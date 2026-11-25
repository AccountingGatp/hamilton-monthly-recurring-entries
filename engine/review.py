"""Read-only processing for the recurring schedule and ResMan exports."""
import calendar
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

from openpyxl import load_workbook

CENT = Decimal("0.01")
OFFSET_ACCOUNTS = {"62010": "15010", "63010": "29030"}
PROPERTY_LABELS = {
    "pinn rd 2226": "PR2226", "oval meadows": "OM7902", "putters cv": "PC4807",
    "alsacetrail": "AT5712", "windsoroaks/7575": "WO7575",
}


def clean(value):
    return str(value).strip() if value is not None else ""


def account(value):
    match = re.match(r"^(?:GL\s*-?\s*)?(\d{5})(?:\b|$)", clean(value), re.I)
    return match[1] if match else ""


def money(value):
    raw = clean(value).replace(",", "").replace("$", "")
    if not raw:
        return None
    if raw.startswith("(") and raw.endswith(")"):
        raw = "-" + raw[1:-1]
    try:
        amount = Decimal(raw)
        return amount.quantize(CENT, rounding=ROUND_HALF_UP) if amount.is_finite() else None
    except InvalidOperation:
        return None


def as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for pattern in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y", "%b %d, %Y", "%B %d, %Y"):
        try:
            return datetime.strptime(clean(value), pattern).date()
        except ValueError:
            pass
    return None


def month_start(day):
    return day.replace(day=1)


def month_end(day):
    return day.replace(day=calendar.monthrange(day.year, day.month)[1])


def preceding(day):
    return date(day.year - (day.month == 1), day.month - 1 or 12, 1)


def month_name(value):
    name = value.lower().rstrip(".")
    lookup = {calendar.month_name[m].lower(): m for m in range(1, 13)}
    lookup.update({calendar.month_abbr[m].lower(): m for m in range(1, 13)})
    lookup["sept"] = 9
    return lookup.get(name)


def tab_month(title):
    value = title.strip().replace("’", "'")
    numeric = re.fullmatch(r"(\d{4})[-_ /.](\d{1,2})", value)
    if numeric:
        year, month = map(int, numeric.groups())
    else:
        named = re.fullmatch(r"([a-z]+)\.?[\s_'/-]*(\d{2}|\d{4})", value, re.I)
        if not named:
            return None
        name, number = named.groups()
        month = month_name(name)
        year = int(number) + (2000 if len(number) == 2 else 0)
    return date(year, month, 1) if year and month and 1 <= month <= 12 else None


def normalized(value):
    return " ".join(re.sub(r"[^\w]+", " ", clean(value).casefold()).split())


def get(row, column):
    return row[column] if column is not None and column < len(row) else None


def headings(row):
    return {normalized(value): index for index, value in enumerate(row) if clean(value)}


def column(columns, *names):
    return next((columns[name] for name in names if name in columns), None)


def workbook_sheets(path):
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        return book.sheetnames
    finally:
        book.close()


def read_chart(path):
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        candidates = []
        for sheet in book:
            rows = sheet.values
            for row in rows:
                columns = headings(row)
                number = column(columns, "account number", "account code", "gl code", "gl account code")
                name = column(columns, "account", "gl account", "account name", "gl account name")
                if name is None or column(columns, "type", "account type", "gl account type") is None:
                    continue
                summary = column(columns, "summary", "is summary", "summary account")
                accounts = {}
                for values in rows:
                    name_value = clean(get(values, name))
                    code = account(get(values, number)) or account(name_value)
                    if code and normalized(get(values, summary)) not in {"yes", "true", "1"}:
                        label = name_value if account(name_value) else f"{code} - {name_value}".rstrip(" -")
                        if code in accounts and accounts[code] != label:
                            raise ValueError(f"Conflicting GL account labels for {code} in {sheet.title}.")
                        accounts[code] = label
                if accounts:
                    candidates.append(accounts)
                break
        if len(candidates) != 1:
            raise ValueError("The GL Summary needs one account list with Account and Type columns.")
        return candidates[0]
    finally:
        book.close()


@dataclass(frozen=True)
class Line:
    row: int
    property: str
    account: str
    debit: Decimal
    credit: Decimal
    memo: str

    @property
    def net(self):
        return self.debit - self.credit


@dataclass
class Entry:
    id: str
    date: date = None
    memo: str = ""
    basis: str = ""
    book: str = ""
    lines: list = field(default_factory=list)


REGISTER_DATE = r"(?:\d{1,2}/\d{1,2}/\d{2,4}|\d{4}-\d{2}-\d{2})"
REGISTER_PERIOD = re.compile(
    rf"({REGISTER_DATE})\s*(?:-|–|—|to)\s*({REGISTER_DATE})\s*[-–—]\s*"
    r"Accrual\s*[-–—]\s*Accounting Book\s*:\s*(.+)", re.I,
)


def cell_after(row, position):
    return next((value for value in row[position + 1:] if clean(value)), None)


def read_register(path):
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        candidates = []
        for sheet in book:
            header = list(sheet.iter_rows(max_row=30, values_only=True))
            text = [clean(value) for row in header for value in row if clean(value)]
            if not any(normalized(value) == "journal entry register" for value in text):
                continue
            period = next((match for value in text if (match := REGISTER_PERIOD.search(value))), None)
            if period:
                candidates.append((sheet, text, period))
        if len(candidates) != 1:
            raise ValueError("Provide one accrual-basis Journal Entry Register with its date range and accounting book.")
        sheet, header, period = candidates[0]
        start, end = (as_date(value) for value in period.groups()[:2])
        book_name = period[3].strip()
        printed_match = next((match for value in header if
                              (match := re.match(rf"Printed\s*:?\s*({REGISTER_DATE})", value, re.I))), None)
        printed = as_date(printed_match[1]) if printed_match else None
        preamble = header[:header.index(period.string)]
        all_scope = {"all", "all properties", "properties all", "property all"}
        scope = "ALL" if any(normalized(value) in all_scope for value in preamble) else "Unverified"

        entries = []
        entry = None
        line_columns = None
        metadata_column = None
        labels = {"date": "date", "description": "memo", "accounting basis": "basis", "accounting book": "book"}
        for row_number, row in enumerate(sheet.values, 1):
            fields = headings(row)
            id_column = column(fields, "id number", "journal id", "journal entry id")
            user_column = column(fields, "user")
            if id_column is not None and user_column is not None:
                entry = Entry(clean(cell_after(row, id_column)))
                entries.append(entry)
                metadata_column = user_column
                line_columns = None
                continue
            if entry is None:
                continue
            label = normalized(get(row, metadata_column))
            if line_columns is None and label in labels:
                value = cell_after(row, metadata_column)
                setattr(entry, labels[label], as_date(value) if label == "date" else clean(value))
                continue

            property_column = column(fields, "property", "property code", "property name")
            account_column = column(fields, "gl account", "account", "account number", "gl code")
            debit_column = column(fields, "debit", "debits")
            credit_column = column(fields, "credit", "credits")
            if None not in (property_column, account_column, debit_column, credit_column):
                line_columns = (property_column, account_column, debit_column, credit_column,
                                column(fields, "description", "memo", "line description"))
                continue
            if line_columns is None:
                continue
            prop, raw_account, raw_debit, raw_credit, memo = (get(row, value) for value in line_columns)
            code = account(raw_account)
            if not clean(prop) and not clean(raw_account):
                if clean(raw_debit) or clean(raw_credit):
                    line_columns = None
                continue
            if clean(prop).startswith("©") and not clean(raw_account):
                continue
            if not code or not clean(prop):
                raise ValueError(f"Missing property or GL account on register row {row_number}.")
            debit, credit = money(raw_debit), money(raw_credit)
            if (clean(raw_debit) and debit is None) or (clean(raw_credit) and credit is None):
                raise ValueError(f"Invalid debit or credit on register row {row_number}.")
            entry.lines.append(Line(row_number, clean(prop).upper(), code,
                                    debit or Decimal(0), credit or Decimal(0), clean(memo)))
        if not start or not end or end < start or any(
            not e.date or not e.id or not e.lines or not e.basis or not e.book
            or not start <= e.date <= end for e in entries
        ) or len({e.id for e in entries}) != len(entries):
            raise ValueError("The register contains incomplete, duplicate or out-of-period journal entries.")
        return entries, start, end, printed, book_name, scope
    finally:
        book.close()


@dataclass
class Expected:
    row: int
    property: str
    account: str
    date: date
    memo: str
    note: str
    amount: Decimal
    county: bool


class ScheduleMonthMissing(ValueError):
    pass


def schedule_columns(row):
    fields = headings(row)
    code = column(fields, "gl", "gl account", "gl code", "account", "account code", "account number")
    amount = column(fields, "amount", "expected amount", "monthly amount", "recurring amount")
    if code is None or amount is None:
        return None
    return (code, column(fields, "date", "posting date", "expected date"),
            column(fields, "description", "memo"), amount,
            column(fields, "notes", "note", "comments", "support"),
            column(fields, "property", "property code", "property name"))


def schedule_rows(sheet):
    prop = ""
    county = False
    default_columns = (0, 1, 2, 3, 4, None)
    columns = default_columns
    for i, row in enumerate(sheet.values, 1):
        heading = clean(get(row, 0))
        key = normalized(heading)
        if key.startswith("accrued county tax disbursement"):
            county = True
            prop = ""
            columns = (0, 1, 2, 5, None, None)
            continue
        if key.startswith("insurance details for"):
            prop = ""
            columns = default_columns
            continue
        header = schedule_columns(row)
        if header:
            columns = header
            continue

        first, when, memo, amount, note, property_value = (get(row, value) for value in columns)
        code = account(first)
        if not code:
            if (not county and heading and not heading.startswith("*")
                and all(not clean(get(row, c)) for c in (1, 2, 3))
                and re.fullmatch(r"[\w][\w ./&-]{1,100}", heading)
                and key not in {"gl", "date", "property", "description", "notes", "amount"}):
                prop = heading
            continue
        memo = clean(memo)
        label = clean(property_value) or prop
        if county or not label:
            label = re.split(r"\s*-\s*", memo)[-1].strip()
            if county and label == memo:
                match = re.search(r"(?:WO|OC|PC|OM|AT|BPT)\d{4}|AlsaceTrail|WindsorOaks/7575", memo, re.I)
                label = match[0] if match else ""
        value = money(amount)
        if value == 0:
            continue
        yield Expected(i, label.upper(), "63010" if county and code == "29030" else code,
                       as_date(when), memo, "" if county else clean(note), value, county)


def read_schedule(path, month, tab_name=None):
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        if tab_name:
            if tab_name not in book.sheetnames:
                raise ValueError(f"Schedule tab {tab_name!r} was not found.")
            sheet = book[tab_name]
            dated = tab_month(sheet.title)
            month_only = month_name(sheet.title.strip())
            if (dated and dated != month) or (month_only and month_only != month.month):
                raise ValueError(f"Schedule tab {tab_name!r} does not match the chosen schedule month.")
            selected = [(sheet, list(schedule_rows(sheet)))]
        else:
            selected = [(sheet, list(schedule_rows(sheet))) for sheet in book if tab_month(sheet.title) == month]
            if not selected:
                for sheet in book:
                    if month_name(sheet.title.strip()) != month.month:
                        continue
                    rows = list(schedule_rows(sheet))
                    years = {row.date.year for row in rows if row.date and row.date.month == month.month}
                    if years == {month.year}:
                        selected.append((sheet, rows))
        if not selected:
            raise ScheduleMonthMissing("No schedule tab could be identified for this month. "
                                       "Select the schedule tab manually if its name is undated or custom.")
        if len(selected) > 1:
            raise ValueError("More than one schedule tab matches this month. Select the correct tab.")
        sheet, expected = selected[0]
        if not expected:
            raise ValueError("No recurring schedule rows were found on the selected tab.")
        return sheet.title, expected
    finally:
        book.close()


def property_key(value):
    return re.sub(r"[^\w]", "", clean(value).casefold()).replace("_", "")


def run(schedule, chart, register, target, allow_rollforward, tab_name=None, tab_is_prior=False):
    prior = preceding(target)
    carried = tab_is_prior
    if tab_name:
        tab, expected = read_schedule(schedule, prior if carried else target, tab_name)
    else:
        try:
            tab, expected = read_schedule(schedule, target)
        except ScheduleMonthMissing:
            if not allow_rollforward:
                raise
            tab, expected = read_schedule(schedule, prior)
            carried = True

    books = read_chart(chart)
    entries, start, end, printed, book_name, properties_scope = read_register(register)
    covered = (properties_scope.upper() == "ALL" and start <= target
               and end >= month_end(target) and printed and printed > month_end(target))
    index = defaultdict(list)
    for entry in entries:
        if entry.basis.lower() not in {"accrual", "both"} or entry.book.lower() != book_name.lower():
            continue
        for line in entry.lines:
            index[month_start(entry.date), line.property, line.account].append((entry, line))

    known_properties = {line.property for e in entries for line in e.lines}
    normalized_properties = defaultdict(set)
    for prop in known_properties:
        normalized_properties[property_key(prop)].add(prop)
    for row in expected:
        if row.property in known_properties:
            continue
        options = normalized_properties[property_key(row.property)]
        if not options:
            alias = PROPERTY_LABELS.get(row.property.lower())
            if row.property == "JMB":
                alias = "JMBGROUP"
            options = {alias} if alias in known_properties else set()
        if len(options) == 1:
            row.property = next(iter(options))

    counts = Counter((row.property, row.account) for row in expected)
    amount_counts = Counter((row.property, row.account, row.amount) for row in expected)
    report = []
    drafts = []
    referenced_rows = set()
    for row in expected:
        key = (row.property, row.account)
        offset = OFFSET_ACCOUNTS.get(row.account)
        current = index[target, *key]
        previous = index[prior, *key]
        proposed_date = row.date
        valid_source_date = row.date and month_start(row.date) == (prior if carried else target)
        if carried and valid_source_date:
            last_current = month_end(target).day
            new_day = last_current if row.date == month_end(prior) else min(row.date.day, last_current)
            proposed_date = target.replace(day=new_day)
        if not valid_source_date:
            proposed_date = None

        def supported(entry, line):
            peers = [other for other in entry.lines if other.row != line.row]
            return (offset and len(peers) == 1 and peers[0].property == row.property
                    and peers[0].account == offset and peers[0].credit == row.amount
                    and not peers[0].debit and line.debit == row.amount and not line.credit)

        match = [(e, l) for e, l in current if l.net == row.amount]
        good_prior = [(e, l) for e, l in previous if supported(e, l)]
        if row.amount is None:
            status = "Review schedule amount"
        elif row.property not in known_properties:
            status = "Verify property code or register scope"
        elif row.account not in books:
            status = "Verify account in chart"
        elif amount_counts[*key, row.amount] > 1 or len(match) > 1:
            status = "Review possible duplicate"
        elif match:
            entry, line = match[0]
            referenced_rows.add(line.row)
            if sum(other.net for other in entry.lines) != 0:
                status = "Amount found - verify journal balance"
            elif offset and not supported(entry, line):
                status = "Amount found - verify offset or split"
            elif proposed_date and entry.date != proposed_date:
                status = "Amount found - verify posting date"
            elif not proposed_date:
                status = "Amount found - verify schedule date"
            elif len(current) > counts[key]:
                status = "Amount found - review additional activity"
            else:
                status = "Matched journal entry"
        elif current:
            status = "Review current journal activity"
        elif not covered:
            status = "Incomplete register coverage - verify"
        elif not proposed_date:
            status = "Verify schedule date"
        elif not offset or offset not in books or row.amount <= 0:
            status = "Review treatment or supporting document"
        elif counts[key] > 1 or len(good_prior) != 1 or len(previous) != 1:
            status = "Verify prior journal and supporting document"
        else:
            status = "Draft available - verify source support"
            entry, prior_line = good_prior[0]
            memo = f"{calendar.month_name[target.month]} {target.year} - {books[row.account]}"
            for code, debit, credit in ((row.account, row.amount, ""), (offset, "", row.amount)):
                drafts.append([f"{target:%Y%m}-{row.row}", proposed_date, row.property, code,
                               books[code], memo, debit, credit, entry.basis, entry.book, entry.id, tab, row.row])

        report.append([status, row.property, row.account, books.get(row.account, ""), row.amount,
                       proposed_date, row.date, row.memo, row.note, tab, row.row,
                       len(current), ", ".join(dict.fromkeys(e.id for e, _ in current)),
                       ", ".join(dict.fromkeys(e.id for e, _ in previous)),
                       sum((line.net for _, line in current), Decimal(0)),
                       sum((line.net for _, line in previous), Decimal(0))])

    watched = {row.account for row in expected}
    activity = [[e.id, e.date, l.property, l.account, l.debit, l.credit, l.memo or e.memo,
                 l.row, "Schedule amount found" if l.row in referenced_rows else "Review against schedule"]
                for (period, _, code), values in index.items() if period == target and code in watched
                for e, l in values]
    details = {
        "Review month": f"{target:%Y-%m}", "Schedule tab": tab,
        "Schedule source": "Prior month carried forward" if carried else "Selected month",
        "Register coverage": f"{start} to {end}, printed {printed or 'unknown'}, scope {properties_scope}",
        "Full review month in register": "Yes" if covered else "No",
        "Register includes prior month": "Yes" if start <= prior and end >= month_end(prior) else "No",
        "Journal entries read": len(entries),
        "Schedule rows reviewed": len(report),
        "Draft entries": len(drafts) // 2,
        "Scope": "Journal entries only. Bills, invoice activity, trailing P&L and loan statements are outside these exports.",
        "Mortgage": "Verify against the current loan statement and invoice or GL Detail export.",
        "Drafts": "Review source support and check for posting since the register was printed. Draft CSV is a workpaper.",
        "Sources": " | ".join(Path(getattr(p, "name", p)).name for p in (schedule, chart, register)),
    }
    return details, report, drafts, activity
