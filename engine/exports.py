import csv
from collections import Counter
from datetime import datetime
from io import StringIO
from zipfile import ZIP_DEFLATED, ZipFile


REVIEW_FIELDS = [
    "Status", "Property", "Account", "Account Name", "Expected Amount", "Expected Date",
    "Schedule Date", "Schedule Description", "Schedule Notes", "Schedule Tab", "Schedule Row",
    "Current Journal Line Count", "Current Journal IDs", "Prior Journal IDs",
    "Current Journal Net", "Prior Journal Net",
]
DRAFT_FIELDS = [
    "Draft ID", "Proposed Date", "Property", "Account", "Account Name", "Description",
    "Debit", "Credit", "Basis", "Book", "Prior Journal ID", "Schedule Tab", "Schedule Row",
]
ACTIVITY_FIELDS = [
    "Journal ID", "Date", "Property", "Account", "Debit", "Credit", "Memo",
    "Register Row", "Schedule Check",
]


def review_month(value):
    parsed = datetime.strptime(value, "%Y-%m").date()
    if value != parsed.strftime("%Y-%m"):
        raise ValueError("Enter the review month as YYYY-MM.")
    return parsed


def excel_safe(value):
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def to_csv(headers, rows):
    output = StringIO(newline="")
    writer = csv.writer(output)
    writer.writerow(headers)
    for row in rows:
        writer.writerow(excel_safe(value) for value in row)
    return output.getvalue().encode("utf-8-sig")


def save_result(path, result):
    details, rows, drafts, activity = result
    summary = list(details.items()) + sorted(Counter(row[0] for row in rows).items())
    exports = {
        "Summary.csv": to_csv(["Check", "Result"], summary),
        "Recurring_Review.csv": to_csv(REVIEW_FIELDS, rows),
        "Journal_Activity.csv": to_csv(ACTIVITY_FIELDS, activity),
    }
    if drafts:
        exports["Draft_Journal_Workpaper.csv"] = to_csv(DRAFT_FIELDS, drafts)
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        for name, contents in exports.items():
            archive.writestr(name, contents)


