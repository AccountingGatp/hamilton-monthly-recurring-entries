import json
from io import BytesIO
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from exports import review_month, save_result
from review import run, workbook_sheets


def workbook(data, name):
    source = BytesIO(data.to_py())
    source.name = name
    return source


def tabs(data):
    try:
        return json.dumps(workbook_sheets(workbook(data, "schedule.xlsx")))
    except (BadZipFile, InvalidFileException, KeyError) as error:
        raise ValueError("The schedule is not a readable .xlsx workbook.") from error


def review(schedule, chart, register, file_names, month, tab, prior):
    sources = (
        workbook(data, name)
        for data, name in zip((schedule, chart, register), json.loads(file_names))
    )
    try:
        result = run(*sources, review_month(month), prior,
                     tab_name=tab or None, tab_is_prior=bool(tab and prior))
    except (BadZipFile, InvalidFileException, KeyError) as error:
        raise ValueError("One of the selected files is not a readable .xlsx workbook.") from error
    output = BytesIO()
    save_result(output, result)
    details, rows, _, _ = result
    display_rows = [[row[i] for i in (1, 2, 4, 14, 15, 0, 12)] for row in rows]
    return json.dumps({"details": details, "rows": display_rows}, default=str), output.getvalue()
