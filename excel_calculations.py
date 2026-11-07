import json
import re
from collections import defaultdict
from decimal import Decimal, localcontext
from html import escape

from fastapi import HTTPException


# Accept plain decimal numbers with:
# - Up to 12 digits before the decimal point.
# - Up to 6 digits after the decimal point.
#
# Currency symbols, commas, and scientific notation are rejected
# rather than interpreted ambiguously.
NUMBER_PATTERN = re.compile(
    r"[+-]?\d{1,12}(?:\.\d{1,6})?",
    flags=re.ASCII,
)

TOTAL_LABELS = {
    "total",
    "subtotal",
    "grand total",
    "grand_total",
}


def calculation_error(filename, worksheet, message):
    raise HTTPException(
        status_code=422,
        detail=f"{filename} — worksheet '{worksheet}': {message}",
    )


def parse_number(
    value,
    filename,
    worksheet,
    row_number,
    column,
):
    text = str(value).strip()

    if not NUMBER_PATTERN.fullmatch(text):
        calculation_error(
            filename,
            worksheet,
            f"Row {row_number}, column '{column}' has an invalid "
            f"or missing number: {text!r}. Use a plain number "
            "without currency symbols or thousands separators, "
            "with at most 12 integer digits and 6 decimal places.",
        )

    return Decimal(text)


def format_number(value):
    text = format(value, "f")

    if "." in text:
        text = text.rstrip("0").rstrip(".")

    if text in {"", "-0"}:
        return "0"

    return text


def format_usd(value):
    # Show at least two decimal places.
    # Preserve extra precision instead of silently rounding.
    text = format_number(value)
    whole, separator, fraction = text.partition(".")

    if not separator:
        return whole + ".00"

    return whole + "." + fraction.ljust(2, "0")


def markdown_cell(value):
    text = escape(str(value), quote=False)

    return (
        text.replace("\\", "\\\\")
        .replace("|", "\\|")
        .replace("\r", " ")
        .replace("\n", " ")
    )


def render_records(headers, records, use_tables):
    if use_tables:
        lines = [
            "| " + " | ".join(map(markdown_cell, headers)) + " |",
            "| " + " | ".join("---" for _ in headers) + " |",
        ]

        for record in records:
            lines.append(
                "| "
                + " | ".join(map(markdown_cell, record))
                + " |"
            )

        return "\n".join(lines)

    lines = []

    for index, record in enumerate(records, start=1):
        details = "; ".join(
            f"{markdown_cell(header)}: {markdown_cell(value)}"
            for header, value in zip(headers, record)
        )

        lines.append(f"{index}. {details}")

    return "\n".join(lines)


def calculate_sheet(filename, section):
    worksheet = section["worksheet"]

    # backend.py stores each worksheet as:
    # Line 1: worksheet title
    # Line 2: explanation of cell order
    # Remaining lines: JSON objects containing excel_row and cells
    try:
        rows = [
            json.loads(line)
            for line in section["text"].splitlines()[2:]
            if line.strip()
        ]

        for row in rows:
            if (
                not isinstance(row, dict)
                or not isinstance(row.get("excel_row"), int)
                or not isinstance(row.get("cells"), list)
            ):
                raise ValueError("Unexpected stored row format.")

    except (ValueError, TypeError, KeyError, AttributeError):
        calculation_error(
            filename,
            worksheet,
            "Could not read stored worksheet rows.",
        )

    if len(rows) < 2:
        calculation_error(
            filename,
            worksheet,
            "A header row and at least one data row are required.",
        )

    # Headers are matched without case sensitivity.
    headers = [
        str(value).strip().casefold()
        for value in rows[0]["cells"]
    ]

    nonblank_headers = [
        header for header in headers if header
    ]

    if len(nonblank_headers) != len(set(nonblank_headers)):
        calculation_error(
            filename,
            worksheet,
            "Duplicate column headers were found.",
        )

    columns = {
        header: index
        for index, header in enumerate(headers)
        if header
    }

    missing = {"category", "unit_price_usd"} - columns.keys()

    if missing:
        calculation_error(
            filename,
            worksheet,
            "Missing required headers: "
            + ", ".join(sorted(missing))
            + ". Headers must be in the first nonempty row.",
        )

    quantity_columns = [
        name
        for name in ("quantity", "quantity_sold")
        if name in columns
    ]

    if len(quantity_columns) != 1:
        calculation_error(
            filename,
            worksheet,
            "Include exactly one of Quantity or Quantity_Sold "
            "so the calculation is unambiguous.",
        )

    quantity_column = quantity_columns[0]
    sales_mode = quantity_column == "quantity_sold"

    def cell(row, column):
        index = columns.get(column)

        if index is None or index >= len(row["cells"]):
            return ""

        value = row["cells"][index]
        return "" if value is None else str(value).strip()

    items = []

    groups = defaultdict(
        lambda: {
            "quantity": Decimal("0"),
            "value": Decimal("0"),
        }
    )

    total_quantity = Decimal("0")
    total_value = Decimal("0")

    # Sufficient precision for the accepted numeric values and
    # the workbook size limits in backend.py.
    with localcontext() as context:
        context.prec = 60

        for row in rows[1:]:
            row_number = row["excel_row"]
            category = cell(row, "category")
            item = cell(row, "item") or f"Row {row_number}"

            if not category:
                calculation_error(
                    filename,
                    worksheet,
                    f"Row {row_number} has a blank Category.",
                )

            if (
                category.casefold() in TOTAL_LABELS
                or item.casefold() in TOTAL_LABELS
            ):
                calculation_error(
                    filename,
                    worksheet,
                    f"Row {row_number} appears to be a total or "
                    "subtotal row. Use detail rows only to prevent "
                    "double-counting.",
                )

            quantity = parse_number(
                cell(row, quantity_column),
                filename,
                worksheet,
                row_number,
                quantity_column,
            )

            price = parse_number(
                cell(row, "unit_price_usd"),
                filename,
                worksheet,
                row_number,
                "Unit_Price_USD",
            )

            row_value = quantity * price

            items.append({
                "excel_row": row_number,
                "item": item,
                "category": category,
                "quantity": quantity,
                "price": price,
                "value": row_value,
            })

            groups[category]["quantity"] += quantity
            groups[category]["value"] += row_value

            total_quantity += quantity
            total_value += row_value

        # Confirm that category totals agree with row totals.
        category_quantity = sum(
            (group["quantity"] for group in groups.values()),
            Decimal("0"),
        )

        category_value = sum(
            (group["value"] for group in groups.values()),
            Decimal("0"),
        )

        if (
            category_quantity != total_quantity
            or category_value != total_value
        ):
            raise HTTPException(
                status_code=500,
                detail="Internal calculation verification failed.",
            )

    return {
        "filename": filename,
        "worksheet": worksheet,
        "sales_mode": sales_mode,
        "items": items,
        "groups": dict(groups),
        "total_quantity": total_quantity,
        "total_value": total_value,
    }


def build_excel_summary(selected_documents, use_tables=True):
    summaries = []

    # Calculate every stored worksheet in every selected workbook.
    # An unsupported worksheet causes an error rather than being
    # silently omitted from the answer.
    for _, document in selected_documents:
        if document["file_type"] != "xlsx":
            raise HTTPException(
                status_code=400,
                detail="Excel totals mode accepts only .xlsx files.",
            )

        for section in document["sections"]:
            summaries.append(
                calculate_sheet(
                    document["filename"],
                    section,
                )
            )

    if not summaries:
        raise HTTPException(
            status_code=422,
            detail="No readable worksheets were available.",
        )

    output = [
        "## Excel calculation results",
        "Calculated directly in Python from stored worksheet values. "
        "Totals are kept separate for each worksheet.",
    ]

    for summary in summaries:
        filename = summary["filename"]
        worksheet = summary["worksheet"]

        quantity_label = (
            "Quantity sold"
            if summary["sales_mode"]
            else "Quantity"
        )

        value_label = (
            "Sales value (USD)"
            if summary["sales_mode"]
            else "Stock value (USD)"
        )

        output.append(
            f"### {markdown_cell(filename)} — "
            f"worksheet: {markdown_cell(worksheet)}"
        )

        output.append(
            "Calculation for each data row: "
            "**quantity × Unit_Price_USD**."
        )

        item_records = [
            [
                item["excel_row"],
                item["item"],
                item["category"],
                format_number(item["quantity"]),
                format_usd(item["price"]),
                format_usd(item["value"]),
            ]
            for item in summary["items"]
        ]

        output.append("#### Row calculations")

        output.append(
            render_records(
                [
                    "Excel row",
                    "Item",
                    "Category",
                    quantity_label,
                    "Unit price (USD)",
                    value_label,
                ],
                item_records,
                use_tables,
            )
        )

        category_records = [
            [
                category,
                format_number(group["quantity"]),
                format_usd(group["value"]),
            ]
            for category, group in sorted(
                summary["groups"].items()
            )
        ]

        category_records.append([
            "TOTAL",
            format_number(summary["total_quantity"]),
            format_usd(summary["total_value"]),
        ])

        output.append("#### Category totals")

        output.append(
            render_records(
                [
                    "Category",
                    quantity_label,
                    value_label,
                ],
                category_records,
                use_tables,
            )
        )

        output.append(
            "Verified: category quantities sum to "
            f"**{format_number(summary['total_quantity'])}**; "
            "category values sum to "
            f"**USD {format_usd(summary['total_value'])}**."
        )

        output.append(
            f"Included **{len(summary['items'])} data rows**, "
            f"from Excel row {summary['items'][0]['excel_row']} "
            f"through row {summary['items'][-1]['excel_row']}."
        )

    output.append(
        "**Scope:** These are totals for the parsed detail rows, "
        "not a financial audit. Use flat worksheets with one header "
        "row and no embedded totals, notes, or repeated headers. "
        "The total-row check catches only a few common labels; "
        "it cannot identify every subtotal. "
        "Quantity totals assume compatible units within each worksheet. "
        "Category names are grouped exactly as written after trimming "
        "surrounding spaces. Hidden and filtered rows are included. "
        "Formula values are the saved results in the workbook; "
        "Python does not recalculate Excel formulas. Recalculate "
        "and save the workbook in Excel before uploading if it "
        "contains formulas. Question-box filters are not applied."
    )

    return "\n\n".join(output)