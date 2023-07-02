from invoice_doc_cfg import *

# Function to set the widths of the table columns
def set_col_widths(table):
    widths = (Inches(0.5), Inches(5), Inches(1.0))
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width


# Function to set the heights of the table rows
def set_row_height(table):
    heights = [Inches(0.2), Inches(0.4), Inches(0.2)]
    for i in range(len(table.rows)):
        table.rows[i].height = heights[i]


# Function to set the margins of the document
def set_margins(document):
    sections = document.sections
    for section in sections:
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)


# Function to add date information to the document
def add_date(document, payment_day):
    date_info = ["Date of Issue: ", "Payment Date: ", "Payment method: "]
    date_value = ["1998-02-21", payment_day.strftime('%Y-%m-%d'), "Bank transfer"]

    date = document.add_paragraph()
    for i in range(len(date_info)):
        date.add_run(date_info[i])
        date.add_run(date_value[i] + "\n").bold = True

    for run in date.runs:
        set_cell_font(run, COMMON_FONT, TABLE_FONT_SIZE_BIG)

    date.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT


# Function to add seller and buyer information to the document
def add_seller_buyer(document, buyer):
    table = document.add_table(rows=1, cols=2)

    seller_cell = table.cell(0, 0)
    buyer_cell = table.cell(0, 1)

    paragraph1 = seller_cell.add_paragraph()
    paragraph1.add_run("Seller:\n").bold = True
    set_cell_font(paragraph1.runs[0], COMMON_FONT, SELLER_BUYER_FONT_SIZE)

    paragraph1.add_run(SELLER_DATA)
    set_cell_font(paragraph1.runs[1], COMMON_FONT, TABLE_FONT_SIZE_BIG)

    paragraph2 = buyer_cell.add_paragraph()
    paragraph2.add_run("Buyer:\n").bold = True

    set_cell_font(paragraph2.runs[0], COMMON_FONT, SELLER_BUYER_FONT_SIZE)

    paragraph2.add_run(buyer).add_break()
    set_cell_font(paragraph2.runs[1], COMMON_FONT, TABLE_FONT_SIZE_BIG)


# Function to add invoice number to the document
def add_invoice_number(document, invoice_nr):
    faktura_inf = f"Invoice {invoice_nr}/{INVOICE_YEAR}"
    faktura = document.add_paragraph()
    faktura.add_run(faktura_inf).bold = True
    set_cell_font(faktura.runs[0], COMMON_FONT, INVOICE_NUMBER_FONT_SIZE)
    faktura.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


# Function to set the font of a cell run
def set_cell_font(cell_run, font_name, font_size):
    cell_run.font.name = font_name
    cell_run.font.size = Pt(font_size)


# Function to add the items table to the document
def add_items_table(document, amount):
    table = document.add_table(rows=3, cols=3)
    table.alignment = 1
    table.style = "Table Grid"
    set_col_widths(table)
    set_row_height(table)

    for i in range(0, 3):
        table.cell(0, i).paragraphs[0].add_run(ITEMS_COL_NAMES[i]).bold = True
        set_cell_font(
            table.cell(0, i).paragraphs[0].runs[0], COMMON_FONT, TABLE_FONT_SIZE_BIG
        )

        table.cell(0, i).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    table.cell(1, 0).paragraphs[0].add_run("1")
    set_cell_font(
        table.cell(1, 0).paragraphs[0].runs[0], COMMON_FONT, TABLE_FONT_SIZE_SMALL
    )
    table.cell(1, 0).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    table.cell(1, 1).paragraphs[0].add_run(
        ITEM_DATA if amount >= 0 else "Return " + ITEM_DATA
    )
    set_cell_font(
        table.cell(1, 1).paragraphs[0].runs[0], COMMON_FONT, TABLE_FONT_SIZE_SMALL
    )

    table.cell(2, 1).paragraphs[0].add_run("All").bold = True
    set_cell_font(
        table.cell(2, 1).paragraphs[0].runs[0], COMMON_FONT, TABLE_FONT_SIZE_SMALL
    )
    table.cell(2, 1).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    for i in range(1, 3):
        table.cell(i, 2).paragraphs[0].add_run(f"{amount}.00{CURRENCY}")
        set_cell_font(
            table.cell(i, 2).paragraphs[0].runs[0], COMMON_FONT, TABLE_FONT_SIZE_SMALL
        )
        table.cell(i, 2).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER


# Function to add balance to the document
def add_balance(document, amount):
    p = document.add_paragraph()

    p.add_run().add_break()

    table = document.add_table(rows=1, cols=3)

    table.cell(0, 0).paragraphs[0].add_run(f"All: {amount} {CURRENCY}")
    table.cell(0, 1).paragraphs[0].add_run(f"Paid: {amount} {CURRENCY}")
    table.cell(0, 2).paragraphs[0].add_run(f"Owing: 0,00{CURRENCY}").bold = True

    for i in range(0, 3):
        set_cell_font(
            table.cell(0, i).paragraphs[0].runs[0], COMMON_FONT, TABLE_FONT_SIZE_SMALL
        )
