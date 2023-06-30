from docx import Document
from docx.shared import Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.enum.text import WD_BREAK

LOGO_FILE_PATH = "example_logo.jpg"
OUTPUT_FILE_PATH = "output.docx"


SELLER_DATA = "Sherlock Holmes \n221B Baker Street\nLondon\nNIP : 111111111111\nBank: Gringotts Wizarding Bank\nAccount Number : 11 2222 3333 4444 5555 7777 9999"

ITEM_DATA = "Example description of selling item (2111/21233)"

INVOICE_YEAR = 2077

COMMON_FONT = "Calibri"

CURRENCY = "GBP"

ITEMS_COL_NAMES = ["#", "Name", "Value"]

TABLE_FONT_SIZE_BIG = 8.5

TABLE_FONT_SIZE_SMALL = 8.0

INVOICE_NUMBER_FONT_SIZE = 12.0

SELLER_BUYER_FONT_SIZE = 10.0
