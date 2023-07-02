from loader import BankTransfers, BankTransfer
from invoice_doc import *


document = Document()

set_margins(document)

asd = BankTransfers()
asd.load_excel()

for bt in asd.bank_transfers:
    bt.print()
    document.add_picture(LOGO_FILE_PATH, width=Inches(1.0))
    add_date(document, bt.accounting_date)
    add_seller_buyer(document, bt.counterparty_data)
    add_invoice_number(document, bt.invoice_number)
    add_items_table(document, bt.amount)
    add_balance(document, bt.amount)
    document.add_page_break()


document.save(OUTPUT_FILE_PATH)
