# InvoiceGenerator
InvoiceGenerator is a tool used for big generation of invoices based on an .xls file containing invoice data in each row and saving them in one docx file for easy print.

## Study Case
I created this generator because I needed to generate over 120 invoices based on information contained in an Excel file. Below is an example of the data:
![Excel Screenshot](readme_img/excel.png)
The goal was to generate an original page in a docx file for each row.
Example page below:
![Doc Screenshot](readme_img/doc.png)

## Usage
The main library used in this project is **python-docx**. Unfortunately, I encountered some issues when using it with Python 3.9, but it works fine with **Python 3.8**.

Here is a list of libraries I have used:
 - *et-xmlfile	1.1.0	1.1.0*
- *lxml	4.9.2	4.9.2*
- *numpy	1.24.4	1.25.0*
- *openpyxl	3.1.2	3.1.2*
- *pandas	2.0.3	2.0.3*
- *pip	23.1.2	23.1.2*
- *python-dateutil	2.8.2	2.8.2*
- *python-docx	0.8.11	0.8.11*
- *python-docx-2023	0.2.17	0.2.17*
- *pytz	2023.3	2023.3*
- *setuptools	68.0.0	68.0.0*
- *six	1.16.0	1.16.0*
- *tzdata	2023.3	2023.3*
- *xlrd	2.0.1	2.0.1 *

and a helpful link to the python-docx documentation:
https://python-docx.readthedocs.io/en/latest/index.html

Here is simple main.py with elements that you could see in example image above
```python 
document = Document()  #creating empty document

set_margins(document)

asd = BankTransfers() 
asd.load_excel() #creating class containg every Bank Transfer from excel file

for bt in asd.bank_transfers:
    bt.print()
    document.add_picture(LOGO_FILE_PATH, width=Inches(1.0)) # logo is at left-top corner
    add_date(document) # date is at right-top corner
    add_seller_buyer(document, bt.counterparty_data) # this data is in the middle
    add_invoice_number(document, bt.invoice_number) # center
    add_items_table(document, bt.amount) # table 
    add_balance(document, bt.amount) # bottom line balance info
    document.add_page_break() # go to next page

document.save(OUTPUT_FILE_PATH)
```
