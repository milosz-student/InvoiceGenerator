from loader_cfg import *


class BankTransfer:
    def __init__(
        self,
        invoice_number,
        accounting_date,
        transaction_type,
        amount,
        currency,
        counterparty_data,
        counterparty_account_number,
        operation_title,
    ):
        self.invoice_number = invoice_number
        self.accounting_date = accounting_date
        self.transaction_type = transaction_type
        self.amount = amount
        self.currency = currency
        self.counterparty_data = counterparty_data
        self.counterparty_account_number = counterparty_account_number
        self.operation_title = operation_title

    def print(self):
        print(
            f"{self.invoice_number} {self.accounting_date} {self.transaction_type} {self.amount} {self.currency} {self.counterparty_data} {self.counterparty_account_number} {self.operation_title}"
        )


class BankTransfers:
    def __init__(self):
        self.bank_transfers = []

    def add(self, bank_transfer):
        self.bank_transfers.append(bank_transfer)

    def print_all(self):
        for bt in self.bank_transfers:
            bt.print()

    def load_excel(self):
        excel_file = pd.read_excel(EXCEL_FILE_PATH)

        for value in excel_file.values:
            invoice_number = value[COLUMNS.index("Nr Invoice")]
            accounting_date = value[COLUMNS.index("Accounting Date")]
            transaction_type = value[COLUMNS.index("Transaction Type")]
            amount = value[COLUMNS.index("Amount")]
            currency = value[COLUMNS.index("Currency")]
            counterparty_data = value[COLUMNS.index("Counterparty Data")]
            counterparty_account_number = value[
                COLUMNS.index("Counterparty Account Number")
            ]
            operation_title = value[COLUMNS.index("Operation Title")]

            self.add(
                BankTransfer(
                    invoice_number,
                    accounting_date,
                    transaction_type,
                    amount,
                    currency,
                    counterparty_data,
                    counterparty_account_number,
                    operation_title,
                )
            )
