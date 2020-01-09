from partners import partners
import xlrd, xlsxwriter
import os, logging
from datetime import date

logging.basicConfig(filename='errors.log', filemode='w', format='%(asctime)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S')

PATH = "./summary_files"
FILE_PREFIX = "MCSUMM"
CUST_NAME_COL = 2
CARD_COL = 3
TRANS_COL = 4
TIME_COL = 5
CLEARED_COL = 6
MERCHANT_COL = 7
AMOUNT_COL = 21

FIRST_CUST_ROW = 3

PROCESSING_FEE = .3 # 30 cent processing fee

partner_transactions = {}
date_range = {
    "start": None,
    "end": None
}

# return list of all summary filenames
def get_summary_files():
    files = []
    for filename in os.listdir(PATH):
        if FILE_PREFIX in filename:
            files.append(filename)
    
    return files


def process_file(filename):
    loc = PATH + "/" + filename
    wb = xlrd.open_workbook(loc, logfile=open(os.devnull, 'w'))
    sheet = wb.sheet_by_index(1)
    num_rows = sheet.nrows

    next_cust_row = FIRST_CUST_ROW
    while (next_cust_row < num_rows) and sheet.cell_value(next_cust_row, CUST_NAME_COL):
        next_cust_row = process_customer(sheet, next_cust_row)
    
    get_date(sheet) # update date_range

def process_customer(sheet, cust_row):
    cust_name = sheet.cell_value(cust_row, CUST_NAME_COL).strip()
    card_num = sheet.cell_value(cust_row, CARD_COL)[-4:]

    cust_key = (cust_name, card_num)

    transaction_row = cust_row + 1

    # loop through all transactions for this customer
    while sheet.cell_value(transaction_row, MERCHANT_COL):
        merchant = sheet.cell_value(transaction_row, MERCHANT_COL)

        if merchant in partners:
            spent = sheet.cell_value(transaction_row, AMOUNT_COL) 

            # create a transaction record
            transaction = [
                sheet.cell_value(transaction_row, TRANS_COL),   # transaction date
                int(sheet.cell_value(transaction_row, TIME_COL)),    # transaction time
                sheet.cell_value(transaction_row, CLEARED_COL), # cleared date
                spent   # transaction amount
            ]

            # check if cust has purchased from merchant already
            if cust_key in partner_transactions[merchant]:
                

                # a purchase already exists, update totals
                partner_transactions[merchant][cust_key]["total_spent"] += spent
                partner_transactions[merchant][cust_key]["trans_count"] += 1

            else:
                # add this customer to list for this partner
                partner_transactions[merchant][cust_key] = {
                    "total_spent": spent,
                    "trans_count": 1,
                    "transactions": []
                }

            # append transaction to list for current merchant
            partner_transactions[merchant][cust_key]["transactions"].append(transaction)

            partner_transactions[merchant]["total"] += spent
            partner_transactions[merchant]["count"] += 1

        transaction_row += 1
    
    # return row of next customer
    return transaction_row + 1 

def get_date(sheet):
    date_str = sheet.cell_value(1,2)[-10:]
    date_obj = date(int(date_str[-4:]), int(date_str[:2]), int(date_str[3:5]))

    #check if earliest date
    if not date_range["start"] or date_obj < date_range["start"]:
        date_range["start"] = date_obj
    #check if latest date
    if not date_range["end"] or date_obj > date_range["end"]:
        date_range["end"] = date_obj

    # print(date_range["start"].strftime("%m/%d/%y") + "-" + date_range["end"].strftime("%m/%d/%y"))


def create_report(merchant):
    filename = f"{partners[merchant]['retailer_name']}_summary_{date_range['start'].strftime('%Y%m%d')}_{date_range['end'].strftime('%Y%m%d')}"
    workbook = xlsxwriter.Workbook(filename + '.xlsx')
    ws = workbook.add_worksheet()

    # adds formatting for $ cells
    currency_format = workbook.add_format({'num_format': '[$$-409]#,##0.00'})

    ws.write(0,0, "American River Gold Transaction Summarization and Invoice")
    ws.write(1,0, "Merchant Name:")
    ws.write(1,1, partners[merchant]['retailer_name'])
    ws.write(2,0, "Appears As:")
    ws.write(2,1, merchant)
    ws.write(3,0, "Contact:")
    ws.write(3,1, partners[merchant]["contact_name"])
    ws.write(4,0, "Phone:")
    ws.write(4,1, partners[merchant]["contact_phone"])
    ws.write(5,0, "Contact:")
    ws.write(5,1, partners[merchant]["contact_email"])
    ws.write(6,0, "Period:")
    ws.write(6,1, f"{date_range['start'].strftime('%m/%d/%Y')} - {date_range['end'].strftime('%m/%d/%Y')}")
    
    row = 8
    ws.write(row,0, "Cardholder")
    ws.write(row,1, "Card last 4")
    ws.write(row,2, "Trans Date")
    ws.write(row,3, "Trans Time")
    ws.write(row,4, "Settled Date")
    ws.write(row,5, "Amount")

    row = 9
    for key in partner_transactions[merchant].keys():
        if key != "total" and key != "count":
            row = print_customer(merchant, key, row, ws, currency_format) 

    cust_disc = partners[merchant]['customer_discount']
    arg_fee = partners[merchant]['ARG_fee']
    total = partner_transactions[merchant]['total']
    ws.write(row, 0, "Total Spent:")
    ws.write(row, 1, total)
    ws.write(row+1, 0, f"Total Discount ({round(cust_disc * 100, 2)}%):")
    ws.write(row+1, 1, round(total * cust_disc, 2), currency_format)
    ws.write(row+2, 0, f"ARG fee ({round(arg_fee * 100, 2)}%):")
    ws.write(row+2, 1, round(total * arg_fee, 2), currency_format)
    ws.write(row+3, 0, "Processing fee:")
    ws.write(row+3, 1, .30, currency_format)

    ws.write(row+5, 0, "Total bill:")
    ws.write(row+5, 1, round(total * cust_disc, 2) + round(total * arg_fee, 2) + .3, currency_format)

    # set column widths
    ws.set_column('A:A', 25)
    ws.set_column('B:B', 12)
    ws.set_column('C:C', 12)
    ws.set_column('D:D', 12)
    ws.set_column('E:E', 15)
    ws.set_column('F:F', 12)
    workbook.close()

def print_customer(merchant, cust_key, row, ws, currency_format):
    
    ws.write(row,0, cust_key[0])
    ws.write(row,1, cust_key[1])
    
    transactions = partner_transactions[merchant][cust_key]['transactions']
    cust_disc = partners[merchant]['customer_discount']
    arg_fee = partners[merchant]['ARG_fee']
    total = partner_transactions[merchant][cust_key]['total_spent']

    row += 1
    # list all transactions for given customer
    for t in transactions:
        col = 2
        ws.write(row, col, t[0])
        ws.write(row, col+1, t[1])
        ws.write(row, col+2, t[2])
        ws.write(row, col+3, t[3], currency_format)
        row += 1 

    ws.write(row, col+2, "Total:")
    ws.write(row, col+3, total, currency_format)

    row += 1

    ws.write(row, col+2, f"discount ({cust_disc * 100}%):")
    ws.write(row, col+3, round(total * cust_disc, 2), currency_format)

    return row + 2

if __name__ == "__main__":
    # create empty list of customers for each partner
    for k in partners.keys():
        partner_transactions[k] = {
            "count": 0,
            "total": 0
        }
    
    files = get_summary_files()
    for f in files:
        # process_file(f)
        try:
            process_file(f)
        except:
            logging.error('unable to process file: "%s"', f)

    # create_report("MR PICKLES - 132")
    [create_report(k) for k in partner_transactions.keys()]
    # for k in partner_transactions.keys():
    #     create_report(k)

    print(partner_transactions)