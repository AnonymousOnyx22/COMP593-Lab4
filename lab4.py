"""
comp 593 - lab 4: business process automation
reads sales_data.csv and generates individual excel order files.
"""

import os
import sys
import pandas as pd
from datetime import date


def get_csv_path():
    """grab the csv path from the command line and make sure it actually exists"""
    if len(sys.argv) < 2:
        print("error: no csv file path provided.")
        print("usage: python lab4.py <path_to_sales_data.csv>")
        sys.exit(1)

    csv_path = sys.argv[1]
    if not os.path.isfile(csv_path):
        print(f"error: couldn't find a file at '{csv_path}'")
        sys.exit(1)

    return csv_path


def create_orders_dir(csv_path):
    """create the Orders_YYYY-MM-DD folder next to wherever the csv lives"""
    csv_dir = os.path.dirname(os.path.abspath(csv_path))
    dir_name = f"Orders_{date.today().isoformat()}"
    orders_dir = os.path.join(csv_dir, dir_name)
    os.makedirs(orders_dir, exist_ok=True)
    return orders_dir


def load_sales_data(csv_path):
    """read the csv into a dataframe and strip any sneaky BOM characters from headers"""
    df = pd.read_csv(csv_path)
    df.columns = df.columns.str.lstrip('\ufeff')
    return df


def get_order_ids(df):
    """pull out the unique order IDs and sort them"""
    return sorted(df['ORDER ID'].unique())


def build_order_df(df, order_id):
    """
    filter down to one order, calculate the total price per line,
    sort by item number, and trim to just the columns we need
    """
    order = df[df['ORDER ID'] == order_id].copy()

    # multiply quantity by price to get the line total
    order['TOTAL PRICE'] = order['ITEM QUANTITY'] * order['ITEM PRICE']

    # smallest item number first
    order = order.sort_values('ITEM NUMBER').reset_index(drop=True)

    # only keep the columns the spec asks for
    cols = [
        'ORDER DATE', 'ITEM NUMBER', 'PRODUCT LINE', 'PRODUCT CODE',
        'ITEM QUANTITY', 'ITEM PRICE', 'TOTAL PRICE', 'STATUS', 'CUSTOMER NAME'
    ]
    return order[cols]


def save_order_excel(order_df, order_id, orders_dir):
    """write the order out to a nicely formatted xlsx file"""
    filename = f"Order_{order_id}.xlsx"
    filepath = os.path.join(orders_dir, filename)

    grand_total = order_df['TOTAL PRICE'].sum()

    with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
        order_df.to_excel(writer, index=False, sheet_name='Order')

        workbook = writer.book
        worksheet = writer.sheets['Order']

        # two formats we'll reuse: money and bold
        money_fmt = workbook.add_format({'num_format': '$#,##0.00'})
        bold_fmt  = workbook.add_format({'bold': True})

        # item price (F) and total price (G) both get the money format
        num_rows = len(order_df)
        worksheet.set_column('F:F', 13, money_fmt)
        worksheet.set_column('G:G', 13, money_fmt)

        # set each column width to match the spec
        worksheet.set_column('A:A', 11)  # order date
        worksheet.set_column('B:B', 13)  # item number
        worksheet.set_column('C:C', 15)  # product line
        worksheet.set_column('D:D', 15)  # product code
        worksheet.set_column('E:E', 15)  # item quantity
        worksheet.set_column('H:H', 10)  # status
        worksheet.set_column('I:I', 30)  # customer name

        # grand total goes one row below the last data row
        grand_total_row = num_rows + 1
        worksheet.write(grand_total_row, 4, 'GRAND TOTAL:', bold_fmt)
        worksheet.write(grand_total_row, 6, grand_total, money_fmt)

    print(f"  created: {filename}")


def main():
    csv_path = get_csv_path()
    orders_dir = create_orders_dir(csv_path)
    df = load_sales_data(csv_path)
    order_ids = get_order_ids(df)

    print(f"generating {len(order_ids)} order files in: {orders_dir}")

    for order_id in order_ids:
        order_df = build_order_df(df, order_id)
        save_order_excel(order_df, order_id, orders_dir)

    print("done.")


if __name__ == '__main__':
    main()