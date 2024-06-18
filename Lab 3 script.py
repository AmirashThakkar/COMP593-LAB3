import os
import sys
import pandas as pd
from datetime import datetime

def input_file(filepath):
    if not filepath:
        print("Error: No path to sales data CSV file provided.")
        sys.exit(1)
    if not os.path.isfile(filepath):
        print("Error: The provided path does not exist or is not a file.")
        sys.exit(1)

def orders_directory(base_dir):
    date_string = datetime.now().strftime("%Y-%m-%d")
    orders_dir = os.path.join(base_dir, f"Orders_{date_string}")
    if not os.path.exists(orders_dir):
        os.makedirs(orders_dir)
    return orders_dir

def Sales_Data(filepath, orders_dir):
    Sales_Data = pd.read_csv(filepath)

    # Print the columns to check their names
    print("Columns in CSV file:", Sales_Data.columns)

    grouped_orders = Sales_Data.groupby('ORDER ID')

    for order_id, order_information in grouped_orders:
        order_information = order_information.sort_values(by='ITEM NUMBER')
        order_information['TOTAL PRICE'] = order_information['ITEM QUANTITY'] * order_information['ITEM PRICE']

        total_price = order_information['TOTAL PRICE'].sum()

        # Update the column names based on the CSV file
        order_info = order_information[['ORDER ID', 'ITEM NUMBER', 'PRODUCT LINE', 'ITEM QUANTITY', 'ITEM PRICE', 'TOTAL PRICE']]

        order_filepath = os.path.join(orders_dir, f"Order_{order_id}.xlsx")
        with pd.ExcelWriter(order_filepath, engine='xlsxwriter') as writer:
            order_information.to_excel(writer, index=False, sheet_name=f'Order_{order_id}')

            workbook  = writer.book
            worksheet = writer.sheets[f'Order_{order_id}']

            money_format_ = workbook.add_format({'num_format': '$#,##0.00'})
            worksheet.set_column('E:F', 18, money_format_)
            worksheet.set_column('A:D', 15)

            worksheet.write(len(order_info) + 1, 4, 'Grand Total')
            worksheet.write(len(order_info) + 1, 5, total_price, money_format_)
            worksheet.set_column('A:A', 12)
            worksheet.set_column('B:B', 10)
            worksheet.set_column('C:C', 20)
            worksheet.set_column('D:D', 14)

def main():
    if len(sys.argv) < 2:
        print("Error: No path to sales data CSV file provided.")
        sys.exit(1)

    filepath = sys.argv[1]
    input_file(filepath)

    base_directory = os.path.dirname(filepath)
    orders_directory = orders_directory(base_directory)

    Sales_Data(filepath, orders_directory)

    print(f"Excel files have been generated in {orders_directory}")

if __name__ == "__main__":
    main()