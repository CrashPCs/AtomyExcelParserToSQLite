import os
import pandas as pd
import sqlite3

def create_database():
    conn = sqlite3.connect('atomy_orders.db')
    cursor = conn.cursor()
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS orders (
        order_id TEXT PRIMARY KEY,
        customer_name TEXT,
        customer_id INTEGER,
        phone_number TEXT,
        order_date TEXT,
        status TEXT DEFAULT 'Not Arrived'
    )
    ''')
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS items (
        order_id TEXT,
        product_code TEXT,
        product_name TEXT,
        quantity INTEGER,
        FOREIGN KEY(order_id) REFERENCES orders(order_id)
    )
    ''')
    conn.commit()
    conn.close()

def read_process_excel(file_path):
    data = pd.read_excel(file_path, header=None, skiprows=9)
    orders = []
    items = []
    current_order_id = None
    for index, row in data.iterrows():
        if any("Итого" in str(cell) for cell in row):
            continue
        if is_order_row(row[0]):
            current_order_id = row[0]
            order_date = extract_date(row[0])
            customer_id = extract_customer_id(row)
            phone_number = str(row[7]).split('.')[0] if pd.notnull(row[7]) else None
            
            orders.append({
                'order_id': current_order_id,
                'customer_name': row[3],
                'customer_id': customer_id,
                'phone_number': phone_number,
                'order_date': order_date
            })
        elif current_order_id and pd.notnull(row[0]) and not is_order_row(row[0]) and not str(row[0]).startswith('G000R'):
            items.append({
                'order_id': current_order_id,
                'product_code': row[0],
                'product_name': row[3],
                'quantity': int(row[8]) if pd.notnull(row[8]) else 0
            })
    return orders, items

def extract_customer_id(row):
    # Проверяем столбцы 6 и 7 на наличие customer_id
    for col_index in [5, 6]:
        if pd.notnull(row[col_index]) and is_valid_customer_id(row[col_index]):
            return int(row[col_index])
    return None

def is_valid_customer_id(value):
    try:
        int(value)
        return True
    except (ValueError, TypeError):
        return False

def is_order_row(number):
    return isinstance(number, str) and number.startswith('R') and len(number) == 12 and number[1:].isdigit()

def extract_date(order_number):
    year = '20' + order_number[1:3]
    month = order_number[3:5]
    day = order_number[5:7]
    return f"{year}-{month}-{day}"

def insert_data_into_db(orders, items):
    conn = sqlite3.connect('atomy_orders.db')
    cursor = conn.cursor()
    for order in orders:
        cursor.execute('''
        INSERT OR REPLACE INTO orders (order_id, customer_name, customer_id, phone_number, order_date)
        VALUES (?, ?, ?, ?, ?)
        ''', (order['order_id'], order['customer_name'], order['customer_id'], 
              order['phone_number'], order['order_date']))
    for item in items:
        cursor.execute('''
        INSERT INTO items (order_id, product_code, product_name, quantity)
        VALUES (?, ?, ?, ?)
        ''', (item['order_id'], item['product_code'], item['product_name'], item['quantity']))
    conn.commit()
    conn.close()

if __name__ == '__main__':
    directory = r'C:\Users\SystemX\Desktop\Консолидации'  # Specify your directory with Excel files
    create_database()
    for filename in os.listdir(directory):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(directory, filename)
            orders, items = read_process_excel(file_path)
            insert_data_into_db(orders, items)
            print(f"Processed {filename}")
