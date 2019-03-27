from wordpress import API
import json
import yaml
import pandas as pd

config = "config.yml"
# old config
config = "config_old.yml"

with open(config, 'r') as f:
  try:
    config = yaml.load(f)
  except yaml.YAMLError as exc:
    print(exc)
        
wcapi=API(
        url=config['wordpress_url'],
        consumer_key=config['woo_key'],
        consumer_secret=config['woo_secret'],
        wp_api=True,
        version="wc/v3",
        basic_auth=True,
        user_auth=True,
        timeout=60
)

with open('static_data.yml', 'r') as f:
  try:
    static_data = yaml.load(f)
  except yaml.YAMLError as exc:
    print(exc)

def product_weight(id):
    return static_data['products'][id]['weight']

def product_code(id):
    return static_data['products'][id]['code']

def woo_orders(wcapi):
    return wcapi.get("orders").json()


required_order_cols = [
    "number",
    "status",
    "date_created",
    "billing",
    "shipping",
    "line_items",
    "meta_data",
    "customer_note"
]

required_order_sub_cols = {
    "billing": [
        'first_name',
        'last_name',
        'company',
        'address_1',
        'address_2',
        'city',
        'state',
        'postcode',
        'country',
        'email',
        'phone'
    ],
    "shipping": [
        'first_name',
        'last_name',
        'company',
        'address_1',
        'address_2',
        'city',
        'state',
        'postcode',
        'country'
    ],
    "meta_data": []
}

required_item_cols = [
    'name',
    'product_id',
    'quantity'
]

def parse_order(order):
    parsed_order = {}
    for col in required_order_cols:
        if col in required_order_sub_cols:
            for sub_col in required_order_sub_cols[col]:
                parsed_order[col + "_" + sub_col] = order[col][sub_col]
        else:
            parsed_order[col] = order[col]
    return parsed_order

def extract_order_items(order):
    order_items = []
    for item in order["line_items"]:
        item_order = dict(order)
        del item_order["line_items"]
        for col in required_item_cols:
            if col in item:
                item_order['product_' + col] = item[col]
        item_order['product_requirement'] = product_weight(item['product_id'])*item['quantity']/1000.0
        item_order['product_code'] = product_code(item['product_id'])
        order_items.append(item_order)
    return order_items

def process_orders(orders):
    all_order_items = []
    for order in orders:
        if len(order["line_items"]) > 0:
            parsed_order = parse_order(order)
            all_order_items.extend(extract_order_items(parsed_order))
    return all_order_items

def orders_to_excel:
    orders = process_orders(woo_orders(wcapi))
    df = pd.DataFrame(orders)
    writer = pd.ExcelWriter('orders.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1')
    writer.save()