from wordpress import API
import json
import yaml
import pandas as pd

with open("config.yml", 'r') as f:
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

def woo_orders(wcapi):
    return wcapi.get("orders").json()


    