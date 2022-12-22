import json
import requests
import pandas as pd

url = 'http://127.0.0.1:5000/model'

dict = {'ID': {'2': 3}, 'n1': {'2': 3}, 'n2': {'2':3}}
r = requests.post(url, json=dict)
print(r.status_code)

r = requests.get(url)
print(r.json())