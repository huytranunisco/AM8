import json

path = 'test.json'

with open('test.json', 'r') as f:
    data = json.loads(f.read())

print(data)