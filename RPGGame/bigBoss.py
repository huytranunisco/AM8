import json

with open('RPGGame\test.json', 'r', encoding='utf-8') as file:
    data = json.loads(file.read())

print(data)
width = data['width']
print(width)

class bigBoss():
    pass
