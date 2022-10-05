import json

path = 'gameconfigs.json'

with open(path, 'r') as f:
    data = json.loads(f.read())

print(data)
boss = data['boss']
print(boss)

class bigBoss():
    pass
