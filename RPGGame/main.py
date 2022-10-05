import json

path = 'RPGGame\gameconfigs.json'

with open(path, 'r') as f:
    data = json.loads(f.read())

bossStats = data['boss']
print(bossStats)