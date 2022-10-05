import json

class bossMan():
    pass

path = 'RPGGame\gameconfigs.json'

with open(path, 'r') as f:
    data = json.loads(f.read())

bossStats = data['boss']

bmhp = bossStats['boss_magic_hp']
bphp = bossStats['boss_piercing_hp']
bbhp = bossStats['boss_blunt_hp']

def bossTotalHP():
    bossTotalHP = bmhp + bbhp + bphp
    return bossTotalHP

def bossHPCheck():
    return print('Boss HP: \n-----------------\n', 'Magic: ', bmhp, '\n', 'Blunt: ', bbhp, '\n', 'Piercing: ', bphp, '\n-----------------')




