class bossMan():
    magicHP = 100
    bluntHP = 100
    piercingHP = 100

def bossTotalHP():
    bossTotalHP = magicHP + bluntHP + piercingHP
    return bossTotalHP

def bossHPCheck():
    return print('Boss HP: \n', 'Magic: ', magicHP, '\n', 'Blunt: ', bluntHP, '\n', 'Piercing: ', piercingHP)

def bossAttack():
    if bossTotalHP() > 0:
        soldierHours -= 1
    return bossMan



