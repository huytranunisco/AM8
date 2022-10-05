import json
import bossMan

path = 'RPGGame\gameconfigs.json'

with open(path, 'r') as f:
    data = json.loads(f.read())

weapons = data['weapons']
soldier = data['soldiers']
allSoldierWeapons = []
boss = bossMan

def getSoldierCount():
    soldierCount = int(input('\nHow many soldiers: '))

    with open(path, 'a') as a:
        json.dump()

    for val in soldier['count']:
        if val == 0:
            soldier['coount'].replace(soldierCount)

def getSoldierHours():
    soldierHours = int(input('How many hours do the soldiers work: '))
    return soldierHours

def getAllSoldierWeapons():
    for soldierNum in range(soldierCount()):
        allSoldierWeapons.append(input('What weapon does soldier ' + str(soldierNum + 1) + ' have (staff, mace, or sword)? '))
        soldierNum += 1

def soldierAttack():
    tempDamage = 0
    if getSoldierCount.soldierCount and getSoldierHours.soldierHours > 0:
        for x in range(getSoldierHours.soldierHours):
            if allSoldierWeapons[x] == 'staff':
                attackMagic = int(weapons['magic_dmg'])
                if boss.mHP <= 0:
                    print('Soldier', x + 1, 'tried to attack the boss but its magic HP is already depleted! \n')
                    continue          
                else:
                    boss.mHP -= attackMagic 
                    tempDamage += attackMagic
                    print('Soldier', x + 1, 'attacked the boss for', attackMagic, 'magic HP! \n')
                x += 1
            elif allSoldierWeapons[x] == 'mace':
                attackBlunt = int(weapons['blunt_dmg'])
                if boss.bHP <= 0:
                    print('Soldier', x + 1, 'tried to attack the boss but its blunt HP is already depleted! \n')
                    continue
                else:
                    boss.bHP -= attackBlunt
                    tempDamage += attackBlunt
                    print('Soldier', x + 1, 'attacked the boss for', attackBlunt, 'blunt HP! \n')
                x += 1
            elif allSoldierWeapons[x] == 'sword':
                attackPiercing = int(weapons['piercing_dmg'])
                if boss.pHP <= 0:
                    print('Soldier', x + 1, 'tried to attack the boss but its piercing HP is already depleted! \n')
                    continue
                else:
                    boss.pHP -= attackPiercing
                    tempDamage += attackPiercing
                    print('Soldier', x + 1, 'attacked the boss for', attackPiercing, 'piercing HP! \n')
                x += 1
            elif allSoldierWeapons[x] == NULL:
                tempDamage += 0
    if boss.bossTotalHP() == 0:
        print('Congrats you defeated the boss!!!\n'
              'Congrats you defeated the boss!!!\n'
              'Congrats you defeated the boss!!!\n')
    else:
        print('No soldiers available to attack')
    totalSoldierDamage = tempDamage
    return totalSoldierDamage


