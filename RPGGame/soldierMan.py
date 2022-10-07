import json
from random import randint
import bossMan
from asyncio.windows_events import NULL


allSoldierWeapons = []
boss = bossMan

path = 'RPGGame\gameconfigs.json'
with open(path, 'r') as f:
    data = json.loads(f.read())

weapons = data['weapons']
soldier = data['soldiers']


def getSoldierCount():
    soldierCount = int(input('\nHow many soldiers: '))
    while soldier.items():
        soldier['count'] = soldierCount
        break
    return soldierCount

def getSoldierHours():
    soldierHours = int(input('How many hours do the soldiers work: '))
    while soldier.items():
        soldier['hours'] = soldierHours
        break
    return soldierHours

def getAllSoldierWeapons():
    for soldierNum in range(soldier['count']):
        allSoldierWeapons.append(input('What weapon does soldier ' + str(soldierNum + 1) + ' have (staff, mace, or sword)? '))
        soldierNum += 1

def getWeaponDamage():
    for x in weapons.keys():
        weaponDamage = input('How much damage does ' + x + ' do? ' )
        x = int(weaponDamage) + randint(-4, 4)
        if x <= 0:
            x = 0
        elif x >= 100:
            x = 100
    return weaponDamage

def soldierAttack():
    tempDamage = 0
    if soldier['count'] and soldier['hours'] > 0:
        for x in range(soldier['count']):
            if allSoldierWeapons[x] == 'staff':
                attackMagic = int(weapons['magic_dmg'])
                if boss.bmhp <= 0:
                    print('Soldier', x + 1, 'tried to attack the boss but its magic HP is already depleted! \n')
                    continue
                elif boss.bmhp < attackMagic:
                    print('Soldier', x + 1, 'depleted the bosses magic HP! \n')
                    boss.bmhp = 0
                    continue
                else:
                    boss.bmhp -= attackMagic
                    tempDamage += attackMagic
                    print('Soldier', x + 1, 'attacked the boss for', attackMagic, 'magic HP! \n')
                x += 1
            elif allSoldierWeapons[x] == 'mace':
                attackBlunt = int(weapons['blunt_dmg'])
                if boss.bbhp <= 0:
                    print('Soldier', x + 1, 'tried to attack the boss but its blunt HP is already depleted! \n')
                    continue
                elif boss.bbhp < attackBlunt:
                    print('Soldier', x + 1, 'depleted the bosses blunt HP! \n')
                    boss.bbhp = 0
                    continue
                else:
                    boss.bbhp -= attackBlunt
                    tempDamage += attackBlunt
                    print('Soldier', x + 1, 'attacked the boss for', attackBlunt, 'blunt HP! \n')
                x += 1
            elif allSoldierWeapons[x] == 'sword':
                attackPiercing = int(weapons['piercing_dmg'])
                if boss.bphp <= 0:
                    print('Soldier', x + 1, 'tried to attack the boss but its piercing HP is already depleted! \n')
                    continue
                elif boss.bphp < attackPiercing:
                    print('Soldier', x + 1, 'depleted the bosses piercing HP! \n')
                    boss.bphp = 0
                    continue
                else:
                    boss.bphp -= attackPiercing
                    tempDamage += attackPiercing
                    print('Soldier', x + 1, 'attacked the boss for', attackPiercing, 'piercing HP! \n')
                x += 1
            elif allSoldierWeapons[x] == NULL:
                tempDamage += 0
    if boss.bossTotalHP() == 0:
        print('Congrats you defeated the boss!!!\n'
              'Congrats you defeated the boss!!!\n'
              'Congrats you defeated the boss!!!\n')
    totalSoldierDamage = tempDamage
    
    return totalSoldierDamage


