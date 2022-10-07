from asyncio.windows_events import NULL
from random import randint


class bossHP:
    magicHP = 100
    bluntHP = 100
    piercingHP = 100

soldierWeapons = {'magic': 'staff', 
                  'blunt': 'mace', 
                  'piercing': 'sword'}
def bossTotalHP():
    bossTotalHP = bossHP.magicHP + bossHP.bluntHP + bossHP.piercingHP
    return bossTotalHP

def bossHPCheck():
    return print('Boss HP: \n', 'Magic: ', bossHP.magicHP, '\n', 'Blunt: ', bossHP.bluntHP, '\n', 'Piercing: ', bossHP.piercingHP)

def bossAttack():
    if bossHP > 0:
        soldierHours -= 1
    return bossHP

def soldierAttack():
    tempDamage = 0
    if soldierCount and soldierHours > 0:
        for x in range(soldierCount):
            if allSoldierWeapons[x] == 'staff':
                attackMagic = int(soldierWeapons['magic'])
                if bossHP.magicHP <= 0:
                    print('Soldier', x + 1, 'tried to attack the boss but its magic HP is already depleted! \n')
                    continue          
                else:
                    bossHP.magicHP -= attackMagic 
                    tempDamage += attackMagic
                    print('Soldier', x + 1, 'attacked the boss for', attackMagic, 'magic HP! \n')
                x += 1
            elif allSoldierWeapons[x] == 'mace':
                attackBlunt = int(soldierWeapons['blunt'])
                if bossHP.bluntHP <= 0:
                    print('Soldier', x + 1, 'tried to attack the boss but its blunt HP is already depleted! \n')
                    continue
                else:
                    bossHP.bluntHP -= attackBlunt
                    tempDamage += attackBlunt
                    print('Soldier', x + 1, 'attacked the boss for', attackBlunt, 'blunt HP! \n')
                x += 1
            elif allSoldierWeapons[x] == 'sword':
                attackPiercing = int(soldierWeapons['piercing'])
                if bossHP.piercingHP <= 0:
                    print('Soldier', x + 1, 'tried to attack the boss but its piercing HP is already depleted! \n')
                    continue
                else:
                    bossHP.piercingHP -= attackPiercing
                    tempDamage += attackPiercing
                    print('Soldier', x + 1, 'attacked the boss for', attackPiercing, 'piercing HP! \n')
                x += 1
            elif allSoldierWeapons[x] == NULL:
                tempDamage += 0
    if bossTotalHP() == 0:
        print('Congrats you defeated the boss!!!\n'
              'Congrats you defeated the boss!!!\n'
              'Congrats you defeated the boss!!!\n')
    else:
        print('No soldiers available to attack')
    totalSoldierDamage = tempDamage
    return totalSoldierDamage

def main():
    global soldierCount
    global soldierHours
    global allSoldierWeapons
    global soldierNum
    soldierNum = 1
    allSoldierWeapons = []

    bossHPCheck()

    soldierCount = int(input('\nHow many soldiers: '))
    soldierHours = int(input('How many hours do the soldiers work: '))
    
    for x in soldierWeapons.keys():
        weaponDamage = input('How much damage does ' + soldierWeapons[x] + ' do? ' )
        soldierWeapons[x] = int(weaponDamage) + randint(-4, 4)
        if soldierWeapons[x] <= 0:
            soldierWeapons[x] = 0
        elif soldierWeapons[x] >= 100:
            soldierWeapons[x] = 100

    for soldierNum in range(soldierCount):
        allSoldierWeapons.append(input('What weapon does soldier ' + str(soldierNum + 1) + ' have (staff, mace, or sword)? '))
        soldierNum += 1

    soldierAttack()

    bossHPCheck()
    
main()
