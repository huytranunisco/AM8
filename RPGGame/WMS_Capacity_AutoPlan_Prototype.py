class bossHP:
        magicHP = 100
        bluntHP = 100
        piercingHP = 100

class soldierWeapons:
        magic = 'book'
        blunt = 'bat'
        piercing = 'spear'
class weaponDamage:
        bookDMG = 10
        batDMG = 25
        spearDMG = 35

def bossHPCheck():
    return print('Boss HP: \n', 'Magic: ', bossHP.magicHP, '\n', 'Blunt: ', bossHP.bluntHP, '\n', 'Piercing: ', bossHP.piercingHP)

def bossAttack():
    if bossHP > 0:
        soldierHours -= 1
    return bossHP

def soldierAttack():
    if soldierCount and soldierHours > 0:
        for x in range(soldierCount):
            if allSoldierWeapons[x] == 'bat':
                attack1 = weaponDamage.bookDMG
                bossHP.magicHP -= attack1
                totalSoldierDamage += attack1
                x += 1
            elif allSoldierWeapons[x] == 'book':
                attack2 = weaponDamage.batDMG
                bossHP.bluntHP -= attack2 
                totalSoldierDamage += attack2
                x += 1
            elif allSoldierWeapons[x] == 'spear':
                attack3 = weaponDamage.spearDMG
                bossHP.piercingHP -= attack3
                totalSoldierDamage += attack3
                x += 1
            else:
                totalSoldierDamage = 0
        return totalSoldierDamage

def main():
    global soldierCount
    global soldierHours
    global soldierWeapon
    global allSoldierWeapons
    global soldierNum
    soldierNum = 1
    allSoldierWeapons = []

    bossHPCheck()

    soldierCount = int(input('\nHow many soldiers: '))
    soldierHours = int(input('How many hours do the soldiers work: '))

    for soldierNum in range(soldierCount):
        allSoldierWeapons.append(input('What weapon does soldier ' + str(soldierNum) + ' have (bat, book, or spear)? '))
        soldierNum += 1

    print(allSoldierWeapons)
    soldierAttack()

    bossHPCheck()

main()


    
