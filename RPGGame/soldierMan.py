class soldierMan():
    pass

soldierWeapons = {'magic': 'staff', 
                  'blunt': 'mace', 
                  'piercing': 'sword'}

def soldierAttack():
    tempDamage = 0
    if soldierCount and soldierHours > 0:
        for x in range(soldierCount):
            if allSoldierWeapons[x] == 'staff':
                attackMagic = int(soldierWeapons['magic'])
                if boss.mHP <= 0:
                    print('Soldier', x + 1, 'tried to attack the boss but its magic HP is already depleted! \n')
                    continue          
                else:
                    boss.mHP -= attackMagic 
                    tempDamage += attackMagic
                    print('Soldier', x + 1, 'attacked the boss for', attackMagic, 'magic HP! \n')
                x += 1
            elif allSoldierWeapons[x] == 'mace':
                attackBlunt = int(soldierWeapons['blunt'])
                if boss.bHP <= 0:
                    print('Soldier', x + 1, 'tried to attack the boss but its blunt HP is already depleted! \n')
                    continue
                else:
                    boss.bHP -= attackBlunt
                    tempDamage += attackBlunt
                    print('Soldier', x + 1, 'attacked the boss for', attackBlunt, 'blunt HP! \n')
                x += 1
            elif allSoldierWeapons[x] == 'sword':
                attackPiercing = int(soldierWeapons['piercing'])
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


