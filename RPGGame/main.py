import json
import soldier
import bigBoss
import numpy as np
import pandas as pd

#Displays a menu of options for users to see different resources available.
def menu():
    print("__Main Menu__")
    print("1. Display all soldiers")
    print("2. Display a certain type of soldiers")
    print("3. Display cost for all soldiers")
    print("4. Display cost for a certain type of soldiers")
    print("5. Display total cost of weapons")
    print("6. Display cost for a certain type of weapons")
    print("0. Quit")

#Displays a menu of options for users to see certain types of an item/entity.
def typeMenu():
    print('Which type do you want to see?')
    print("1. Pierce")
    print("2. Magic")
    print("3. Blunt")
    print("0. Back")

if __name__ == '__main__':

    #Opens the game's configs file where certain resources are changed.
    with open('RPGGame\gameconfigs.json', 'r') as f:
        data = json.loads(f.read())

    #Retrieves boss's data from the configs file.
    bossStats = data['boss']
    bossPHp = bossStats['boss_p_hp']
    bossMHp = bossStats['boss_m_hp']
    bossBHp = bossStats['boss_b_hp']

    #Retrieves weapons' data from the configs file.
    weaponStats = data['weapon']
    pDmg, pCost = weaponStats['p_dmg'], weaponStats['p_cost']
    mDmg, mCost = weaponStats['m_dmg'], weaponStats['m_cost']
    bDmg, bCost = weaponStats['b_dmg'], weaponStats['b_cost']

    #Retrieves soldier's data from the configs file.
    soldierStats = data['soldier']
    soldierCost = soldierStats['s_cost']

    #Creates a boss entity with parameters of three different types of health.
    bossMan = bigBoss.BigBoss(bossPHp, bossMHp, bossBHp)

    #Creates an empty 2D array for any created soldiers that will be categorized into types.
    soldiersList = []

    #Calculates how many soldiers are needed to take down each of the boss's health.
    #For each turn, a soldier is created to attack the boss.
    #If the boss's total health is depleted to 0 or less, the loop exits.
    print("Attacking boss with soldiers...")

    #
    soldierID = 0

    while bossMan.totalHp > 0:
        if bossMan.pHp > 0:
            s = soldier.Soldier(1, 1, 'Pierce', soldierCost, pDmg, pCost)
            bossMan.takeDamage(s.type, s.attack())
            sValues = [soldierID, s.atkCount, s.dmgMod, s.type, s.cost, s.weapon.dmg, s.weapon.cost]
            soldiersList.append(sValues)
            soldierID += 1
        if bossMan.mHp > 0:
            s = soldier.Soldier(1, 1, 'Magic', soldierCost, mDmg, mCost)
            bossMan.takeDamage(s.type, s.attack())
            sValues = [soldierID, s.atkCount, s.dmgMod, s.type, s.cost, s.weapon.dmg, s.weapon.cost]
            soldiersList.append(sValues)
            soldierID += 1
        if bossMan.bHp > 0:
            s = soldier.Soldier(1, 1, 'Blunt', soldierCost, bDmg, bCost)
            bossMan.takeDamage(s.type, s.attack())
            sValues = [soldierID, s.atkCount, s.dmgMod, s.type, s.cost, s.weapon.dmg, s.weapon.cost]
            soldiersList.append(sValues)
            soldierID += 1
        
    totalSoldierCost = len(soldiersList[0]) + len(soldiersList[1]) + len(soldiersList[2])
    totalWeaponCost = (len(soldiersList[0]) * pCost) + (len(soldiersList[1]) * mCost) + (len(soldiersList[2]) * bCost)

    soldierdf = pd.DataFrame(soldiersList, columns=['ID', 'Attack Count', 'Damage Modifier', 'Type', 'Soldier Cost', 'Weapon Damage', 'Weapon Cost'])

    menuInput = -1
    while menuInput != 0:
        menu()
        try:
            menuInput = int(input("Enter an option (0-6): "))
        except:
            print("Not a valid option.")
            menuInput = int(input("Enter an option (0-6): "))

        while menuInput < 0 or menuInput > 6:
            print("Not a valid option.")
            menuInput = int(input("Enter an option (0-6): "))
        
        if menuInput == 1:
            print(soldierdf)

        elif menuInput == 2:
            typeMenu()
            submenuInput = int(input("Enter an option (1-3): "))
            while submenuInput < 0 or submenuInput > 3:
                print("Not a valid option.")
                submenuInput = int(input("Enter an option (1-3): "))

            if submenuInput == 1:
                piercedf = soldierdf[soldierdf['Type'] == 'Pierce']
                print(piercedf)

            elif submenuInput == 2:
                magicdf = soldierdf[soldierdf['Type'] == 'Magic']
                print(magicdf)
            
            elif submenuInput == 3:
                bluntdf = soldierdf[soldierdf['Type'] == 'Blunt']
                print(bluntdf)

        elif menuInput == 3:
            x = soldierdf['Soldier Cost'].sum()
            print('Total soldier cost: ', x)

        elif menuInput == 4:
            typeMenu()
            submenu2Input = int(input("Enter an option (1-3): "))
            while submenu2Input < 0 or submenu2Input > 3:
                print("Not a valid option.")
                submenu2Input = int(input("Enter an option (1-3): "))

            if submenu2Input == 1:
                sum = soldierdf[soldierdf['Type'] == 'Pierce']['Soldier Cost'].sum()
                print('Total cost for pierce soldier: ', sum)

            elif submenu2Input == 2:
                sum = soldierdf[soldierdf['Type'] == 'Magic']['Soldier Cost'].sum()
                print('Total cost for magic soldier: ', sum)            
            elif submenu2Input == 3:
                sum = soldierdf[soldierdf['Type'] == 'Blunt']['Soldier Cost'].sum()
                print('Total cost for blunt soldier: ', sum)

            elif submenu2Input == 4:
                menuInput = 1

        elif menuInput == 5:
            print("\nTotal weapon cost for all soldiers: ", soldierdf['Weapon Cost'].sum(), "\n")

        elif menuInput == 6:
            typeMenu()
            submenu3Input = int(input("Enter an option (1-3): "))
            while submenu3Input < 0 or submenu3Input > 3:
                print("Not a valid option.")
                submenu3Input = int(input("Enter an option (1-3): "))
            if submenu3Input == 1:
                sum = soldierdf[soldierdf['Type'] == 'Pierce']['Weapon Cost'].sum()
                print('Total weapon cost for pierce soldier: ', sum)

            elif submenu3Input == 2:
                sum = soldierdf[soldierdf['Type'] == 'Magic']['Weapon Cost'].sum()
                print('Total weapon cost for magic soldier: ', sum)
            
            elif submenu3Input == 3:
                sum = soldierdf[soldierdf['Type'] == 'Blunt']['Weapon Cost'].sum()
                print('Total weapon cost for blunt soldier: ', sum)

            elif submenu3Input == 4:
                menuInput = 1
