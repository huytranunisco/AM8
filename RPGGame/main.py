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
    soldiersList = [[],[],[]]

    #Calculates how many soldiers are needed to take down each of the boss's health.
    #For each turn, a soldier is created to attack the boss.
    #If the boss's total health is depleted to 0 or less, the loop exits.
    print("Attacking boss with soldiers...")
    while bossMan.totalHp > 0:
        if bossMan.pHp > 0:
            s = soldier.Soldier(1, 1, 'Pierce', soldierCost, pDmg, pCost)
            bossMan.takeDamage(s.type, s.attack())
            soldiersList[0].append(s)
        if bossMan.mHp > 0:
            s = soldier.Soldier(1, 1, 'Magic', soldierCost, mDmg, mCost)
            bossMan.takeDamage(s.type, s.attack())
            soldiersList[1].append(s)
        if bossMan.bHp > 0:
            s = soldier.Soldier(1, 1, 'Blunt', soldierCost, bDmg, bCost)
            bossMan.takeDamage(s.type, s.attack())
            soldiersList[2].append(s)
        
        print(f'Total HP: {bossMan.totalHp}')
        print(f'Pierce HP: {bossMan.pHp}')
        print(f'Magic HP: {bossMan.mHp}')
        print(f'Blunt HP: {bossMan.bHp}')

    totalSoldierCost = len(soldiersList[0]) + len(soldiersList[1]) + len(soldiersList[2])
    totalWeaponCost = (len(soldiersList[0]) * pCost) + (len(soldiersList[1]) * mCost) + (len(soldiersList[2]) * bCost)

    menuInput = 1
    while menuInput != 0:
        menu()
        menuInput = input("Enter an option (1-6): ")
        while menuInput < 0 or menuInput > 6:
            print("Not a valid option.")
            menuInput = input("Enter an option (1-6): ")
        
        if menuInput == '1':
            sArray = np.array(soldiersList).T
            print(sArray)