import json
import bossMan
import soldierMan

boss = bossMan
sold = soldierMan

def main():


    boss.bossHPCheck() 
    print('Total HP: ', boss.bossTotalHP())

    soldierMan.getSoldierCount()
    soldierMan.getSoldierHours()
    soldierMan.getAllSoldierWeapons()


main()