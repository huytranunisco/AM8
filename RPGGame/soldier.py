class Soldier:
    sTypes = ['Pierce', 'Magic', 'Blunt']
    count = 0

    def __init__(self, turn, dmgMod, type, hasWep = False) -> None:
        self.turnCount = turn
        self.dmgMod = dmgMod

        while type not in self.sTypes:
            print("Not a valid type.")
            type = input('Set type (Pierce, Magic, Blunt): ')

        self.type = type

        if hasWep:
            self.weapon = Weapon(type, 5, 16)
        else: None
        
        Soldier.count += 1
    
    def __str__(self) -> str:
        soldierString = f"Soldier #{self.count}"
        attributeString = "\nType: " + self.type + "\nAmount of Turns: " + str(self.turnCount) + "\nDamage Modifier: x" + str(self.dmgMod)
        if self.weapon == None:
            weaponString = "\nNo Weapon"
        else:
            weaponString = self.weapon.__str__()
        
        return soldierString + attributeString + weaponString
    


class Weapon(object):

    def __init__(self, type, wDmg, wCost) -> None:
        self.type = type
        self.dmg = wDmg
        self.cost = wCost

    def __str__(self) -> str:
        return "\nDamage: " + str(self.dmg) + "\nCost: " + str(self.cost) + "\n"

s = Soldier(2, 1, 'Pierce', True)
print(s)

d = Soldier(2, 1, 'Pierce', True)
print(d)