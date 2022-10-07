class BigBoss:
    def __init__(self, pierceHp, magicHp, bluntHp) -> None:
        self.pHp = pierceHp
        self.mHp = magicHp
        self.bHp = bluntHp
        self.totalHp = self.pHp + self.mHp + self.bHp

    def getHPs(self):
        stringOutput = 'Total HP = ' + str(self.totalHp) + '\nPierce HP = ' + str(self.pHp) + '\nMagic HP = ' + str(self.mHp) + '\nBlunt HP = ' + str(self.bHp)
        return stringOutput
    
    def takeDamage(self, type, dmg):
        if type == 'Pierce':
            self.pHp -= dmg
            if self.pHp < 0:
                self.pHp = 0

        if (type == 'Magic' or type == 'Blunt') and self.mHp > 0:
            self.mHp -= dmg
            if self.mHp < 0:
                self.mHp = 0

        if type == 'Blunt':
            self.bHp -= dmg
            if self.bHp < 0:
                self.bHp = 0
        
        self.setTotalHp()
    
    def setTotalHp(self):
        self.totalHp = self.pHp + self.mHp + self.bHp