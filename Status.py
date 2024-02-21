class Status:
    charactor_limit = 0
    strength = []
    experience = {}

    def getCharactorLimit(self):
        return self.charactor_limit
    
    def setCharactorLimit(self, charactor_limit):
        self.charactor_limit = charactor_limit

    def getStrength(self, num):
        return self.strength[num]
    
    def getStrengthNum(self):
        return len(self.strength)
    
    def setStrength(self, strength):
        self.strength.append(strength)

    def getAllStrength(self):
        return self.strength

    def getExperience(self):
        return self.experience
    
    def getAllExperience(self):
        return self.experience
    
    def setExperience(self, experience, learned):
        self.experience[experience] = learned

    def getLearned(self, experience):
        return self.experience[experience]
    
    def getAllLearned(self):
        return self.experience.values()
    
    def getExperienceName(self, learned):
        key = [k for k, v in self.experience.items() if v == learned]
        return key[0]