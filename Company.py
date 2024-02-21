class Company:
    name = ""
    job = ""
    question = ""

    def getName(self):
        return str(self.name)
    
    def setName(self, name):
        self.name = name

    def getJob(self):
        return self.job
    
    def setJob(self, job):
        self.job = job

    def setQuestion(self, question):
        self.question = question

    def getQuestion(self):
        return self.question