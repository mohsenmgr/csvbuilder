class Populator:

    _instance = None

   
    def __new__(mflr, *args, **kwargs):
        if mflr._instance is None:
            mflr._instance = super().__new__(mflr)
        return mflr._instance
    

    def __init__(self,inputFile):
        self.measures = []
        self.inputFile = inputFile
        self.populate_measures()

    def getMeasures(self):
        return self.measures
    

    def populate_measures(self):
        with open(self.inputFile, 'r') as file:
            self.measures = [line.strip() for line in file]
    
