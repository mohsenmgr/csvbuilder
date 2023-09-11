import os

class MyUtils:
    _instance = None

    def __new__(mu, *args, **kwargs):
        if mu._instance is None:
            mu._instance = super().__new__(mu)
        return mu._instance
    

    def __init__(self, filesDirectory):
        self.mainFilesDirectory = filesDirectory


    def getMainFilesDirectory(self):
        return self.mainFilesDirectory
    
    def listFilesInDirectory(self):
        files = [f for f in os.listdir(self.mainFilesDirectory) if os.path.isfile(os.path.join(self.mainFilesDirectory,f))] 
        return files
    
    def listCsvFilesInDirectory(self):
        file_list = []
        for filename in os.listdir(self.mainFilesDirectory):
            if filename.endswith(".csv"):
                file_list.append(filename)

        return file_list
    
    def deleteAllFilesInDirectory(self):
        files = self.listCsvFilesInDirectory()
        idx = 0
        count = len(files)
        for file in files:
            if file.endswith(".csv"):
                os.remove(self.mainFilesDirectory + file)
                idx = idx + 1

        if count == idx:
            return 0
        return 1
    
    def deleteOneFileInDirectory(self,fileName):
        fileName = fileName + ".7z"
        files = self.listFilesInDirectory()
        if fileName in files:
            try:
               os.remove(self.mainFilesDirectory + fileName)
            except Exception as e:
               error_message = e.args[0] if e.args else str(e)
               print(f"Error: {error_message}")
               raise Exception(f"The File Deletion failed Error: {error_message}")

    def checkIfFileExistsInPath(self,filePath):
        if os.path.exists(filePath):
            return 1
        return 0