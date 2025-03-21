from scripts.readFile import connectDB, getData
import pandas as pd


# Define Load Combination
#The load combo should create a dataframe with following columns:
# Case, LoadType, LoadName, Function, LoadSF, TransAccSF, RotAccSF, TimeFactor, ArrivalTime, CoordSys, Angle
# LoadType is "Load Pattern", TransAccSF is empty, RotAccSF is empty, TimeFactor is 1, ArrivalTime is 0, CoordSys is empty, Angle is empty

# The function should take in prefix for Function, ScaleAmp as Nmax at Zmax and Nmin at Zmin
# The function should return a dataframe with the above columns 

class LoadCombo:
    def __init__(self, filePath, Nmax, Nmin, Zmax, Zmin, prefixList):
        self.Nmax = Nmax
        self.Nmin = Nmin
        self.Zmax = Zmax
        self.Zmin = Zmin
        self.prefixList = prefixList
        self.conn = connectDB(filePath)
        # Get the Loas Patterns 
        self.baseX = getXGroup()
        self.baseY = getYGroup()
        self.baseZ = getZGroup()
        #self.df = self.defineLoadCombo()

    def defineJointDisp(self):
        pass
         

    def defineLoadCombo(self):
        loadCombo = getData(self.conn, tableName='Case - Direct History 2 - Load Assignments')
        for prefix in self.prefixList:
            comboName = [prefix + "HorOnly", prefix + "VertOnly", prefix + "HorVert"]

        # Case should be prefix + "HorOnly" or prefix + "VertOnly" or prefix + "HorVert"
        # LoadType is "Load Pattern", TransAccSF is empty, RotAccSF is empty, TimeFactor is 1, ArrivalTime is 0, CoordSys is empty, Angle is empty
        # LoadName is prefix + "HorOnly" or prefix + "VertOnly" or prefix + "HorVert"
        # Function is prefix + "HorOnly" or prefix + "VertOnly" or prefix + "HorVert"
        # LoadSF is Nmax at Zmax and Nmin at Zmin
        # Create a dataframe with the above columns
        return data



class LoadPattern:
    def __init__(self, filePath):
        self.conn = connectDB(filePath)
        self.df = self.defineLoadPattern()


    def addJointGroup(self):
        pass
    

    def defineLoadPattern(self):
        loadPattern = getData(self.conn, tableName='Case - Direct History 2 - Load Assignments')
        # Case should be prefix + "HorOnly" or prefix + "VertOnly" or prefix + "HorVert"
        # LoadType is "Load Pattern", TransAccSF is empty, RotAccSF is empty, TimeFactor is 1, ArrivalTime is 0, CoordSys is empty, Angle is empty
        # LoadName is prefix + "HorOnly" or prefix + "VertOnly" or prefix + "HorVert"
        # Function is prefix + "HorOnly" or prefix + "VertOnly" or prefix + "HorVert"
        # LoadSF is Nmax at Zmax and Nmin at Zmin
        # Create a dataframe with the above columns
        return data

    def getXGroup(self):
        return getData(self.conn, tableName='Groups 2 - Assignments')

    def getYGroup(self):
        return getData(self.conn, tableName='Groups 3 - Assignments')

    def getZGroup(self):
        return getData(self.conn, tableName='Groups 4 - Assignments')
