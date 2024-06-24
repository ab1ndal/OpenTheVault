
import numpy as np
import pandas as pd

class CreateCut:
    def __init__(self, cutName='', diagCoord=[], cutDirection='Z', cutStep=0, startCoord=0, endCoord=0, groupName = 'ALL', unit='m', vec1=0, vec2=0):
        self.cutName = cutName
        self.diagCoord = diagCoord
        self.cutDirection = cutDirection
        self.cutStep = cutStep
        self.startCoord = startCoord
        self.endCoord = endCoord
        self.groupName = groupName
        self.unit = unit
        self.vec1 = vec1
        self.vec2 = vec2
        self.quad = pd.DataFrame(columns=['SectionCut', 'X', 'Y', 'Z'])
        self.general = pd.DataFrame(columns=['CutName', 'DefinedBy', 'Group', 'ResultType', 'DefaultLoc', 'GlobalX','GlobalY', 'GlobalZ', 'AngleA', 'AngleB', 'AngleC', 'DesignType', 'DesignAngle', 'ElemSide'])
        self.advAxis = pd.DataFrame(columns=['SectionCut', 'LocalPlane', 'AxOption1', 'AxCoordSys', 'AxCoordDir', 'AxVecJt1', 'AxVecJt2', 'PlOption1', 'PlCoordSys', 'CoordDir1', 'CoordDir2', 'PlVecJt2', 'PlVecJt2', 'AxVecX', 'AxVecY', 'AxVecZ', 'PlVecX', 'PlVecY', 'PlVecZ'])
        self.quad.loc[len(self.quad)] = ['',self.unit, self.unit, self.unit]
        self.general.loc[len(self.general)] = ['','','','','',self.unit, self.unit, self.unit, 'Degrees', 'Degrees', 'Degrees', '', 'Degrees', '']
        self.advAxis.loc[len(self.advAxis)] = 19*['']  
    
    def inputCutName(self, **kwargs):
        if kwargs.get('cutName'):
            self.cutName = kwargs.get('cutName')
        else:
            self.cutName = input('Enter the name of the cut: ')
    
    def inputDiagCoord(self, **kwargs):
        if kwargs.get('diagCoord'):
            self.diagCoord = kwargs.get('diagCoord')
        else:
            self.diagCoord = [float(x) for x in input('Enter the diagonal coordinates of the cut: ').split()]
    
    def inputCutDirection(self, **kwargs):
        if kwargs.get('cutDirection'):
            self.cutDirection = kwargs.get('cutDirection')
        else:
            self.cutDirection = input('Enter the cut direction (X, Y, Z): ')
    
    def inputCutStep(self, **kwargs):
        if kwargs.get('cutStep'):
            self.cutStep = kwargs.get('cutStep')
        else:
            self.cutStep = float(input('Enter the cut step: '))
    
    def inputStartCoord(self, **kwargs):
        if kwargs.get('startCoord'):
            self.startCoord = kwargs.get('startCoord')
        else:
            self.startCoord = float(input('Enter the start coordinate of the normal: '))
    
    def inputEndCoord(self, **kwargs):
        if kwargs.get('endCoord'):
            self.endCoord = kwargs.get('endCoord')
        else:
            self.endCoord = float(input('Enter the end coordinate of the normal: '))
    
    def inputGroupName(self, **kwargs):
        if kwargs.get('groupName'):
            self.groupName = kwargs.get('groupName')
        else:
            self.groupName = input('Enter the group name: ')
    
    def inputUnit(self, **kwargs):
        if kwargs.get('unit'):
            self.unit = kwargs.get('unit')
        else:
            self.unit = input('Enter the unit of the coordinates: ')
    
    def inputVec1(self, **kwargs):
        if kwargs.get('vec1'):
            self.vec1 = kwargs.get('vec1')
        else:
            self.vec1 = input('Enter the first Plane Vector Joint: ')
    
    def inputVec2(self, **kwargs):
        if kwargs.get('vec2'):
            self.vec2 = kwargs.get('vec2')
        else:
            self.vec2 = input('Enter the second Plane Vector Joint: ')

    
    def defineCut(self):
        # create a list of cuts from the start to the end with the cut step include both start and end
        cutList = list(np.arange(self.startCoord, self.endCoord, self.cutStep))
        cutList = [round(x, 3) for x in cutList]
        cutList.append(self.endCoord)
        
        # create a list of cuts in the cut direction
        # The diagonal coordinates are the coordinates of the rectangle diagonal.
        # For each cut, create 4 entries in dataframe with the coordinates of the rectangle vertices
        # The rectangle vertices are defined as follows:
        # For a cut in the Z direction, the vertices are defined as follows: (x1, y1, z1), (x2, y1, z1), (x1, y2, z1), (x1, y1, z1)
        # For a cut in the Y direction, the vertices are defined as follows: (x1, y1, z1), (x1, y1, z2), (x2, y1, z1), (x2, y1, z2)
        # For a cut in the X direction, the vertices are defined as follows: (x1, y1, z1), (x1, y2, z1), (x1, y1, z2), (x1, y2, z2)
        
             

        for c in cutList:
            cutNameInList = f'{self.cutName} - {self.cutDirection}={c}{self.unit}'
            if self.cutDirection == 'Z':
                self.quad.loc[len(self.quad)] = [cutNameInList, self.diagCoord[0], self.diagCoord[1], c]
                self.quad.loc[len(self.quad)] = [cutNameInList, self.diagCoord[0], self.diagCoord[3], c]
                self.quad.loc[len(self.quad)] = [cutNameInList, self.diagCoord[2], self.diagCoord[1], c]
                self.quad.loc[len(self.quad)] = [cutNameInList, self.diagCoord[2], self.diagCoord[3], c]
            elif self.cutDirection == 'Y':
                self.quad.loc[len(self.quad)] = [cutNameInList, self.diagCoord[0], c, self.diagCoord[1]]
                self.quad.loc[len(self.quad)] = [cutNameInList, self.diagCoord[0], c, self.diagCoord[3]]
                self.quad.loc[len(self.quad)] = [cutNameInList, self.diagCoord[2], c, self.diagCoord[1]]
                self.quad.loc[len(self.quad)] = [cutNameInList, self.diagCoord[2], c, self.diagCoord[3]]
            elif self.cutDirection == 'X':
                self.quad.loc[len(self.quad)] = [cutNameInList, c, self.diagCoord[0], self.diagCoord[1]]
                self.quad.loc[len(self.quad)] = [cutNameInList, c, self.diagCoord[0], self.diagCoord[3]]
                self.quad.loc[len(self.quad)] = [cutNameInList, c, self.diagCoord[2], self.diagCoord[1]]
                self.quad.loc[len(self.quad)] = [cutNameInList, c, self.diagCoord[2], self.diagCoord[3]]
            self.general.loc[len(self.general)] = [cutNameInList, 'Quad', self.groupName, 'Analysis', 'Yes',0,0,0,0,0,0,'', '', 'Positive']
            self.advAxis.loc[len(self.advAxis)] = [cutNameInList,'31', 'Coord Dir', 'GLOBAL', 'Z', 'None', 'None', 'Two Joints', 'GLOBAL', 'X', 'Y', self.vec1, self.vec2, 0, 0, 1, 1, 0, 0]
    
    def printExcel(self, fileLoc=''):

        with pd.ExcelWriter(fileLoc + '\\cuts.xlsx') as writer:
            self.general.to_excel(writer, index=False, sheet_name='Section Cuts 1 - General')
            self.advAxis.to_excel(writer, index=False, sheet_name='Section Cuts 2 - Advanced Axes')
            self.quad.to_excel(writer, index=False, sheet_name='Section Cuts 3 - Quadrilaterals')
        # Create an excel file with multiple sheets
        # The first sheet is the section cuts 
        # The second sheet is the group names

