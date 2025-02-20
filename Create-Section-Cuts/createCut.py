
import numpy as np
import pandas as pd

class CreateCut:
    def __init__(self, cutName='', diagCoord=[], cutDirection='Z', 
                 cutStep=0, startCoord=0, endCoord=0, groupName = 'ALL', 
                 unit='m', vec1='', vec2='', advAxisExists=False, localPlane='31', 
                 is4Pt=False, **kwargs):
        if 'addSpecialCoord' in kwargs.keys():
            self.addSpecialCoord = kwargs.get('addSpecialCoord')
        else:
            self.addSpecialCoord = []
        if 'rmSpecialCoord' in kwargs.keys():
            self.rmSpecialCoord = kwargs.get('rmSpecialCoord')
        else:
            self.rmSpecialCoord = []
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
        self.advAxisExists = advAxisExists
        self.is4Pt = is4Pt
        self.localPlane = localPlane
        self.quad = pd.DataFrame(columns=['SectionCut', 'X', 'Y', 'Z'])
        self.general = pd.DataFrame(columns=['CutName', 'DefinedBy', 'Group', 'ResultType', 'DefaultLoc', 'GlobalX','GlobalY', 'GlobalZ', 'AngleA', 'AngleB', 'AngleC', 'DesignType', 'DesignAngle', 'ElemSide'])
        self.advAxis = pd.DataFrame(columns=['SectionCut', 'LocalPlane', 'AxOption1', 'AxCoordSys', 'AxCoordDir', 'AxVecJt1', 'AxVecJt2', 'PlOption1', 'PlCoordSys', 'CoordDir1', 'CoordDir2', 'PlVecJt2', 'PlVecJt2', 'AxVecX', 'AxVecY', 'AxVecZ', 'PlVecX', 'PlVecY', 'PlVecZ'])
        self.quad.loc[len(self.quad)] = ['',self.unit, self.unit, self.unit]
        self.general.loc[len(self.general)] = ['','','','','',self.unit, self.unit, self.unit, 'Degrees', 'Degrees', 'Degrees', '', 'Degrees', '']
        self.advAxis.loc[len(self.advAxis)] = 19*['']
        self.fileName = 'cuts.xlsx'
    
    def inputCutName(self, **kwargs):
        if kwargs.get('cutName'):
            self.cutName = kwargs.get('cutName')
        else:
            self.cutName = input('Enter the name of the cut: ')
    
    def inputFileName(self, **kwargs):
        if kwargs.get('fileName'):
            self.fileName = kwargs.get('fileName')
        else:
            self.fileName = input('Enter the name of the file: ')
    
    def inputDiagCoord(self, **kwargs):
        if kwargs.get('diagCoord'):
            self.diagCoord = kwargs.get('diagCoord')
        else:
            self.diagCoord = [float(x) for x in input('Enter the diagonal coordinates of the cut: ').split()]
    
    def inputEdgeCoord(self, **kwargs):
        if kwargs.get('edgeCoord'):
            self.edgeCoord = kwargs.get('edgeCoord')
        else:
            self.edgeCoord = [float(x) for x in input('Enter the edge coordinates of the cut: ').split()]

    def inputCutDirection(self, **kwargs):
        if kwargs.get('cutDirection'):
            self.cutDirection = kwargs.get('cutDirection')
        else:
            self.cutDirection = input('Enter the cut direction (X, Y, Z): ')

    def inputPlaneShiftValue(self, **kwargs):
        if 'planeShiftValue' in kwargs.keys():
            self.planeShiftValue = kwargs.get('planeShiftValue')
        else:
            self.planeShiftValue = float(input('Enter the plane shift value: '))
    
    def inputCutStep(self, **kwargs):
        if 'cutStep' in kwargs.keys():
            self.cutStep = kwargs.get('cutStep')
        else:
            self.cutStep = float(input('Enter the cut step: '))
    
    def inputPlaneShift(self, **kwargs):
        if 'planeShift' in kwargs.keys():
            self.planeShift = kwargs.get('planeShift')
        else:
            self.planeShift = bool(input('Should the cut plane be shifted? (True/False): '))
    
    def inputCutDistance(self, **kwargs):
        if 'cutDelta' in kwargs.keys():
            self.cutDelta = kwargs.get('cutDelta')
        else:
            self.cutDelta = float(input('Enter the height from the plane where the cut should extend: '))

    def inputCutHeight(self, **kwargs):
        if 'cutH' in kwargs.keys():
            self.cutH = kwargs.get('cutH')
        else:
            self.cutH = float(input('Enter the height of the cut plane: '))

    def inputStartCoord(self, **kwargs):
        if 'startCoord' in kwargs.keys():
            self.startCoord = kwargs.get('startCoord')
        else:
            self.startCoord = float(input('Enter the start coordinate of the normal: '))
    
    def inputEndCoord(self, **kwargs):
        if 'endCoord' in kwargs.keys():
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
        if self.advAxisExists:
            if 'vec1' in kwargs.keys():
                self.vec1 = kwargs.get('vec1')
            else:
                self.vec1 = input('Enter the first Plane Vector Joint: ')
    
    def inputVec2(self, **kwargs):
        if self.advAxisExists:
            if 'vec2' in kwargs.keys():
                self.vec2 = kwargs.get('vec2')
            else:
                self.vec2 = input('Enter the second Plane Vector Joint: ')
    
    def inputLocalPlane(self, **kwargs):
        if self.advAxisExists:
            if kwargs.get('localPlane'):
                self.localPlane = kwargs.get('localPlane')
            else:
                self.localPlane = input('Enter the local plane: ')
        
    def inputAdvAxis(self, **kwargs):
        if 'advAxisExists' in kwargs.keys():
            self.advAxisExists = kwargs.get('advAxisExists')
        else:
            self.advAxisExists = input('Enter the advanced axis: ')

    def inputIs4Pt(self, **kwargs):
        if 'is4Pt' in kwargs.keys():
            self.is4Pt = kwargs.get('is4Pt')
        else:
            self.is4Pt = input('Enter if the cut is a 4 point cut: ')
    
    def inputIsCustom(self, **kwargs):
        if 'isCustom' in kwargs.keys():
            self.isCustom = kwargs.get('isCustom')
        else:
            self.isCustom = input('Enter if the cut is a 4 point cut: ')
    
    def input4PtCoord(self, **kwargs):
        self.quadCoord = []
        for param in kwargs.keys():
            self.quadCoord.append(kwargs.get(param))
        if not self.quadCoord:
            self.quadCoord = [float(x) for x in input('Enter the coordinates of the 4 points: ').split()]

    def inputSpecialCoord(self, **kwargs):
        if 'addSpecialCoord' in kwargs.keys():
            self.addSpecialCoord = kwargs.get('addSpecialCoord')
        else:
            self.addSpecialCoord = [float(x) for x in input('Enter the special coordinates to add: ').split()]
        if 'rmSpecialCoord' in kwargs.keys():
            self.rmSpecialCoord = kwargs.get('rmSpecialCoord')
        else:
            self.rmSpecialCoord = [float(x) for x in input('Enter the special coordinates to remove: ').split()]
    
    def checkQuadCoord(self):
        if len(self.quadCoord) != 8:
            raise ValueError('Please enter 8 coordinates for the 4 points')
        # Check if 4 points create a quadrilateral
        # The points should be in cyclic order

        return True
    
    def addQuadCoord(self, direction='Z', cutName = '', constCoord=0, coordList = [0,0]):
        if direction == 'Z':
            self.quad.loc[len(self.quad)] = [cutName, coordList[0], coordList[1], constCoord]
        if direction == 'Y':
            self.quad.loc[len(self.quad)] = [cutName, coordList[0], constCoord, coordList[1]]
        if direction == 'X':
            self.quad.loc[len(self.quad)] = [cutName, constCoord, coordList[0], coordList[1]]

    def defineCustomQuad(self, cutName, gridStart, gridEnd, z, delta=2):
        if not self.planeShift:
            d = 0
        else:
            d = self.planeShiftValue
        dx, dy = np.subtract(gridEnd, gridStart)
        length = np.hypot(dx, dy)

        if length == 0:
            raise ValueError('The length of the grid is 0')
        
        normal = d*np.array([-dy, dx])/length
        gridStart = np.add(gridStart, normal).tolist()
        gridEnd = np.add(gridEnd, normal).tolist()

        self.addQuadCoord('Z', cutName, z+delta, gridStart)
        self.addQuadCoord('Z', cutName, z+delta, gridEnd)
        self.addQuadCoord('Z', cutName, z-delta, gridEnd)
        self.addQuadCoord('Z', cutName, z-delta, gridStart)

    def define4PtCut(self, direction, cutName, constCoord, coordList):
        self.addQuadCoord(direction, cutName, constCoord, coordList[0:2])
        self.addQuadCoord(direction, cutName, constCoord, coordList[2:4])
        self.addQuadCoord(direction, cutName, constCoord, coordList[4:6])
        self.addQuadCoord(direction, cutName, constCoord, coordList[6:8])
    
    def define2PtCut(self, direction, cutName, constCoord, coordList):
        self.addQuadCoord(direction, cutName, constCoord, [coordList[0], coordList[1]])
        self.addQuadCoord(direction, cutName, constCoord, [coordList[0], coordList[3]])
        self.addQuadCoord(direction, cutName, constCoord, [coordList[2], coordList[3]])
        self.addQuadCoord(direction, cutName, constCoord, [coordList[2], coordList[1]])



    def defineCut(self):
        # create a list of cuts from the start to the end with the cut step include both start and end
        cutList = []
        if not self.isCustom:
            if round(self.cutStep,3) != 0:
                cutList = list(np.arange(self.startCoord, self.endCoord, self.cutStep))
                cutList = [round(x, 3) for x in cutList]
            cutList.append(self.endCoord)

            # Add the special coordinates to the cut list
            for c in self.addSpecialCoord:
                cutList.append(c)
            for c in self.rmSpecialCoord:
                if c in cutList:
                    cutList.remove(c)
        else:
            cutList = [self.cutH]
        
        for c in cutList:
            cutNameInList = f'{self.cutName} - {self.cutDirection}={c}{self.unit}'
            if self.is4Pt:
                self.define4PtCut(self.cutDirection, cutNameInList, c, self.quadCoord)
            elif not self.isCustom:
                self.define2PtCut(self.cutDirection, cutNameInList, c, self.diagCoord)
            else:
                cutNameInList = f'{self.cutName}'
                self.defineCustomQuad(cutName=cutNameInList, 
                                      gridStart = self.edgeCoord[0:2], 
                                      gridEnd = self.edgeCoord[2:4], 
                                      z = self.cutH, 
                                      delta = self.cutDelta)
            self.general.loc[len(self.general)] = [cutNameInList, 'Quad', self.groupName, 'Analysis', 'Yes',0,0,0,0,0,0,'', '', 'Positive']
            if self.advAxisExists:
                self.advAxis.loc[len(self.advAxis)] = [cutNameInList, ''.join(self.localPlane.split('-')), 'Coord Dir', 'GLOBAL', 'Z', 'None', 'None', 'Two Joints', 'GLOBAL', 'X', 'Y', self.vec1, self.vec2, 0, 0, 1, 1, 0, 0]
    
    def printExcel(self, fileLoc=''):
        with pd.ExcelWriter(fileLoc + f'\\{self.fileName}') as writer:
            self.general.to_excel(writer, index=False, sheet_name='Section Cuts 1 - General')
            self.advAxis.to_excel(writer, index=False, sheet_name='Section Cuts 2 - Advanced Axes')
            self.quad.to_excel(writer, index=False, sheet_name='Section Cuts 3 - Quadrilaterals')
        # Create an excel file with multiple sheets
        # The first sheet is the section cuts 
        # The second sheet is the group names

