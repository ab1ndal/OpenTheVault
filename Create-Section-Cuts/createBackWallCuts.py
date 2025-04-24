from createCut import CreateCut
import pandas as pd

fileLoc = r'W:\\2023\\23184 - Trojena Neom PBD\\3 Engineering\\1 Calculations\\_Stage 3C Calc Package (100%)\\4.0 - Concrete Wall Design\\4.3 - Building 305 Concrete Wall Design\\4.3.1 - Wall shear design (in and out of plane)\\Model Outputs\\20250414\\'
fileName = 'BackWallCuts.xlsx'

import math

def create_shifted_rectangle(x1, y1, x2, y2, d):
    # Step 1: Direction vector
    dx = x2 - x1
    dy = y2 - y1

    # Step 2: Perpendicular vector (normal)
    length = math.hypot(dx, dy)
    nx = -dy / length
    ny = dx / length

    # Step 3: Scale by distance
    sx = nx * d
    sy = ny * d

    # Step 4: Calculate 4 corner points
    p1 = (x1 + sx, y1 + sy)
    p2 = (x2 + sx, y2 + sy)
    p3 = (x2 - sx, y2 - sy)
    p4 = (x1 - sx, y1 - sy)

    return [p1, p2, p3, p4]


file = fileLoc + fileName
db = pd.read_excel(file, sheet_name='Sheet1', header=0)
print(db.columns)

cut = CreateCut(unit='m', advAxisExists=True, localPlane='32', 
                is4Pt=True)
cut.inputIsCustom(isCustom=True)
cut.defineCustomCutName(renameCut=False)
shift = 0.25
for i in range(len(db)):
    cut.inputCutName(cutName=db['Cut Name'][i])
    cut.inputVec1(vec1=db['I'][i])
    cut.inputVec2(vec2=db['J'][i])
    cut.inputGroupName(groupName=db['Group Name'][i])
    cut.inputCentroidCoord(defaultLoc='No', globalX=db['Xa'][i], globalY=db['Ya'][i], globalZ=db['Za'][i])
    cut.inputCutHeight(cutH=db['Za'][i])
    p1, p2, p3, p4 = create_shifted_rectangle(db['Xi'][i], db['Yi'][i], db['Xj'][i], db['Yj'][i], shift)
    cut.input4PtCoord(x1 = p1[0],
                       y1 = p1[1],
                       x2 = p2[0],
                       y2 = p2[1],
                       x3 = p3[0],
                       y3 = p3[1],
                       x4 = p4[0],
                       y4 = p4[1])
    #cut.inputStartCoord(startCoord = db['Za'][i])
    #cut.inputEndCoord(endCoord = db['Za'][i])
    cut.defineCut()
cut.inputFileName(fileName = 'BackWallCuts_out.xlsx')
cut.printExcel(fileLoc = './')