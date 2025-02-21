import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from createCut import CreateCut

fileLocation = r".\\customCutLocations.xlsx"
sheetName = '205'
outFileLoc = r'.\\'
outFileName = 'customCutDefn.xlsx'
localPlane = '32'

df = pd.read_excel(fileLocation, sheet_name=sheetName, header=0)
print(df)
grids = pd.read_excel(fileLocation, sheet_name='GridInfo', header=0)

cut = CreateCut(cutDirection='Z', localPlane=localPlane)
cut.inputFileName(fileName = outFileName)
cut.inputIsCustom(isCustom = True)
cut.inputIs4Pt(is4Pt = False)
cut.inputAdvAxis(advAxisExists = True)

fig, ax = plt.subplots()
# set figure height and width
fig.set_figheight(10)
fig.set_figwidth(12)
for index, row in grids.iterrows():
    ax.plot([row['Xs'], row['Xe']], [row['Ys'], row['Ye']], label=row['Gridlines'], color='black', linewidth = 0.25, linestyle = 'dashed')
    # Calculate rotation angle for the text
    theta = np.degrees(np.arctan2(row['Ye'] - row['Ys'], row['Xe'] - row['Xs']))
    ax.annotate(row['Gridlines'], (row['Xs'], row['Ys']),rotation=theta, va='center', ha='right', color='blue')
ax.axis('equal')

for index, row in df.iterrows():
    cut.inputCutName(cutName=row['Name'])
    cut.inputVec1(vec1=row['Plane_Joint_1'])
    cut.inputVec2(vec2=row['Plane_Joint_2'])
    cut.inputPlaneShift(planeShift=row['Shift_Cut'])
    cut.inputPlaneShiftValue(planeShiftValue=row['Shift_Value'])
    cut.inputEdgeCoord(edgeCoord=[row['Start_X'], row['Start_Y'], row['End_X'], row['End_Y']])
    cut.inputCutHeight(cutH=row['Plane_Height'])
    cut.inputCutDistance(cutDelta=row['Z_Length']/2)
    cut.inputGroupName(groupName=row['Group'])
    cutStart, cutEnd = cut.defineCut()
    # Plot the cut
    ax.plot([cutStart[0], cutEnd[0]], [cutStart[1], cutEnd[1]], linestyle = 'solid', color = 'red', linewidth = 0.5)
    # Rotate and place the cut name on the line between cutStart and cutEnd offset by the text height
    # Calculate text rotation
    textRotation = np.degrees(np.arctan2(cutEnd[1] - cutStart[1], cutEnd[0] - cutStart[0]))
    # Calculate text position
    textX = cutStart[0]
    textY = cutStart[1]
    if row['Shift_Value'] > 0:
        va = 'bottom'
    else:
        va = 'top'
    # Place the text
    ax.annotate(row['Name'], (textX, textY),rotation=textRotation, va=va, ha='center', color='blue', fontsize=5)

# dont show the axes
plt.axis('off')
plt.tight_layout()
plt.savefig(outFileLoc+'customCutDefinition.png', dpi=300)

cut.printExcel(fileLoc=outFileLoc)
