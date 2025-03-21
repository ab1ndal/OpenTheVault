import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from createCut import CreateCut

fileLocation = r"C:\\Users\\abindal\\OneDrive - Nabih Youssef & Associates\\Documents\\00_Projects\\06_The Vault\\Processing-Scripts\\OpenTheVault\\Create-Section-Cuts\\gridlineCutDefinition.xlsx"
spacing = 1
sheetName = '302'
groupName = 'All'
cutSuffix = 'All'
#addSpecialCoord = [-0.95, -10.2]
#rmSpecialCoord = [-1, -10]
addSpecialCoord = []
rmSpecialCoord =  []
GridList = []

outFileLoc = r'C:\\Users\\abindal\\OneDrive - Nabih Youssef & Associates\\Documents\\00_Projects\\06_The Vault\\Processing-Scripts\\OpenTheVault\\Create-Section-Cuts\\'


df = pd.read_excel(fileLocation, sheet_name=sheetName, header=0)
# add a new column to the dataframe for calculating theta using (Xs, Ys) and (Xe, Ye) in degrees
df['theta'] = df.apply(lambda row: np.degrees(np.arctan2(row['Ye'] - row['Ys'], row['Xe'] - row['Xs'])), axis=1)
df['Length'] = df.apply(lambda row: np.sqrt((row['Xe'] - row['Xs'])**2 + (row['Ye'] - row['Ys'])**2), axis=1)
df['R1x'] = df.apply(lambda row: round(row['Xe'] - spacing*np.sqrt(2) * np.sin(np.radians(row['theta']-45)),3), axis=1)
df['R1y'] = df.apply(lambda row: round(row['Ye'] + spacing*np.sqrt(2) * np.cos(np.radians(row['theta']-45)),3), axis=1)
df['R2x'] = df.apply(lambda row: round(row['R1x'] + 2*spacing * np.sin(np.radians(row['theta'])),3), axis=1)
df['R2y'] = df.apply(lambda row: round(row['R1y'] - 2*spacing * np.cos(np.radians(row['theta'])),3), axis=1)
df['R3x'] = df.apply(lambda row: round(row['R2x'] - (row['Length']+2*spacing) * np.cos(np.radians(row['theta'])),3), axis=1)
df['R3y'] = df.apply(lambda row: round(row['R2y'] - (row['Length']+2*spacing) * np.sin(np.radians(row['theta'])),3), axis=1) 
df['R4x'] = df.apply(lambda row: round(row['R3x'] - 2*spacing * np.sin(np.radians(row['theta'])),3), axis=1)
df['R4y'] = df.apply(lambda row: round(row['R3y'] + 2*spacing * np.cos(np.radians(row['theta'])),3), axis=1)


# Draw a rectangle with coordinates (R1x, R1y), (R2x, R2y), (R3x, R3y), (R4x, R4y) for each row in the dataframe
# Draw a line between (Xs, Ys) and (Xe, Ye) for each row in the dataframe
# Add the cut name to the rectangle and the line
# Save the drawing as a .pdf file

fig, ax = plt.subplots()
# set figure height and width
fig.set_figheight(10)
fig.set_figwidth(12)
for index, row in df.iterrows():
    ax.plot([row['Xs'], row['Xe']], [row['Ys'], row['Ye']], label=row['Gridlines'], color='black')
    ax.plot([row['R1x'], row['R2x'], row['R3x'], row['R4x'], row['R1x']], [row['R1y'], row['R2y'], row['R3y'], row['R4y'], row['R1y']], linestyle='--', color='red')
    ax.annotate('R1', (row['R1x'], row['R1y']))
    ax.annotate('R2', (row['R2x'], row['R2y']))
    ax.annotate('R3', (row['R3x'], row['R3y']))
    ax.annotate('R4', (row['R4x'], row['R4y']))
    ax.annotate(row['Gridlines'], ((row['Xs']+row['Xe']+spacing)/2, (row['Ys']+row['Ye']+spacing)/2),rotation=row['theta'], va='center', ha='center', color='blue')
ax.axis('equal')
# dont show the axes
plt.axis('off')
plt.tight_layout()
plt.savefig(outFileLoc+'cutDefinition.png', dpi=300)

# Create a new instance of CreateCut
cut = CreateCut(cutDirection='Z', cutStep=1.0, startCoord=-24.0, endCoord=126.0,
                groupName=groupName, unit='m', advAxisExists=True, localPlane='32',
                is4Pt=True, addSpecialCoord=addSpecialCoord, rmSpecialCoord=rmSpecialCoord)
if not GridList:
    GridList = df['Gridlines'].unique()
print(GridList)

for index, row in df.iterrows():
    if row['Gridlines'] in GridList:
        cut.inputCutName(cutName=row['Gridlines']+'-'+cutSuffix)
        cut.inputVec1(vec1=row['Point I'])
        cut.inputVec2(vec2=row['Point J'])
        cut.input4PtCoord(r1x=row['R1x'], 
                        r1y=row['R1y'],
                        r2x=row['R2x'],
                        r2y=row['R2y'],
                        r3x=row['R3x'],
                        r3y=row['R3y'],
                        r4x=row['R4x'],
                        r4y=row['R4y'])
        cut.inputIsCustom(isCustom=False)
        cut.inputPlaneShift(planeShift=False)
        cut.inputFileName(fileName='gridlineCuts.xlsx')
        cut.defineCut()
cut.printExcel(fileLoc=outFileLoc)

