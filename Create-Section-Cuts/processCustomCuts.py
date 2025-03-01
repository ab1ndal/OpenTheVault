import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from createCut import CreateCut
from matplotlib.patches import Polygon
from pandasql import sqldf
import distinctipy

fileLocation = r".\\customCutLocations.xlsx"
sheetName = '205'
outFileLoc = r'.\\'
outFileName = 'customCutDefn.xlsx'
localPlane = '32'

# Add functionality to read local plane and the cut read direction

def read_file(file_path, sheet_name, colNames=None):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
        if colNames:
            df=df[colNames]
        return df
    except Exception as e:
        print(e)
        return None

df = read_file(file_path=fileLocation, sheet_name=sheetName)
floorelev = read_file(file_path=fileLocation, sheet_name='FloorElevations')
grids = read_file(file_path=fileLocation, sheet_name='GridInfo')
jointCoord = read_file(file_path=fileLocation, sheet_name='JointCoordinates', colNames=['Joint', 'GlobalX', 'GlobalY', 'GlobalZ'])
areasCoord = read_file(file_path=fileLocation, sheet_name='Slabs', colNames=['Area', 'NumJoints', 'Joint1', 'Joint2', 'Joint3', 'Joint4', 'CentroidZ'])
frameCoord = read_file(file_path=fileLocation, sheet_name='Beams', colNames=['Frame', 'JointI', 'JointJ', 'CentroidZ'])

#print(areasCoord)
#print(frameCoord)

# Round the values to 2 decimal places
areasCoord['CentroidZ'] = areasCoord['CentroidZ'].round(3)
frameCoord['CentroidZ'] = frameCoord['CentroidZ'].round(3)

query = 'Select areasCoord.*, p1.GlobalX as x1, p1.GlobalY as y1, p2.GlobalX as x2, p2.GlobalY as y2, p3.GlobalX as x3, p3.GlobalY as y3, p4.GlobalX as x4, p4.GlobalY as y4 ' \
        'from areasCoord '\
        'left join jointCoord as p1 on areasCoord.Joint1 = p1.Joint '\
        'left join jointCoord as p2 on areasCoord.Joint2 = p2.Joint '\
        'left join jointCoord as p3 on areasCoord.Joint3 = p3.Joint '\
        'left join jointCoord as p4 on areasCoord.Joint4 = p4.Joint'
slabCoord = sqldf(query, locals())

query = 'Select frameCoord.*, p1.GlobalX as x1, p1.GlobalY as y1, p2.GlobalX as x2, p2.GlobalY as y2 ' \
        'from frameCoord '\
        'left join jointCoord as p1 on frameCoord.JointI = p1.Joint '\
        'left join jointCoord as p2 on frameCoord.JointJ = p2.Joint'
beamCoord = sqldf(query, locals())

def draw_slab(ax, slabCoord, **kwargs):
    #slab_color = kwargs.get('slab_color', 'grey')
    alpha = 0.3

    slab_height = kwargs.get('slab_height', 0.0)

    floor_num = kwargs.get('floor_num', "001")

    min_value = floorelev[floorelev['FloorLabel'] == floor_num]['MinElev'].values[0]
    max_value = floorelev[floorelev['FloorLabel'] == floor_num]['MaxElev'].values[0] 

    # round to nearest 2 m
    #slab_height = round(slab_height/2)*2

    slab_at_specified_height = slabCoord[(slabCoord['CentroidZ'] >= min_value) & (slabCoord['CentroidZ'] <= max_value)].reset_index()

    unique_heights = sorted(slab_at_specified_height['CentroidZ'].unique())
    colors = distinctipy.get_colors(len(unique_heights), exclude_colors = [(0, 1, 0), (1, 1, 1)])
    height_color_map = dict(zip(unique_heights, colors))

    for i in range(len(slab_at_specified_height)):
        slab_color = height_color_map[slab_at_specified_height.CentroidZ[i]]
        if slab_at_specified_height.NumJoints[i] == 4.0:
            vertices = [[slab_at_specified_height.x1[i], slab_at_specified_height.y1[i]], 
                        [slab_at_specified_height.x2[i], slab_at_specified_height.y2[i]], 
                        [slab_at_specified_height.x3[i], slab_at_specified_height.y3[i]], 
                        [slab_at_specified_height.x4[i], slab_at_specified_height.y4[i]], 
                        [slab_at_specified_height.x1[i], slab_at_specified_height.y1[i]]]
            
        else:
            vertices = [[slab_at_specified_height.x1[i], slab_at_specified_height.y1[i]], 
                        [slab_at_specified_height.x2[i], slab_at_specified_height.y2[i]], 
                        [slab_at_specified_height.x3[i], slab_at_specified_height.y3[i]], 
                        [slab_at_specified_height.x1[i], slab_at_specified_height.y1[i]]]
            

        poly = Polygon(vertices, facecolor = slab_color, alpha = alpha, edgecolor = (0,0,0,1), 
                       linewidth=0.1, antialiased=True)
        ax.add_patch(poly)
    ax.set_aspect('equal')
    # Add a legend for colors and heights
    for height, color in height_color_map.items():
        ax.plot([], [], color=color, label=f'{height}m')
    ax.legend(title='Slab Heights', loc='upper right', fontsize='x-small')
    print(f'Slab drawn for z = "{unique_heights}m"')

def draw_beam(ax, beamCoord, **kwargs):
    beam_color = kwargs.get('beam_color', 'green')
    alpha = 0.3

    beam_height = kwargs.get('beam_height', 0.0)

    floor_num = kwargs.get('floor_num', "001")
    min_value = floorelev[floorelev['FloorLabel'] == floor_num]['MinElev'].values[0]
    max_value = floorelev[floorelev['FloorLabel'] == floor_num]['MaxElev'].values[0] 
    

    beam_at_specified_height = beamCoord[(beamCoord['CentroidZ'] >= min_value) & (beamCoord['CentroidZ'] <= max_value)].reset_index()

    for i in range(len(beam_at_specified_height)):
        # Plat beam as a straight line
        ax.plot([beam_at_specified_height.x1[i], beam_at_specified_height.x2[i]],
                [beam_at_specified_height.y1[i], beam_at_specified_height.y2[i]], 
                color = beam_color, linewidth = 0.2)
    ax.set_aspect('equal')
    print(f'Beam drawn for z = "{beam_height}m"')
    return



cut = CreateCut(cutDirection='Z')
cut.inputFileName(fileName = outFileName)
cut.inputIsCustom(isCustom = True)
cut.inputIs4Pt(is4Pt = False)
cut.inputAdvAxis(advAxisExists = True)

# Sort df by Plane_Height
df = df.sort_values(by='Plane_Height')
# set type of Floor to string
df['Floor'] = df['Floor'].astype(str)

image_list = []
for floor in df['Floor'].unique():
    df_floor = df[df['Floor'] == floor]
    print(f'Processing floor {floor}')

    fig, ax = plt.subplots()
    # set figure height and width
    fig.set_figheight(10)
    fig.set_figwidth(12)
    for index, row in grids.iterrows():
        ax.plot([row['Xs'], row['Xe']], [row['Ys'], row['Ye']], color='black', linewidth = 0.1, linestyle = 'dashed')
        # Calculate rotation angle for the text
        theta = np.degrees(np.arctan2(row['Ye'] - row['Ys'], row['Xe'] - row['Xs']))
        ax.annotate(row['Gridlines'], (row['Xs'], row['Ys']),rotation=theta, va='center', ha='right', color='blue')
    draw_slab(ax, slabCoord, floor_num=str(floor).zfill(3))#slab_height=round(df_floor['Plane_Height'].values[0],1))
    draw_beam(ax, beamCoord, floor_num=str(floor).zfill(3))
    ax.axis('equal')
    ax.set_title(f'Floor {floor}')

    for index, row in df_floor.iterrows():
        cut.inputElementSide(elementSide=row['elementSide'])
        cut.inputLocalPlane(localPlane=row['localPlane'])
        cut.inputCutName(cutName=row['Name'])
        cut.inputVec1(vec1=row['Plane_Joint_1'])
        cut.inputVec2(vec2=row['Plane_Joint_2'])
        cut.inputPlaneShift(planeShift=(row['Shift_Value']!=0))
        cut.inputPlaneShiftValue(planeShiftValue=row['Shift_Value'])
        cut.inputEdgeCoord(edgeCoord=[row['Start_X'], row['Start_Y'], row['End_X'], row['End_Y']])
        cut.inputCutHeight(cutH=row['Plane_Height'])
        cut.inputCutDistance(cutDelta=row['Z_Length']/2)
        cut.inputGroupName(groupName=row['Group'])
        cut.inputExtensionValue(extensionValue=row['Extend_Cut_Value'])
        cutStart, cutEnd = cut.defineCut()
        # Plot the cut
        ax.plot([cutStart[0], cutEnd[0]], [cutStart[1], cutEnd[1]], linestyle = 'solid', color = 'red', linewidth = 0.5)
        # Rotate and place the cut name on the line between cutStart and cutEnd offset by the text height
        # Calculate text rotation
        textRotation = np.degrees(np.arctan2(cutEnd[1] - cutStart[1], cutEnd[0] - cutStart[0]))
        # Calculate text position
        textX = cutStart[0]
        textY = cutStart[1]
        if 'Long' in row['Name']:
            va = 'bottom'
        else:
            va = 'top'
        # Place the text
        ax.annotate(row['Name'], (textX, textY),rotation=textRotation, va=va, ha='left', color='blue', fontsize=3)

    # dont show the axes
    plt.axis('off')
    plt.tight_layout()
    plt.savefig(outFileLoc+f'{floor.zfill(3)}_customCuts.png', dpi=300)
    # Add to a list of images
    image_list.append(outFileLoc+f'{floor.zfill(3)}_customCuts.png')
    plt.close()

cut.printExcel(fileLoc=outFileLoc)

# Create a pdf
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import os

# Define output PDF path
output_pdf_path = outFileLoc + "combined_customCuts_new2.pdf"

# Create a PDF
c = canvas.Canvas(output_pdf_path, pagesize=letter)

for image_path in image_list:
    # Open image using Pillow
    img = Image.open(image_path)
    img_width, img_height = img.size
    
    # Convert image size to fit within the PDF while preserving aspect ratio
    max_width, max_height = letter  # Page size
    scale = min(max_width / img_width, max_height / img_height)
    new_width, new_height = int(img_width * scale), int(img_height * scale)
    
    # Add a new page and draw the image
    c.drawImage(image_path, x=0, y=letter[1] - new_height, width=new_width, height=new_height)
    c.showPage()  # Save the current page

    # Remove the image
    img.close()
    os.remove(image_path)

# Save the final PDF
c.save()

print("All images converted into a high-quality single PDF")
print("Process completed")