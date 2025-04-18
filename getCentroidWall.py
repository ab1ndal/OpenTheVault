import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
from matplotlib.patches import Polygon

def find_extreme_intersections(polygon_vertices, z_cut, reference_point):
    """
    Finds the extreme intersection points (leftmost & rightmost) of a horizontal plane at z_cut with a vertical polygon.

    :param polygon_vertices: List of (X, Y, Z) tuples defining the closed polygon.
    :param z_cut: Z-level of the cutting plane.
    :return: List of extreme intersection points [(X_min, Y, Z_cut), (X_max, Y, Z_cut)]
    """
    intersection_points = []

    for i in range(len(polygon_vertices)):
        v1 = np.array(polygon_vertices[i])
        v2 = np.array(polygon_vertices[(i + 1) % len(polygon_vertices)])  # Wrap-around at end

        # Check if the edge spans Z_cut
        if (v1[2] - z_cut) * (v2[2] - z_cut) <= 0:
            # Compute interpolation factor
            if v2[2] == v1[2]:
                #add both points if the edge is horizontal
                intersection_points.append((v1[0], v1[1], z_cut))
                intersection_points.append((v2[0], v2[1], z_cut))
            else:
                t = (z_cut - v1[2]) / (v2[2] - v1[2])

                # Compute intersection point
                x_cut = v1[0] + t * (v2[0] - v1[0])
                y_cut = v1[1] + t * (v2[1] - v1[1])

            intersection_points.append((x_cut, y_cut, z_cut))

    distance = lambda point: ((point[0] - reference_point[0]) ** 2 + (point[1] - reference_point[1]) ** 2) ** 0.5
    intersection_points.sort(key=distance)

    if len(intersection_points) >= 2:
        return [intersection_points[0], intersection_points[-1]]  # Leftmost and rightmost
    else:
        return intersection_points  # If only one point exists

def plot_projection(X, Y, Z, ax, reference_point, line = 'b-', linewidth=2, marker = '.'):
    """
    Plots 2D projection of the 3D polygon on the XY plane.
    X axis is the reference axis ie. distance from the reference point on xy plane
    Y axis is the height of the point from the reference point

    """
    # Get the reference point
    ref_x, ref_y, ref_z = reference_point

    # Compute the distance from the reference point
    dist = [((x - ref_x) ** 2 + (y - ref_y) ** 2) ** 0.5 for x, y in zip(X, Y)]

    # Compute the height from the reference point
    height = [z - ref_z for z in Z]

    ax.plot(dist, height, line, linewidth=linewidth, marker = marker, ms = 3)
    
    return dist, height

def plot_polygon_with_intersections(polygon, z_levels, reference_point, wall_name):
    """
    Plots the 3D polygon and its extreme intersections with horizontal planes at different Z-levels.

    :param polygon: List of (X, Y, Z) tuples defining the closed polygon.
    :param z_levels: List of Z-values where horizontal planes will cut through the polygon.
    """
    fig = plt.figure(figsize=(8, 8))
    ax = fig.add_subplot(111)

    # Extract X, Y, Z components
    X, Y, Z = zip(*polygon)

    # Close the polygon for plotting
    X_closed, Y_closed, Z_closed = list(X) + [X[0]], list(Y) + [Y[0]], list(Z) + [Z[0]]

    # Plot the polygon
    proj_dist, proj_height = plot_projection(X_closed, Y_closed, Z_closed, ax, reference_point)
    #ax.plot(X_closed, Y_closed, Z_closed, 'b-', marker = '.', label="Polygon Edges", linewidth=2)

    # Create polygon face
    verts = [list(zip(X_closed, Y_closed, Z_closed))]
    # Add a patch to the plot in 2D
    ax.add_patch(Polygon(list(zip(proj_dist, proj_height)), closed=True, fill=True, color='cyan', alpha=0.2))
    #ax.add_collection3d(Poly3DCollection(verts, alpha=0.2, facecolor='cyan'))

    centroid_list = []
    length_list = []

    # Process each Z-level intersection
    for z_cut in z_levels:
        extreme_points = find_extreme_intersections(polygon, z_cut, reference_point)
        if len(extreme_points) == 2:
            x_vals, y_vals, z_vals = zip(*extreme_points)
            plot_projection(x_vals, y_vals, z_vals, ax, reference_point, line='r--', linewidth=0.5, marker = 'o')
            #ax.plot(x_vals, y_vals, z_vals, 'r--', linewidth=2)
            #ax.scatter(x_vals, y_vals, z_vals, color='red', s=50)  # Mark extreme points
            #mark the centroid
            centroid = np.mean(extreme_points, axis=0)
            # Using extreme points calculate length between extreme points
            length_cut = ((extreme_points[0][0] - extreme_points[1][0])**2 + 
                          (extreme_points[0][1] - extreme_points[1][1])**2)**0.5
            
            #ax.scatter(centroid[0], centroid[1], centroid[2], color='black', s=50)
            ax.scatter(((centroid[0] - reference_point[0])**2+(centroid[1] - reference_point[1])**2)**0.5, (centroid[2] - reference_point[2]), color='black', s=20)
            centroid_list.append(centroid)
            length_list.append(length_cut)
        else:
            print("No extreme points found at Z =", z_cut)
            centroid_list.append((0,0,0))

    # Labels and view
    ax.set_xlabel('Distance from Reference Point on XY Plane')
    ax.set_ylabel('Height from the reference Point (m)')
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    #ax.set_zlabel('Z-axis')
    ax.set_title(f'Centroid for wall {wall_name}')

    # Set equal aspect
    #ax.set_box_aspect([1, 1, 1])
    ax.set_aspect('equal', 'box')
    plt.tight_layout()
    plt.savefig(f"Wall-{wall_name}.png", dpi=300)
    return centroid_list, length_list

# Example closed polygon with an opening

fileLocation = r'C:\\Users\\abindal\\OneDrive - Nabih Youssef & Associates\\Documents - The Vault\\Calculations\\2025 -  Stage 3C\\205 - SAP Models\\205_Walls.xlsx'
cut_Locs = r'C:\\Users\\abindal\\OneDrive - Nabih Youssef & Associates\\Documents - The Vault\\Calculations\\2025 -  Stage 3C\\205 - SAP Models\\SectionCutLocations_205.xlsx'

wallNames = ['205-N12A', '205-N12B', '205-N12C', '205-N12D', '205-N12', '205-N13A']
for wall in wallNames:
    print("Processing Wall:", wall)
    df = pd.read_excel(fileLocation, sheet_name=wall, header=0)
    cuts = pd.read_excel(cut_Locs, sheet_name=wall, header=0)

    # Extract X, Y, Z coordinates and create a list of tuples
    polygon = list(zip(df['GlobalX'], df['GlobalY'], df['GlobalZ']))
    reference_point = (df['GlobalX'][0], df['GlobalY'][0], df['GlobalZ'][0])

    # Choose different Z levels for horizontal planes
    z_levels = list(cuts['GlobalZ'])

    # Plot polygon with extreme intersection points
    centroid_list, length_list = plot_polygon_with_intersections(polygon, z_levels, reference_point,wall)
    # Add centroid_list to Cuts
    cuts['CentroidX'] = [centroid[0] for centroid in centroid_list]
    cuts['CentroidY'] = [centroid[1] for centroid in centroid_list]
    cuts['CentroidZ'] = [centroid[2] for centroid in centroid_list]
    cuts['Length'] = length_list
    #Create a new sheet in the excel file
    with pd.ExcelWriter(cut_Locs, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        cuts.to_excel(writer, sheet_name=f'{wall}_Centroids', index=False)