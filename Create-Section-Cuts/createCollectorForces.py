import pyodbc
import pandas as pd

# Connect to the database
folderLoc = r"W:\\2023\\23184 - Trojena Neom PBD\\3 Engineering\\1 Calculations\\_Stage 3C Calc Package (100%)\\5.0 - Diaphragm Design\\5.4 - Building 305 Diaphragm Design\\5.4.4 - Diaphragm Collector Design\\20250414\\"
fileName = "20250414_305_UB_Dia_TH.mdb"
THfile = folderLoc + fileName
modelName = "20250414_305_UB"
conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + THfile)
cursor = conn.cursor()

# Read into a dataframe
sql = 'SELECT SectionCut, OutputCase, StepNum, F1, F2, F3, M1, M2, M3 FROM [Section Cut Forces - Analysis]'
df = pd.read_sql_query(sql, conn)
cuts = df['SectionCut'].unique()

# Find cuts that have the follow pattern like XX-XXXX-East and XX-XXXX-West both
east_west_groups = {}
a_b_groups = {}

for cut in cuts:
    if cut.endswith('-East') or cut.endswith('-West'):
        base_name = cut.rsplit('-', 1)[0]
        east_west_groups.setdefault(base_name, set()).add(cut.split('-')[-1])

    if cut.endswith('a') or cut.endswith('b'):
        base_name = cut[:-1]
        a_b_groups.setdefault(base_name, set()).add(cut[-1])

# Finding the cuts that have both "-East" and "-West"
east_west_cuts = [base for base, suffixes in east_west_groups.items() if {'East', 'West'}.issubset(suffixes)]

# Finding the cuts that have both "a" and "b"
a_b_cuts = [base for base, suffixes in a_b_groups.items() if {'a', 'b'}.issubset(suffixes)]

# Combining the results
print(east_west_cuts)
print(a_b_cuts)

# take each cut in east_west_cuts
# for each cut get results for cut-East and cut-West. Subtract at each StepNum and OutputCase for F1, F2, F3. Take absolute and max for each cut and OutputCase

results = []

for base_cut in east_west_cuts:
    df_east = df[df['SectionCut'] == f"{base_cut}-East"]
    df_west = df[df['SectionCut'] == f"{base_cut}-West"]

    # Merge on OutputCase and StepNum
    merged_df = df_east.merge(df_west, on=['OutputCase', 'StepNum'], suffixes=('_East', '_West'))

    # Compute absolute difference for F1, F2, F3
    merged_df['AbsDiff_F1'] = round(abs(merged_df['F1_East'] + merged_df['F1_West']),0)
    merged_df['AbsDiff_F2'] = round(abs(merged_df['F2_East'] + merged_df['F2_West']),0)
    merged_df['AbsDiff_F3'] = round(abs(merged_df['F3_East'] + merged_df['F3_West']),0)
    merged_df['AbsDiff_M1'] = round(abs(merged_df['M1_East'] + merged_df['M1_West']),0)
    merged_df['AbsDiff_M2'] = round(abs(merged_df['M2_East'] + merged_df['M2_West']),0)
    merged_df['AbsDiff_M3'] = round(abs(merged_df['M3_East'] + merged_df['M3_West']),0)

    # Find max absolute difference for each OutputCase
    max_diff = merged_df.groupby('OutputCase')[['AbsDiff_F1', 'AbsDiff_F2', 'AbsDiff_F3', 'AbsDiff_M1', 'AbsDiff_M2', 'AbsDiff_M3']].max().reset_index()
    mce_cases = max_diff[max_diff['OutputCase'].str.contains("MCE", na=False)]

    # take an average of all the gorund motions
    # Compute the average of the absolute differences
    average_row = mce_cases[['AbsDiff_F1', 'AbsDiff_F2', 'AbsDiff_F3', 'AbsDiff_M1', 'AbsDiff_M2', 'AbsDiff_M3']].mean().round(0).to_frame().T

    # Assign "average" as the OutputCase
    average_row['OutputCase'] = 'Average MCE'

    # Append the average row to max_diff
    max_diff = pd.concat([max_diff, average_row], ignore_index=True)

    max_diff.insert(0, 'SectionCut', base_cut)  # Add base cut name
    results.append(max_diff)

for base_cut in a_b_cuts:
    df_a = df[df['SectionCut'] == f"{base_cut}a"]
    df_b = df[df['SectionCut'] == f"{base_cut}b"]

    # Merge on OutputCase and StepNum
    merged_df = df_a.merge(df_b, on=['OutputCase', 'StepNum'], suffixes=('_a', '_b'))

    # Compute absolute difference for F1, F2, F3
    merged_df['AbsDiff_F1'] = round(abs(merged_df['F1_a'] + merged_df['F1_b']),0)
    merged_df['AbsDiff_F2'] = round(abs(merged_df['F2_a'] + merged_df['F2_b']),0)
    merged_df['AbsDiff_F3'] = round(abs(merged_df['F3_a'] + merged_df['F3_b']),0)
    merged_df['AbsDiff_M1'] = round(abs(merged_df['M1_a'] + merged_df['M1_b']),0)
    merged_df['AbsDiff_M2'] = round(abs(merged_df['M2_a'] + merged_df['M2_b']),0)
    merged_df['AbsDiff_M3'] = round(abs(merged_df['M3_a'] + merged_df['M3_b']),0)

    # Find max absolute difference for each OutputCase
    max_diff = merged_df.groupby('OutputCase')[['AbsDiff_F1', 'AbsDiff_F2', 'AbsDiff_F3', 'AbsDiff_M1', 'AbsDiff_M2', 'AbsDiff_M3']].max().reset_index()
    mce_cases = max_diff[max_diff['OutputCase'].str.contains("MCE", na=False)]

    # take an average of all the gorund motions
    # Compute the average of the absolute differences
    # Compute average only for output cases containing "MCE"

    average_row = mce_cases[['AbsDiff_F1', 'AbsDiff_F2', 'AbsDiff_F3', 'AbsDiff_M1', 'AbsDiff_M2', 'AbsDiff_M3']].mean().round(0).to_frame().T

    
    # Assign "average" as the OutputCase
    average_row['OutputCase'] = 'Average MCE'

    # Append the average row to max_diff
    max_diff = pd.concat([max_diff, average_row], ignore_index=True)

    max_diff.insert(0, 'SectionCut', base_cut)  # Add base cut name
    results.append(max_diff)

# Combine results into a single DataFrame
final_df = pd.concat(results, ignore_index=True)

# Display final result
print(final_df)

# Melt the dataframe to have F1, F2, and F3 as rows
melted_df = final_df.melt(id_vars=['SectionCut', 'OutputCase'], value_vars=['AbsDiff_F1', 'AbsDiff_F2', 'AbsDiff_F3', 'AbsDiff_M1', 'AbsDiff_M2', 'AbsDiff_M3'], 
                     var_name='Force', value_name='Value')

# Pivot to make OutputCase values into columns
pivot_df = melted_df.pivot(index=['SectionCut', 'Force'], columns='OutputCase', values='Value')

# Reset index for better readability
pivot_df.reset_index(inplace=True)

#

# Save output as excel
pivot_df.to_excel(folderLoc + f"{modelName}_CollectorResults.xlsx", index=False)


