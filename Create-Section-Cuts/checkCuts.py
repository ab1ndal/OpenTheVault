import pyodbc
import pandas as pd
import matplotlib.pyplot as plt

# Connect to the database
fileName = r"C:\\Users\\abindal\\OneDrive - Nabih Youssef & Associates\\Desktop\\20250220_205_LB_FineMesh_TH_Export_N13E_All.mdb"
conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + fileName)
cursor = conn.cursor()

# Read into a dataframe
sql = 'SELECT SectionCut, OutputCase, StepNum, F1, F2, F3 FROM [Section Cut Forces - Analysis]'
df = pd.read_sql_query(sql, conn)
print(df.head())
print(df['SectionCut'].unique())

GM = 'MCE-GM01'
plot_diaphragm = False
plot_diaphragm_close = False
plot_SteelOnly = False
plot_all = True
plot_collector = True
plot_all_close = False
plot_top_bot = False
plot_tests_diff = True
plot_tests = False
plot_col = ['F1', 'F2','F3']
plot_diff = True


if plot_diaphragm:
    cut1 = '6-13E-East'
    cut2 = '6-13E-West'
    dia_east = df[(df['SectionCut'] == cut1) & (df['OutputCase'] == GM)].reset_index()
    dia_west = df[(df['SectionCut'] == cut2) & (df['OutputCase'] == GM)].reset_index()
if plot_diaphragm_close:
    cut1_2 = '6-13E-East-Close'
    cut2_2 = '6-13E-West-Close'
    dia_east_2 = df[(df['SectionCut'] == cut1_2) & (df['OutputCase'] == GM)].reset_index()
    dia_west_2 = df[(df['SectionCut'] == cut2_2) & (df['OutputCase'] == GM)].reset_index()
if plot_SteelOnly:
    cut3 = '6-13E-East-SteelOnly'
    cut4 = '6-13E-West-SteelOnly'
    dia_east_steel = df[(df['SectionCut'] == cut3) & (df['OutputCase'] == GM)].reset_index()
    dia_west_steel = df[(df['SectionCut'] == cut4) & (df['OutputCase'] == GM)].reset_index()
if plot_all:
    cut5 = '6-13E-East-All'
    cut6 = '6-13E-West-All'
    dia_east_all = df[(df['SectionCut'] == cut5) & (df['OutputCase'] == GM)].reset_index()
    dia_west_all = df[(df['SectionCut'] == cut6) & (df['OutputCase'] == GM)].reset_index()

if plot_all_close:
    cut5_2 = '6-13E-East-Close-All'
    cut6_2 = '6-13E-West-Close-All'
    dia_east_all_2 = df[(df['SectionCut'] == cut5_2) & (df['OutputCase'] == GM)].reset_index()
    dia_west_all_2 = df[(df['SectionCut'] == cut6_2) & (df['OutputCase'] == GM)].reset_index()

if plot_collector:
    cut7 = 'N13E_Collector'
    col_n13e = df[(df['SectionCut'] == 'N13E_Collector') & (df['OutputCase'] == GM)].reset_index()
    max_val = col_n13e['StepNum'].max()

if plot_tests:
    col_test_left = df[(df['SectionCut'] == '****TestCutLeft') & (df['OutputCase'] == GM)].reset_index()
    col_test_right = df[(df['SectionCut'] == '****TestCutRight') & (df['OutputCase'] == GM)].reset_index()

if plot_tests_diff:
    col_test_left_diff = df[(df['SectionCut'] == '****TestCutLeft&Right') & (df['OutputCase'] == GM)].reset_index()
    col_test_top_diff = df[(df['SectionCut'] == '****TestCutTop&Bot') & (df['OutputCase'] == GM)].reset_index()

if plot_top_bot:
    cut8 = '6-13E-Top'
    cut9 = '6-13E-Bottom'
    dia_top = df[(df['SectionCut'] == cut8) & (df['OutputCase'] == GM)].reset_index()
    dia_bot = df[(df['SectionCut'] == cut9) & (df['OutputCase'] == GM)].reset_index()

# Plot motions
fig, ax = plt.subplots(figsize = (12,10), nrows = 3, ncols = 1)

for i in range(3):

    if plot_diaphragm:
        ax[i].plot(dia_east['StepNum'], 0*dia_east[plot_col[i]], color = 'black', linewidth = 0.25, linestyle = 'dashed')
        if plot_diff:
            max_diff_df = abs(dia_east[plot_col[i]] + dia_west[plot_col[i]])
            ax[i].plot(dia_east['StepNum'], max_diff_df, color = 'red', linewidth = 0.5, label = 'Diaphragm Cut')
            max_diff = max_diff_df.max()
            max_diff_index = max_diff_df.idxmax()
            ax[i].plot(dia_east['StepNum'][max_diff_index], max_diff_df[max_diff_index], 'ro', 
                       label = f'Max Diff: {max_diff:.0f} @ {dia_east["StepNum"][max_diff_index]:.2f}s')
        else:
            ax[i].plot(dia_east['StepNum'], dia_east[plot_col[i]], color = 'red',linewidth = 0.5, label = 'east-dia-cut')
            ax[i].plot(dia_west['StepNum'], dia_west[plot_col[i]], color = 'red',linewidth = 0.5, linestyle = 'dashed', label = 'west-dia-cut')

    if plot_SteelOnly:
        if plot_diff:
            max_diff_df_steel = abs(dia_east_steel[plot_col[i]] + dia_west_steel[plot_col[i]])
            ax[i].plot(dia_east_steel['StepNum'], max_diff_df_steel, color = 'blue',linewidth = 0.5, label = 'inPlane Steel Framing Cut')   
            max_diff_steel = max_diff_df_steel.max()
            max_diff_steel_index = max_diff_df_steel.idxmax()
            ax[i].plot(dia_east_steel['StepNum'][max_diff_steel_index], max_diff_df_steel[max_diff_steel_index], 'bo', label = f'Max Diff: {max_diff_steel:.0f} @ {dia_east_steel["StepNum"][max_diff_steel_index]:.2f}s')
        else:
            ax[i].plot(dia_east_steel['StepNum'], dia_east_steel[plot_col[i]], color = 'blue',linewidth = 0.5, label = 'east-dia-cut-steel')
            ax[i].plot(dia_west_steel['StepNum'], dia_west_steel[plot_col[i]], color = 'blue',linewidth = 0.5, linestyle = 'dashed', label = 'west-dia-cut-steel')

    if plot_all_close:
        if plot_diff:
            max_diff_df_all_close = abs(dia_east_all_2[plot_col[i]] + dia_west_all_2[plot_col[i]])
            ax[i].plot(dia_east_all_2['StepNum'], max_diff_df_all_close, color = 'olive', linewidth = 0.5, label = 'inPlane Cut Close')
            max_diff_close = max_diff_df_all_close.max()
            max_diff_close_index = max_diff_df_all_close.idxmax()
            ax[i].plot(dia_east_all_2['StepNum'][max_diff_close_index], max_diff_df_all_close[max_diff_close_index], 'o', color = 'olive', label = f'Max Diff: {max_diff_close:.0f} @ {dia_east_all_2["StepNum"][max_diff_close_index]:.2f}s')
        else:
            ax[i].plot(dia_east_all_2['StepNum'], dia_east_all_2[plot_col[i]], color = 'olive',linewidth = 0.5, label = 'east-dia-cut-close-all')
            ax[i].plot(dia_west_all_2['StepNum'], dia_west_all_2[plot_col[i]], color = 'olive',linewidth = 0.5, linestyle = 'dashed', label = 'west-dia-cut-close-all')

    if plot_all:
        if not plot_diff:
            ax[i].plot(dia_east_all['StepNum'], dia_east_all[plot_col[i]], color = 'green',linewidth = 0.5, label = 'east-dia-cut-all')
            ax[i].plot(dia_west_all['StepNum'], dia_west_all[plot_col[i]], color = 'green',linewidth = 0.5, linestyle = 'dashed', label = 'west-dia-cut-all')
        else:
            max_diff_df_all = abs(dia_east_all[plot_col[i]] + dia_west_all[plot_col[i]])
            ax[i].plot(dia_east_all['StepNum'], max_diff_df_all, color = 'green',linewidth = 0.5, label = 'inPlane Cut')
            max_diff_all = max_diff_df_all.max()
            max_diff_all_index = max_diff_df_all.idxmax()
            ax[i].plot(dia_east_all['StepNum'][max_diff_all_index], max_diff_df_all[max_diff_all_index], 'go', label = f'Max Diff: {max_diff_all:.0f} @ {dia_east_all["StepNum"][max_diff_all_index]:.2f}s')
    
    if plot_collector:
        ax[i].plot(col_n13e['StepNum'], abs(col_n13e[plot_col[i]]), color = 'purple',linewidth = 1, label = 'Collector Cut')
        max_col = abs(col_n13e[plot_col[i]]).max()
        max_col_index = abs(col_n13e[plot_col[i]]).idxmax()
        ax[i].plot(col_n13e['StepNum'][max_col_index], abs(col_n13e[plot_col[i]][max_col_index]), 'o', color = 'purple', label = f'Max Collector: {max_col:.0f} @ {col_n13e["StepNum"][max_col_index]:.2f}s')
    
    if plot_tests:
        ax[i].plot(col_test_left['StepNum'], col_test_left[plot_col[i]], color = 'cyan', linewidth = 1, label = 'Test Cut Left')
        ax[i].plot(col_test_right['StepNum'], col_test_right[plot_col[i]], color = 'cyan', linewidth = 1, linestyle = 'dashed', label = 'Test Cut Right')
    
    if plot_tests_diff:
        ax[i].plot(col_test_left_diff['StepNum'], abs(col_test_left_diff[plot_col[i]]), color = 'pink', linewidth = 0.5, label = 'Test Cut Left & Right')
        max_test_diff = abs(col_test_left_diff[plot_col[i]]).max()
        max_test_diff_index = abs(col_test_left_diff[plot_col[i]]).idxmax()
        ax[i].plot(col_test_left_diff['StepNum'][max_test_diff_index], abs(col_test_left_diff[plot_col[i]][max_test_diff_index]), 'o', color = 'pink', label = f'Max Test Diff: {max_test_diff:.0f} @ {col_test_left_diff["StepNum"][max_test_diff_index]:.2f}s')
        ax[i].plot(col_test_top_diff['StepNum'], abs(col_test_top_diff[plot_col[i]]), color = 'brown', linewidth = 0.5, label = 'Test Cut Top & Bot')
        max_test_top_diff = abs(col_test_top_diff[plot_col[i]]).max()
        max_test_top_diff_index = abs(col_test_top_diff[plot_col[i]]).idxmax()
        ax[i].plot(col_test_top_diff['StepNum'][max_test_top_diff_index], abs(col_test_top_diff[plot_col[i]][max_test_top_diff_index]), 'o', color = 'brown', label = f'Max Test Top Diff: {max_test_top_diff:.0f} @ {col_test_top_diff["StepNum"][max_test_top_diff_index]:.2f}s')

    if plot_top_bot:
        if not plot_diff:
            ax[i].plot(dia_top['StepNum'], dia_top[plot_col[i]], color = 'orange',linewidth = 0.5, label = 'top')
            ax[i].plot(dia_bot['StepNum'], dia_bot[plot_col[i]], color = 'orange',linewidth = 0.5, linestyle = 'dashed', label = 'bot')
        else:
            max_diff_top_bot_df = abs(dia_top[plot_col[i]] + dia_bot[plot_col[i]])
            ax[i].plot(dia_top['StepNum'], max_diff_top_bot_df, color = 'orange',linewidth = 0.5, label = 'Out-of-Plane Cut')
            max_diff_top_bot = max_diff_top_bot_df.max()
            max_diff_top_bot_index = max_diff_top_bot_df.idxmax()
            ax[i].plot(dia_top['StepNum'][max_diff_top_bot_index], max_diff_top_bot_df[max_diff_top_bot_index], 'o', color = 'orange', label = f'Max Diff: {max_diff_top_bot:.0f} @ {dia_top["StepNum"][max_diff_top_bot_index]:.2f}s')
    
    if plot_diaphragm_close:
        if plot_diff:
            max_diff_dia_close = abs(dia_east_2[plot_col[i]] + dia_west_2[plot_col[i]])
            ax[i].plot(dia_east_2['StepNum'], max_diff_dia_close, color = 'gold', linewidth = 0.5, label = 'Diaphragm Cut Close')
            max_diff_close = max_diff_dia_close.max()
            max_diff_close_index = max_diff_dia_close.idxmax()
            ax[i].plot(dia_east_2['StepNum'][max_diff_close_index], max_diff_dia_close[max_diff_close_index], 'o', color = 'gold', label = f'Max Diff: {max_diff_close:.0f} @ {dia_east_2["StepNum"][max_diff_close_index]:.2f}s')
        else:
            ax[i].plot(dia_east_2['StepNum'], dia_east_2[plot_col[i]], color = 'gold',linewidth = 0.5, label = 'east-dia-cut-close')
            ax[i].plot(dia_west_2['StepNum'], dia_west_2[plot_col[i]], color = 'gold',linewidth = 0.5, linestyle = 'dashed', label = 'west-dia-cut-close')

    ax[i].set_xlabel('Time Step (s)')
    ax[i].set_ylabel(f'{plot_col[i]} (kN)')
    ax[i].legend(loc = 'upper left', fontsize = 'x-small')
    ax[i].set_title(f'Comparison of {plot_col[i]} for N13E Cuts in 205[{GM}]')
    ax[i].set_xlim(0, max_val)
plt.tight_layout()
plt.savefig( r"C:\\Users\\abindal\\OneDrive - Nabih Youssef & Associates\\Desktop\\N13E_Cuts.png", dpi = 600)
plt.show()