import pandas as pd

def read_file(file_path, sheet_name, colNames=None, header=0, native=False):
    try:
        if not native:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header)
            if colNames:
                df=df[colNames]
            return df
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
            # Remove the first row
            df = df.iloc[1:].reset_index(drop=True)  # Reset the index
            if colNames:
                df.columns = colNames
            return df
    except Exception as e:
        print(e)
        return None