import pandas as pd
import re


                                # Global variables #

# defining the variable for columns name for increased readablity without
# changing the column names in original data file
occurence = "Occurrence_Occurrence Decision Tree Code"
correlation = "Unit Operation_Correlation (↑,↓ or testing only)"
stringency = "Detection at Unit Operation_Stringency (i)"
risk_score = "Overall Unit Operation Risk Level_Overall Unit Operation Risk Level"
method_num = "Detection at Unit Operation_Detection Method"

path = r"C:\Users\nraj01\OneDrive - Amgen\Nikhil\Python related projects\pythonProject1\PQRA initiative"
unit_ops_key = pd.read_excel(path + r"\Unit operations key.xlsx")
unit_ops = unit_ops_key["Unit Operations Clean up"].values.tolist()


def is_nan(strg):
    """
    A function to check whether a cell of a dataframe is a string or NA
    Args: string
    return: bool
    """
    return strg != strg


def fillna_ce(df):
    """
    Changes 'x' of control elements in 1 and empties(or NA) to 0

    Args: DF
    return: binary

    """
    if df["variable1"] == "Control Element":
        if not (df["value"]):
            return 0
        elif df["value"] == "x" or df["value"] == "X":
            return 1
        else:
            return 0
    else:
        return df["value"]


def method_extract(df):
    """
    extract MET-XXXXXX text from a string of a column and store it in a different column
    Args: df
    return: string in form of MET-XXXXXX

    """
    pattern = re.compile(r"M[A-Z][A-Z]-\d\d\d\d\d\d")
    if not is_nan(df[method_num]):
        match = pattern.search(df[method_num])
        if match:
            return match.group(0)
        else:
            return ""
    else:
        return ""


def variable1(df):
    """
    Creates a column with variable1 name. Self explainatory.
    Args: df
    returns: string

    """

    if "Control" in df["variable"]:
        return "Control Element"
    elif "Occurrence" in df["variable"]:
        return "Occurrence"
    elif "Detection" in df["variable"]:
        return "Detection"
    else:
        return "Correlation"


def merge_occurrence_DTC(df):
    """
    Merge two decision tree code column into one.
    """

    if is_nan(df["Occurrence_Decision Tree Code"]):
        return df[occurence]
    else:
        return df["Occurrence_Decision Tree Code"]


def merge_occurrence_OS(df):
    """
    merge two occurrence score columns into one.
    """
    if is_nan(df["Occurrence_Occurrence Score"]):
        return df["Occurrence_Likelihood of Occurrence Score"]
    else:
        return df["Occurrence_Occurrence Score"]


def value_string(df):
    """
    This function is used to separate integer values and string values and store them in
    two different columns called "value" and "value_string". This is done to aid spotfire in visualization.

    """

    if is_nan(df["value"]):
        return ""
    else:
        if isinstance(df["value"], str):
            return df["value"]
        else:
            return ""


def remove_strings(df):
    """

    used to remove all the string datatype from value column

    """

    if isinstance(df["value"], str):
        return np.nan
    else:
        return df["value"]


def separate_654_SKU(df):
    """
    Separates 654 SKU in the dataset into PFS and vial

    """
    if "Vial" in df["Presentation"] and "654" in df["Product"]:
        return "ABP 654 (vial)"
    elif "PFS" in df["Presentation"] and "654" in df["Product"]:
        return "ABP 654 (PFS)"
    else:
        return df["Product"]


def fillna_overall_risk_level(df):
    """
    If correlation is NA or Testing only then fill the overall unit op risk
    level with NA

    """
    if "NA" in df[correlation] or "Testing" in df[correlation]:
        return "NA"
    else:
        return df[risk_score]


def helper_correlation(df):
    """
    Used to create a helper column that codes string to a particular integer. This helps with
    excel pivot table creation.
    """
    word = df[correlation]
    if word == "NA":
        return 1
    elif word == "↑":
        return 2
    elif word == "↓":
        return 3
    elif word == "↑↓":
        return 4
    elif word == "Testing only":
        return 5
    else:
        return 6

def helper_risk_score(df):
    """
    Used to create a helper column that codes string to a particular integer. This helps with
    excel pivot table creation.
    """
    word = df[risk_score]
    if word == "NA":
        return 1
    elif word == "Low":
        return 2
    elif word == "Medium":
        return 3
    elif word == "High":
        return 4
    else:
        return ""


def helper_occurence_code(df):
    """
    Used to create a helper column that codes string to a particular integer. This helps with
    excel pivot table creation.
    """
    letter = df[occurence]
    position_A = ord("A")

    if letter == "NA":
        return 15
    else:
        return ord(letter) - position_A + 1



def sorting_column(col, unit_ops_list=unit_ops):
    """
    A column that is inserted in the excel to sort the pivot table according to the unit operation.
    """
    index = 0
    for index, unit_op in enumerate(unit_ops_list):
        if unit_op == col:
            return index

    return "missing"


# import the main file, unit ops key, worksheet name key, and column name key
df = pd.read_excel(path + "\Compiled PQRA_modified.xlsx")
unit_ops_key = pd.read_excel(path + r"\Unit operations key.xlsx")
ws_name_key = pd.read_excel(path + r"\Worksheet name key.xlsx")
column_key = pd.read_excel(path + r"\Column select key.xlsx")

# Merge two decision tree columns into one. Do the same for Occurrence score column.
df[occurence] = df.apply(merge_occurrence_DTC, axis=1)
df["Occurrence_Occurrence Score"] = df.apply(merge_occurrence_OS, axis=1)

# Select the columns of interest
df = df[column_key["Columns"].values.tolist()]

# Select all the relevant unit ops using unit ops key file
unit_ops_key["count"] = 1
df = df.merge(unit_ops_key, how="left")
df = df[df["count"].notna()]

# Select all the quality attribute using worksheet name key file
ws_name_key["count2"] = 1
df = df.merge(ws_name_key, how="left")
df = df[df["count2"].notna()]

# setting up and cleaning the data for visualization
df[correlation] = df[correlation].fillna("NA")
df[occurence] = df[occurence].fillna("NA")
df["Product"] = df.apply(separate_654_SKU, axis=1)
df[risk_score] = df.apply(fillna_overall_risk_level, axis=1)
df[risk_score] = df[risk_score].fillna("NA")
df["Helper_correlation"] = df.apply(helper_correlation, axis=1)
df["Helper_decision tree code"] = df.apply(helper_occurence_code, axis=1)
df["Helper_Overall unit op risk"] = df.apply(helper_risk_score, axis=1)
df["Helper_Steringency score"] = df[stringency]
df["Helper_Steringency score"] = df["Helper_Steringency score"].fillna(10)
df["Sorting"] = df["Unit Operations Clean up"].apply(sorting_column)
df = df.drop_duplicates()
df = df[~df[correlation].str.contains("N/a")]

# Exporting to Excel
# df.to_excel(
#     r"C:\Users\nraj01\OneDrive - Amgen\Nikhil\Python related projects\pythonProject1\PQRA initiative\excel1 table.xlsx",
#     index=False,
# )