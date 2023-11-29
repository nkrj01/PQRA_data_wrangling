import pandas as pd
import re

                                # Global variables #

# defining the variable for columns name for increased readablity without
# changing the column names in original data file
occ = "occur"
corr = "corr"
stri = "stri"
risk = "rl"
method = "MET"

path = r"C:\path\to\file"
unit_ops_key = pd.read_excel(path + r"\file.xlsx")
unit_ops = unit_ops_key["UOLP"].values.tolist()


def is_nan(strg):
    return strg != strg


def fillna_ce(df):

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

    if "Control" in df["variable"]:
        return "Control Element"
    elif "Occurrence" in df["variable"]:
        return "Occurrence"
    elif "Detection" in df["variable"]:
        return "Detection"
    else:
        return "Correlation"


def merge_occurrence_DTC(df):

    if is_nan(df["ODCD"]):
        return df[occ]
    else:
        return df["ODCD"]


def merge_occurrence_OS(df):

    if is_nan(df["OOC"]):
        return df["OLOC"]
    else:
        return df["OcOC"]


def value_string(df):
  

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
        return "ABP (vial)"
    elif "PFS" in df["Presentation"] and "654" in df["Product"]:
        return "ABP (PFS)"
    else:
        return df["Product"]


def fillna_overall_risk_level(df):
    """
    If correlation is NA or Testing only then fill the overall unit op risk
    level with NA

    """
    if "NA" in df[corr] or "Testing" in df[corr]:
        return "NA"
    else:
        return df[risk]


def helper_correlation(df):
    """
    Used to create a helper column that codes string to a particular integer. This helps with
    excel pivot table creation.
    """
    word = df[corr]
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
    letter = df[occ]
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
df[occ] = df.apply(merge_occurrence_DTC, axis=1)
df["OOC"] = df.apply(merge_occurrence_OS, axis=1)

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
df[corr] = df[corr].fillna("NA")
df[occ] = df[occ].fillna("NA")
df["Product"] = df.apply(separate_654_SKU, axis=1)
df[risk] = df.apply(fillna_overall_risk_level, axis=1)
df[risk] = df[risk].fillna("NA")
df["Helper_cor"] = df.apply(helper_correlation, axis=1)
df["Helper_dtc"] = df.apply(helper_occurence_code, axis=1)
df["Helper_risk"] = df.apply(helper_risk_score, axis=1)
df["Helper_Steri"] = df[stringency]
df["Helper_Stri_score"] = df["Helper_Stri"].fillna(10)
df["Sorting"] = df["unit_ops"].apply(sorting_column)
df = df.drop_duplicates()
df = df[~df[corr].str.contains("N/a")]

Exporting to Excel
df.to_excel( r"C:\outout_table.xlsx", index=False)
