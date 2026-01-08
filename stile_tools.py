import pandas as pd

# CONSTANTS FOR STILE DIMENSIONS, WEIGHTS, AND COSTS
# END STILES AND CENTER STILES HAVE THE SAME DIMENSIONS AND WEIGHTS
# Length in inches
STILE_18_LENGTH = 18.196
STILE_21_LENGTH = 21.196
STILE_24_LENGTH = 24.196

# Weight in pounds
STILE_18_33_WEIGHT = 0.925
STILE_21_33_WEIGHT = 1.077
STILE_24_33_WEIGHT = 1.230

STILE_18_44_WEIGHT = 1.243
STILE_21_44_WEIGHT = 1.448
STILE_24_44_WEIGHT = 1.653

STILE_18_55_WEIGHT = 1.547
STILE_21_55_WEIGHT = 1.802
STILE_24_55_WEIGHT = 2.057

# Cost per pound in dollars
STILE_33_COST_PER_POUND = 0.5641
STILE_44_COST_PER_POUND = 0.5558
STILE_55_COST_PER_POUND = 0.5549


def calculate_stile_cost(row):
    gauge = row['StileGauge']
    weight = row['StileWeight']

    if gauge == 3:
        cost_per_pound = STILE_33_COST_PER_POUND
    elif gauge == 4:
        cost_per_pound = STILE_44_COST_PER_POUND
    elif gauge == 5:
        cost_per_pound = STILE_55_COST_PER_POUND
    else:
        cost_per_pound = 0

    total_cost = cost_per_pound * weight
    return total_cost


def calculate_stile_weight(row):
    gauge = row['StileGauge']
    height = row['SectionHeight']
    quantity = row['StileQuantity']

    if gauge == 3:
        if height == 18:
            weight_per_stile = STILE_18_33_WEIGHT
        elif height == 21:
            weight_per_stile = STILE_21_33_WEIGHT
        elif height == 24:
            weight_per_stile = STILE_24_33_WEIGHT
        else:
            weight_per_stile = 0
    elif gauge == 4:
        if height == 18:
            weight_per_stile = STILE_18_44_WEIGHT
        elif height == 21:
            weight_per_stile = STILE_21_44_WEIGHT
        elif height == 24:
            weight_per_stile = STILE_24_44_WEIGHT
        else:
            weight_per_stile = 0
    elif gauge == 5:
        if height == 18:
            weight_per_stile = STILE_18_55_WEIGHT
        elif height == 21:
            weight_per_stile = STILE_21_55_WEIGHT
        elif height == 24:
            weight_per_stile = STILE_24_55_WEIGHT
        else:
            weight_per_stile = 0
    else:
        weight_per_stile = 0

    total_weight = weight_per_stile * quantity
    return total_weight


# Convert all the 3 gauge stiles to 4 gauge stiles and recalculate the stile weight and cost columns
def convert_3_to_4_gauge(row):
    gauge = row['StileGauge']
    height = row['SectionHeight']
    quantity = row['StileQuantity']

    if gauge == 3:
        # Convert to 4 gauge
        gauge = 4

    # Calculate the weight based on the new gauge
    if gauge == 4:
        if height == 18:
            weight_per_stile = STILE_18_44_WEIGHT
        elif height == 21:
            weight_per_stile = STILE_21_44_WEIGHT
        elif height == 24:
            weight_per_stile = STILE_24_44_WEIGHT
        else:
            weight_per_stile = 0
    elif gauge == 5:
        if height == 18:
            weight_per_stile = STILE_18_55_WEIGHT
        elif height == 21:
            weight_per_stile = STILE_21_55_WEIGHT
        elif height == 24:
            weight_per_stile = STILE_24_55_WEIGHT
        else:
            weight_per_stile = 0
    else:
        weight_per_stile = 0

    total_weight = weight_per_stile * quantity

    # Calculate the cost based on the new weight
    if gauge == 4:
        cost_per_pound = STILE_44_COST_PER_POUND
    elif gauge == 5:
        cost_per_pound = STILE_55_COST_PER_POUND
    else:
        cost_per_pound = 0

    total_cost = cost_per_pound * total_weight

    return pd.Series([gauge, total_weight, total_cost])


# Convert all the 3 and 4 gauge stiles to 5 gauge stiles and recalculate the stile weight and cost columns
def convert_3_4_to_5_gauge(row):
    gauge = row['StileGauge']
    height = row['SectionHeight']
    quantity = row['StileQuantity']

    if gauge == 3 or gauge == 4:
        # Convert to 5 gauge
        gauge = 5

    # Calculate the weight based on the new gauge
    if gauge == 5:
        if height == 18:
            weight_per_stile = STILE_18_55_WEIGHT
        elif height == 21:
            weight_per_stile = STILE_21_55_WEIGHT
        elif height == 24:
            weight_per_stile = STILE_24_55_WEIGHT
        else:
            weight_per_stile = 0
    else:
        weight_per_stile = 0

    total_weight = weight_per_stile * quantity

    # Calculate the cost based on the new weight
    if gauge == 5:
        cost_per_pound = STILE_55_COST_PER_POUND
    else:
        cost_per_pound = 0

    total_cost = cost_per_pound * total_weight

    return pd.Series([gauge, total_weight, total_cost])


def gen_stile_key_table(df, panel_locations_file, output=False):
    # Make a list of all the unique StileCodes
    unique_stile_codes = df['StileCode'].unique()

    # Create an ExcelFile object
    xls_file = panel_locations_file
    xls = pd.ExcelFile(xls_file)

    # Get the list of sheet names
    sheet_names = xls.sheet_names

    # Create an empty dictionary to hold the data for each stile code
    stile_dict = {}

    # Make a list of orphan stile codes
    orphan_stile_codes = []

    # For each unique stile code, check if there is a corresponding sheet in the Excel file
    # if there is, get the stile counts for each length and add to the stile dictionary
    for stile_code in unique_stile_codes:
        if stile_code in sheet_names:
            # add the excel tab length/stile count data to the dictionary
            stile_dict[stile_code] = xls.parse(sheet_name=stile_code, usecols=[
                2, 3], header=None, names=['Length', stile_code], skiprows=4)
            stile_dict[stile_code].set_index('Length', inplace=True)
        else:
            if output:
                print(
                    f"Sheet {stile_code} does NOT exist. Adding to orphan stile codes.")
            orphan_stile_codes.append(stile_code)

    # Concatenate all the stile dataframes into a single stile key table
    stile_key_table = pd.concat(stile_dict.values(), axis=1)

    # Get all the rows from df where there is no corresponding stile code sheet
    orphan_stile_df = df[df['StileCode'].isin(orphan_stile_codes)].copy()

    if output:
        output_file_name = 'data/orphan_stiles_sections.csv'
        orphan_stile_df.to_csv(output_file_name, index=False)
        print(f"Unique Stile Codes: {unique_stile_codes}")
        print(
            f"There were {len(orphan_stile_df)} sections with orphan stile codes added to the report: '{output_file_name}'.")

    return stile_key_table, orphan_stile_codes, orphan_stile_df
