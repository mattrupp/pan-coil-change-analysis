# Constants for coil change cost calculations
# 4.25 min/coil change, $33/hr, 13 people on line 3
LABOR_COST_PER_HOUR = 33  # labor cost per hour
LABOR_UNITS_PER_CHANGE = 13  # employees affected per coil change
HOURS_PER_CHANGE = 4.25 / 60  # hours per coil change (converted from minutes)
LABOR_COST_PER_CHANGE = (LABOR_COST_PER_HOUR * HOURS_PER_CHANGE) * LABOR_UNITS_PER_CHANGE

def get_coil_color(coil_color_code):
    """Returns the coil color based on the provided color code."""
    color_mapping = {
        'PA': 'Almond',
        'PR': 'Bronze (Terratone)',
        'PB': 'Brown',
        'PD': 'Charcoal',
        'PG': 'Gray', 
        'PH': 'Hunter Green',
        'PK': 'Carbon Black',
        'PM': 'Espresso',
        'PS': 'Sandstone',
        'PT': 'Sahara Tan',
        'PW': 'Polar White',
        'UU': 'Custom Color',
        'WE': 'English Oak',
        'WF': 'Graywood',
        'WO': 'Oak Woodgrain',
        'WN': 'Embossed Ash',
        'WQ': 'American Walnut',
        'WY': 'Cherry Woodgrain',
        'WX': 'Embossed Mahogany',
        'P1': 'Trinar White',
        'P2': 'Trinar Brown',
        'P3': 'Trinar Beige',
        'P4': 'Trinar Polar White',
        'P5': 'Trinar Almond'
    }
    return color_mapping.get(coil_color_code)
