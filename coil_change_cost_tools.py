# Constants for coil change cost calculations
# 4.25 min/coil change, $33/hr, 13 people on line 3
LABOR_COST_PER_HOUR = 33  # labor cost per hour
LABOR_UNITS_PER_CHANGE = 13  # employees affected per coil change
HOURS_PER_CHANGE = 4.25 / 60  # hours per coil change (converted from minutes)
LABOR_COST_PER_CHANGE = (LABOR_COST_PER_HOUR * HOURS_PER_CHANGE) * LABOR_UNITS_PER_CHANGE