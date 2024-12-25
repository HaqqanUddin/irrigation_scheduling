import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
import datetime
import numpy as np

# Function to calculate evapotranspiration (ET)
def calculate_et(crop_kc, etr):
    return etr * crop_kc

# Function to interpolate monthly data to daily data
def interpolate_monthly_to_daily(monthly_data, season_months):
    daily_data = []
    for i in range(len(season_months) - 1):
        start = monthly_data[i]
        end = monthly_data[i + 1]
        # Calculate the number of days in the month
        days_in_month = (datetime.date(2024, season_months[i + 1], 1) - datetime.date(2024, season_months[i], 1)).days
        # Generate daily values within the month
        daily_values = np.linspace(start, end, days_in_month, endpoint=False)
        daily_data.extend(daily_values)
    # Handle the transition from the last month to the next year
    start = monthly_data[-1]
    end = monthly_data[0]
    days_in_month = (datetime.date(2024, 12, 31) - datetime.date(2024, season_months[-1], 1)).days + 1
    daily_values = np.linspace(start, end, days_in_month, endpoint=False)
    daily_data.extend(daily_values[:days_in_month])  # Ensure this slice does not exceed the intended range
    return daily_data

# Main function to generate daily irrigation schedule for the growing season
def daily_irrigation_schedule(soil, crop, climate, season_months, rainfall_data):
    schedule = []
    daily_etr = interpolate_monthly_to_daily(climate['ETr'], season_months)
    daily_rainfall = interpolate_monthly_to_daily(rainfall_data, season_months)
    day_count = 1
    cumulative_soil_water_deficit = 0
    max_allowable_depletion = 0.7 * (soil['field_capacity'] - soil['wilting_point']) * crop['root_depth'] * 10

    for stage, properties in crop['growth_stages'].items():
        days = properties['days']
        kc = properties['kc']
        for day in range(days):
            daily_climate = {
                'ETr': daily_etr[day_count - 1],
                'Rainfall': daily_rainfall[day_count - 1]
            }

            # Calculate ETc
            etc = calculate_et(kc, daily_climate['ETr'])

            # Calculate effective rainfall
            effective_rainfall = max(daily_climate['Rainfall'] - 0, 0)

            # Determine net irrigation application
            soil_water_deficit = effective_rainfall - etc

            # Update cumulative soil water deficit
            cumulative_soil_water_deficit += soil_water_deficit

            # Check if irrigation is required
            if cumulative_soil_water_deficit > max_allowable_depletion:
                irrigation_required = True
                daily_irrigation_req = max_allowable_depletion
                applied_water = daily_irrigation_req
                cumulative_soil_water_deficit = 0  # Reset deficit after irrigation
            else:
                irrigation_required = False
                daily_irrigation_req = 0
                applied_water = 0

            schedule.append({
                'day': day_count,
                'growing_season': 'Kharif',
                'growth_stage': stage,
                'crop_type': crop['crop_type'],
                'kc': kc,
                'ETo (mm/day)': daily_climate['ETr'],
                'Crop water use (Etc) (mm/day)': etc,
                'Rainfall (mm)': daily_climate['Rainfall'],
                'Net Irrigation application (mm)': round(daily_irrigation_req, 2),
                'Cumulative soil water deficit (mm)': round(cumulative_soil_water_deficit, 2),
                'Irrigation Required': 'Yes' if irrigation_required else 'No'
            })
            day_count += 1
    return schedule

# Correct file path
file_path = r"D:\E\Master Courses\Semesters\Fifth Semester\CE-577 Irrigation System Design and Management\Term Project 2\Climatic_Data.xlsx"

# Read data from Excel file
soil_data = pd.read_excel(file_path, sheet_name="Soil")
crop_data = pd.read_excel(file_path, sheet_name="Crops")
climatic_data = pd.read_excel(file_path, sheet_name="Climate")

# Handle potential missing sheet or columns
try:
    rainfall_data = pd.read_excel(file_path, sheet_name="Rainfall")
except KeyError:
    print("Error: 'Rainfall' sheet or column not found in Excel file.")
    exit()

# Ask user for inputs
soil_type = input("Enter the soil type (sand, sandy loam, loam, clay loam, silty clay, clay): ").strip().lower()
crop_type = input("Enter the crop type (cotton, sugarcane, rice, maize, wheat): ").strip().lower()

# Extract soil properties
soil = soil_data[soil_data['soil_type'] == soil_type].iloc[0].to_dict()

# Extract crop properties
crop = crop_data[crop_data['crop_type'] == crop_type].iloc[0].to_dict()
crop['growth_stages'] = {
    'initial': {'kc': crop['initial_kc'], 'days': crop['initial_days']},
    'development': {'kc': crop['development_kc'], 'days': crop['development_days']},
    'mid-season': {'kc': crop['mid_season_kc'], 'days': crop['mid_season_days']},
    'late-season': {'kc': crop['late_season_kc'], 'days': crop['late_season_days']}
}

# Extract monthly climatic data for the Kharif season (April to September)
season_months = [4, 5, 6, 7, 8, 9]
monthly_etr = climatic_data['ETr'][0:6].tolist()

# Extract monthly rainfall data for the Kharif season (April to September)
monthly_rainfall = rainfall_data['Rainfall (mm)'][0:6].tolist()

# Update climatic conditions with ET and rainfall
climatic_conditions = {
    'ETr': monthly_etr,
    'Rainfall': monthly_rainfall
}

# Generate the irrigation schedule
schedule = daily_irrigation_schedule(soil, crop, climatic_conditions, season_months, monthly_rainfall)

# Create a DataFrame from the schedule
schedule_df = pd.DataFrame(schedule)

# Save to Excel file with specific formatting
output_file = r"D:\E\Master Courses\Semesters\Fifth Semester\CE-577 Irrigation System Design and Management\Term Project 2\Irrigation_Schedule.xlsx"

# Set up the workbook and worksheet
wb = Workbook()
ws = wb.active

# Write headers
headers = ['Day', 'Growing Season', 'Growth Stage', 'Crop Type', 'Kc', 'ETo (mm/day)', 'Crop water use (Etc) (mm/day)', 'Rainfall (mm)', 'Net Irrigation application (mm)', 'Cumulative soil water deficit (mm)', 'Irrigation Required']
for col_idx, header in enumerate(headers):
    ws.cell(row=1, column=col_idx+1, value=header)
    ws.cell(row=1, column=col_idx+1).font = Font(bold=True)

# Write data starting from row 2
for idx, row in schedule_df.iterrows():
    ws.append([
        row['day'],
        row['growing_season'],
        row['growth_stage'],
        row['crop_type'],
        row['kc'],
        round(row['ETo (mm/day)'], 2),  # Round ETo to 2 decimal places
        round(row['Crop water use (Etc) (mm/day)'], 2),  # Round Etc to 2 decimal places
        round(row['Rainfall (mm)'], 2),  # Round rainfall to 2 decimal places
        round(row['Net Irrigation application (mm)'], 2),  # Round irrigation depth to 2 decimal places
        round(row['Cumulative soil water deficit (mm)'], 2),  # Round soil water deficit to 2 decimal places
        row['Irrigation Required']
    ])

# Save the workbook
wb.save(output_file)
print(f"Irrigation Schedule results saved to {output_file}")