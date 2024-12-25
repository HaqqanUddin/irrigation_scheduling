import numpy as np
import pandas as pd
import math

def modified_penman_method(T_max, T_min, RH_mean, E, z, U_day_night, U_z, R_s, R_n):
    """
    Calculate Evapotranspiration (ET) using the Modified Penman Method.

    Parameters:
    - T_max: Maximum Temperature (°C)
    - T_min: Minimum Temperature (°C)
    - RH_mean: Maximum Relative Humidity (%)
    - E: Elevation (m)
    - z: Height (m)
    - U_day_night: Day to Night Wind Ratio (daytime arbitrarily chosen as 0700-1900 hours)
    - U_z: Wind Speed at Certain Height (km/day)
    - R_s: Solar Radiation (cal/cm^2/day)
    - R_n: Net Solar Radiation (mm/day)

    Returns:
    - ET: Estimated Evapotranspiration (mm/day)
    """
    # Constants
    a1 = 0.39
    b1 = -0.05
    
    # Calculations
    T_mean = (T_max + T_min) / 2
    U_2 = U_z * (pow((2/z), 0.2))
    U_2day = (U_day_night / (U_day_night + 1)) * U_2 *(1000/43200)
    Delta = 2.00 * (pow((0.00738*T_mean) + 0.8072, 7)) - 0.00116
    P = 1013 - (0.1055 * E)
    L = 2500.78 - (2.3601 * T_mean)
    Gamma = 1.6134 * (P / L)
    C1 = Delta / (Delta + Gamma)
    C2 = 1 - C1
    R_s = (R_s * 41868) / (L * 1000)
    e_s = 33.8639 * (((0.00738 * T_mean + 0.8072)**8) - (0.000019 * ((1.8 * T_mean) + 48)) + 0.001316)
    e_a = e_s * (RH_mean / 100)
    c = 0.68 + (0.0028 * RH_mean) + (0.018 * R_s) - (0.068 * U_2day) + (0.013 * U_day_night) + (0.0097 * U_2day * U_day_night) + ((0.43*10**-4) * RH_mean * R_s * U_2day)
    ET_r = c * ((C1 * R_n) + (C2 * 0.27 * (1.0 + (0.01 * U_2)) * (e_s - e_a)))

    return ET_r

# Read data from Excel file
file_path = "D:/E/Master Courses/Semesters/Fifth Semester/CE-577 Irrigation System Design and Management/Term Project 2/Climatic_Data.xlsx"
sheet_name = "Kharif"  # Make sure the sheet name matches the actual sheet name in your Excel file

df = pd.read_excel(file_path, sheet_name=sheet_name)

# Calculate ET for each month and store results in a list
et_results = []

for index, row in df.iterrows():
    et = modified_penman_method(
        row['T_max'], 
        row['T_min'], 
        row['RH_mean'],
        row['E'], 
        row['z'], 
        row['U_day_night'], 
        row['U_z'], 
        row['R_s'], 
        row['R_n']
    )
    et_results.append(round(et,2))

# Create a new DataFrame with results
et_df = pd.DataFrame({
    'Month': df['Month'],
    'ETr': et_results
})

# Save the results to a new Excel file
output_file_path = "D:/E/Master Courses/Semesters/Fifth Semester/CE-577 Irrigation System Design and Management/Term Project 2/ETr_Monthly.xlsx"
et_df.to_excel(output_file_path, index=False)

print(f"ETr results saved to {output_file_path}")