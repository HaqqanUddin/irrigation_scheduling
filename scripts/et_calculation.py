# Import necessary libraries
import numpy as np  # For numerical operations, especially for mathematical functions
import pandas as pd  # For reading and manipulating data from Excel files
import math  # For mathematical functions (although not explicitly used in the code, it can be useful)

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
    # Constants for the Modified Penman equation
    a1 = 0.39  # Constant a1
    b1 = -0.05  # Constant b1
    
    # Calculate the average temperature
    T_mean = (T_max + T_min) / 2
    
    # Calculate wind speed at the 2m height (U_2)
    U_2 = U_z * (pow((2/z), 0.2))
    
    # Adjust wind speed for day-night ratio
    U_2day = (U_day_night / (U_day_night + 1)) * U_2 * (1000/43200)
    
    # Delta is the slope of the saturation vapor pressure curve
    Delta = 2.00 * (pow((0.00738*T_mean) + 0.8072, 7)) - 0.00116
    
    # Atmospheric pressure (P) at the given elevation
    P = 1013 - (0.1055 * E)
    
    # Latent heat of vaporization (L) as a function of temperature
    L = 2500.78 - (2.3601 * T_mean)
    
    # Psychrometric constant (Gamma)
    Gamma = 1.6134 * (P / L)
    
    # Coefficients used in the modified Penman equation
    C1 = Delta / (Delta + Gamma)
    C2 = 1 - C1
    
    # Adjust solar radiation (R_s) from cal/cm^2/day to mm/day
    R_s = (R_s * 41868) / (L * 1000)
    
    # Calculate the saturation vapor pressure (e_s)
    e_s = 33.8639 * (((0.00738 * T_mean + 0.8072)**8) - (0.000019 * ((1.8 * T_mean) + 48)) + 0.001316)
    
    # Calculate the actual vapor pressure (e_a) using relative humidity
    e_a = e_s * (RH_mean / 100)
    
    # Calculate the final ET using the modified Penman method
    c = 0.68 + (0.0028 * RH_mean) + (0.018 * R_s) - (0.068 * U_2day) + (0.013 * U_day_night) + (0.0097 * U_2day * U_day_night) + ((0.43*10**-4) * RH_mean * R_s * U_2day)
    
    # Calculate the Evapotranspiration (ET)
    ET_r = c * ((C1 * R_n) + (C2 * 0.27 * (1.0 + (0.01 * U_2)) * (e_s - e_a)))

    return ET_r  # Return the estimated evapotranspiration

# Read data from Excel file containing climatic data
file_path = "D:/E/Master Courses/Semesters/Fifth Semester/CE-577 Irrigation System Design and Management/Term Project 2/Climatic_Data.xlsx"
sheet_name = "Kharif"  # Specify the sheet name that contains the data

# Load the data into a pandas DataFrame
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Initialize an empty list to store ET results for each month
et_results = []

# Loop through each row in the DataFrame to calculate ET for each month
for index, row in df.iterrows():
    # Call the modified_penman_method function with values from the current row
    et = modified_penman_method(
        row['T_max'],  # Maximum temperature
        row['T_min'],  # Minimum temperature
        row['RH_mean'],  # Relative humidity
        row['E'],  # Elevation
        row['z'],  # Height
        row['U_day_night'],  # Day to night wind ratio
        row['U_z'],  # Wind speed at height z
        row['R_s'],  # Solar radiation
        row['R_n']  # Net solar radiation
    )
    
    # Append the rounded ET result to the list
    et_results.append(round(et, 2))

# Create a new DataFrame to store the calculated ET results alongside the month names
et_df = pd.DataFrame({
    'Month': df['Month'],  # Month column from the input data
    'ETr': et_results  # Estimated evapotranspiration results
})

# Define the path to save the output results to a new Excel file
output_file_path = "D:/E/Master Courses/Semesters/Fifth Semester/CE-577 Irrigation System Design and Management/Term Project 2/ETr_Monthly.xlsx"

# Save the results to an Excel file
et_df.to_excel(output_file_path, index=False)

# Print confirmation message
print(f"ETr results saved to {output_file_path}")