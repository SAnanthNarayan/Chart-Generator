import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import griddata

# Read Excel data into a pandas DataFrame
file_path = 'your_file.xlsx'  # replace with your file path
df = pd.read_excel(file_path)

# Extract speed, torque, and temperature values
speed = df['speed'].values
torque = df['torque'].values
temperature = df['temperature'].values

# Create a grid for speed and torque
speed_grid = np.linspace(min(speed), max(speed), 100)  # 100 points for speed axis
torque_grid = np.linspace(min(torque), max(torque), 100)  # 100 points for torque axis
X_grid, Y_grid = np.meshgrid(speed_grid, torque_grid)

# Interpolate the temperature data onto the grid
Z_grid = griddata((speed, torque), temperature, (X_grid, Y_grid), method='cubic')

# Plot the 2D contour color map
plt.figure(figsize=(8, 6))
contour = plt.contourf(X_grid, Y_grid, Z_grid, 20, cmap='viridis')  # Adjust the number of contours (20) and colormap
plt.colorbar(contour)  # Add color bar
plt.title('Contour Plot of Temperature')
plt.xlabel('Speed')
plt.ylabel('Torque')
plt.show()
