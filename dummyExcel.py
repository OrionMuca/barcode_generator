import pandas as pd
import numpy as np
import os

# Generate dummy data
np.random.seed(42)
data = {
    'Code': [f'CODE_{i}' for i in range(1, 11)],
    'Description': [f'Description {i}' for i in range(1, 11)],
    'Price': np.random.uniform(10, 100, 10)
}

# Create a DataFrame
df = pd.DataFrame(data)

# Specify the full path to your desktop
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

# Save the DataFrame to an Excel file on the desktop
excel_filename = os.path.join(desktop_path, 'dummy_data.xlsx')
df.to_excel(excel_filename, index=False)

print(f"Excel file saved to: {excel_filename}")
