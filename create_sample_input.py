import pandas as pd

# Create sample input Excel file
data = {
    'Vehicle Number': [
        'RJ45CR3119',
        'DL3CAA1234',
        'MH12AB5678',
        'KA01MN9999'
    ]
}

df = pd.DataFrame(data)
df.to_excel('input.xlsx', index=False)
print("Sample input.xlsx created successfully")
