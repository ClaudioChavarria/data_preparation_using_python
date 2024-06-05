# %% [markdown]
# ### Import Pandas
# `Pandas` is a powerful data analysis and manipulation tool, offering easy-to-use data structures and analysis tools for Python.
# 

# %%
import pandas as pd

# %% [markdown]
# ### Define the file path

# %%
file_path = r"C:\Users\claud\Desktop\data_preparation_using_python\exercise_01_data_preparation_using_pandas\exercise_01.xlsx"

# %% [markdown]
# ### Load the excel file

# %%
# Load the Excel file
try:
    with pd.ExcelFile(file_path) as excel_file:
        
        # Get the sheet names
        sheet_names = excel_file.sheet_names
        print("Sheet Names in the Excel File:" , sheet_names)
        
        # Create a dictionary to store the DataFrames
        dataframes = {}

        # Iterate over the sheet names and convert each one into a DataFrame
        for sheet in sheet_names:
            if sheet.lower() != "diagrama": # Diagrama sheet excluded
                dataframes[sheet] = excel_file.parse(sheet)
            else:
                print(f"Sheet  {sheet} excluded.")

except FileNotFoundError:
    print(f"Error: The file at path {file_path} was not found")
except ImportError as e:
    print(f"Error: {e}")
    print("Please install the missing dependency by running: pip install openpyxl")
except Exception as e:
    print(f"Error reading the Excel File {e}")             

# %% [markdown]
# ### Preparation of 'DimCustomer'

# %%
dataframes['DimCustomer'].shape

# %%
dataframes['DimCustomer'].head(5)

# %%
# Select the columns 'customerKey, 'FirstName', 'LastName' and 'GeographyKey'
stg_dim_customer = dataframes['DimCustomer'].iloc[ :, [0, 1, 4, 6]].copy()

# Concatenate the columns 'FirstName' and 'LastName' in a single column 'Full Name'
stg_dim_customer['Full Name'] = stg_dim_customer[['FirstName', 'LastName']].apply(lambda x: ' '.join(x), axis=1)
# Remove the columns 'FirstName' and 'LastName
stg_dim_customer = stg_dim_customer.drop(columns=['FirstName', 'LastName'])

# Rename the columns
stg_dim_customer.columns = ['id_customer', 'id_geography', 'customer']

# Show the result
stg_dim_customer


# %% [markdown]
# ### Preparation of 'DimEmployee'

# %%
dataframes['DimEmployee'].shape

# %%
dataframes['DimEmployee'].head(5)

# %%
# Selec the columns 'EmployeeKey', 'FirstName', 'LastName', 'Department' and 'Position'
stg_dim_employee = dataframes['DimEmployee'].iloc[ :, [0 , 4, 5, -2, -1]].copy()

# Filter 'DepartmentName' is equal to Sales
stg_dim_employee = stg_dim_employee[stg_dim_employee.iloc[ :, -2].str.contains('Sales')]

# Filter 'Position' is equal to Sales Representative
stg_dim_employee = stg_dim_employee[stg_dim_employee.iloc[ :, -1].str.contains('Sales Representative')]

# Remove the columns 'DeparmentName' and 'Position'"
stg_dim_employee = stg_dim_employee.drop(columns= ['DepartmentName', 'Position'])

# Concatenate the columns 'FirstName and 'LastName' in a single column 'Employee'
stg_dim_employee['Employee'] = stg_dim_employee[['FirstName', 'LastName']].apply(lambda x: ' '.join(x), axis = 1)

# Remove the columns 'FirstName' and 'LastName'
stg_dim_employee = stg_dim_employee.drop(columns= ['FirstName', 'LastName'])

# Rename the columns
stg_dim_employee.columns = ['id_employee', 'Employee']

# Show the result
stg_dim_employee

# %% [markdown]
# ### Preparation of 'DimSalesTerritory'

# %%
dataframes['DimSalesTerritory'].shape

# %%
dataframes['DimSalesTerritory']

# %%
# Create the DataFrame
stg_dim_sales_territory = dataframes['DimSalesTerritory'].copy()

# Filter the columns SalesTerritoryRegion != 'Corporate HQ'
stg_dim_sales_territory = stg_dim_sales_territory[~stg_dim_sales_territory.iloc[: , 2].str.contains('Corporate HQ')]

# Remove SalesTerritoryAlternativeKey
stg_dim_sales_territory = stg_dim_sales_territory.drop(columns = ['SalesTerritoryAlternateKey'])

# Rename the columns
stg_dim_sales_territory.columns = ['id_territory' , 'Region', 'Country', 'Group' ]

# Show the result
stg_dim_sales_territory

# %% [markdown]
# ### Preparation of 'DimGeography'

# %%
dataframes['DimGeography'].shape

# %%
dataframes['DimGeography'].head(5)

# %%
# Select the columns 'geographyKey', 'City', 'StateProvinceName, 'EnglishCountryRegionName' and 'SalesTerritoryKey'
stg_dim_geography = dataframes['DimGeography'].iloc[: , [0, 1, 3, 5, 7]].copy()

stg_dim_geography

# %%
stg_dim_geography['StateProvinceName'].unique()

# %%
grouped_df = stg_dim_geography.groupby('StateProvinceName').size().reset_index(name='Count').sort_values(by='Count', ascending=False)
print(grouped_df)

# %%
stg_dim_geography['StateProvinceName'] = stg_dim_geography['StateProvinceName'].replace({'CALIFORNIA': 'California', 'Nueva York': 'New York'})

# %%
stg_dim_geography['StateProvinceName'].unique()

# %%
stg_dim_geography.columns = ['id_geography', 'city', 'state', 'Region', 'id_salesterritory']

# %%
stg_dim_geography.head()

# %% [markdown]
# ### Preparation of 'DimReseller'

# %%
dataframes['DimReseller'].shape

# %%
dataframes['DimReseller'].head(2)

# %%
stg_dim_reseller = dataframes['DimReseller'].iloc[: , [0, 1, 4, 5]].copy()
stg_dim_reseller.head(3)

# %%
grouped_reseller_df = stg_dim_reseller.groupby('BusinessType').size().reset_index(name= 'Count').sort_values(by = 'Count', ascending = False)
grouped_reseller_df

# %%
stg_dim_reseller['BusinessType'] = stg_dim_reseller['BusinessType'].replace({'Ware House': 'Warehouse'})


# %%
grouped_reseller_df = stg_dim_reseller.groupby('BusinessType').size().reset_index(name= 'Count').sort_values(by = 'Count', ascending = False)
grouped_reseller_df

# %%
stg_dim_reseller['BusinessType'] = stg_dim_reseller['BusinessType'].str.strip()
grouped_reseller_df = stg_dim_reseller.groupby('BusinessType').size().reset_index(name= 'Count').sort_values(by = 'Count', ascending = False)
grouped_reseller_df

# %%
stg_dim_reseller.columns = ['id_reseller', 'id_geography', 'Business Type', 'reseller']
stg_dim_reseller

# %% [markdown]
# ### Preparation of 'DimEmployeeSalesTerritory'

# %%
stg_dim_employee_salesterritory = dataframes['DimEmployeeSalesTerritory']
stg_dim_employee_salesterritory
stg_dim_employee_salesterritory.columns = ['id_employee', 'id_sales_territory']
stg_dim_employee_salesterritory.head(3)

# %% [markdown]
# ### Preparation of 'DimProduct'

# %%
dataframes['DimProduct'].shape

# %%
dataframes['DimProduct'].head(3)

# %%
stg_dim_product = dataframes['DimProduct'].iloc[: , [0, 5, 6, 7, 11]]
stg_dim_product.head(3)

# %%
stg_dim_product = stg_dim_product[stg_dim_product.iloc[:, 3].astype(str).str.contains('True')]
stg_dim_product.head(3)

# %%
stg_dim_product = stg_dim_product.iloc[: , [0, 1, 2, 4]]
stg_dim_product.columns = ['id_product', 'product', 'standard cost', 'list price']
stg_dim_product

# %% [markdown]
# ### Preparation of 'FactResellerSales'

# %%
stg_fact_reseller_sales = dataframes['FactResellerSales'].iloc[: , [0, 2, 3, 4, 5, 6, 8, 9, 10]]
stg_fact_reseller_sales.head(3)
stg_fact_reseller_sales.columns = ['id_order', 'Order date', 'Due date', 'Ship date', 'id_product', 'id_reseller', 
                                    'id_employee', 'id_sales_territory', 'Order Quantity']
stg_fact_reseller_sales.head(5)

# %% [markdown]
# ### Export data in an excel file

# %%
# Dataframes
clear_dataframes = [stg_dim_customer, stg_dim_employee, stg_dim_employee_salesterritory, stg_dim_geography,
                  stg_dim_reseller, stg_dim_product, stg_fact_reseller_sales]

# Name of Dataframes
sheet_names = ['dim_customer', 'dim_employee', 'dim_employee_sales_territory', 'dim_geography', 
               'dim_reseller', 'dim_product',  'fact_reseller_sales']

# export as a xlsx file
with pd.ExcelWriter('output_Exercise_01.xlsx') as writer:
    for df, sheet in zip(clear_dataframes, sheet_names):
        df.to_excel(writer, sheet_name=sheet, index=False)

print("Dataframes successfully exported to 'output_Exercise_01'")


