## Script to calculate the sliding variance of two columns with some missing values ##


# Import of necessary libraries.

import numpy as np
import pandas as pd
import openpyxl 




# Loading the csv into a dataframe of pandas.

df = pd.read_csv('Data.csv', sep=';')

print('df:', df)


# Printing some general statistics of the both columns.

print('Statistical night description:', df['Night'].describe())
print('Statistical day description::', df['Day'].describe())



# Looking for missing values in the columns.

print('Missing values:\n', df.isnull().sum())


# There is one missing value in the middle of both columns with similar values before and after. 
# I decided to use the linear interpolation method for imputation. It is not bad that the new values aren't integers.
# The missing value at the end of the day column will be treated differently later. 

df = df.interpolate(method='linear')

print(df)



# Preparing the Night-Column by converting it to a numpy-array.

night = df['Night']
arr = night.to_numpy()


# Preparing the for-loop by building up a length for the iteration. 

length = arr.size

# Preparing a list to save the sliding variance.

var_night = []



# In the for-loop in each iteration five values of the Night-Column are taken to calculate the variance of it.
# This collection of five values slides from the beginning to the end of the column always taking the next and leaving the oldest row of the column.

for i in range(length):
    fives = arr[i:i+5]
    print('Values taken into account:', fives)
    
    variance = np.var(fives)         # The variance of the five values is calculated.   
    variance = variance.tolist()     # Return a copy of the array data as a Python list.
    #stadev = math.sqrt(variance)    # Standard-Deviation if wanted.
    
    variance = round(variance, 4)    # Rounding with four digits after the decimalpoint.
    var_night.append(variance)       # The calculated variances are saved in the list called var_night.    
    print('Variance:', variance)
    
print('Sliding variance of the night:', var_night)




# To prepare the export of the results the list is converted to a dataframe.

df_var_night = pd.DataFrame(var_night)

# The results are exported as an excel-file.

df_var_night.to_excel('var_night.xlsx', index=False)







# Same procedure as above for the Day-Column in the following.

day = df['Day']
arr = day.to_numpy()


# The missing value at the end of the column is deleted by reducing the length of the array. 
# Interpolation seemed for me to uncertain because of the lack of a following value.

arr = np.delete(arr, length-1)      # Attention for later usage!
length = arr.size


# Same procedure for the Day-Column as for the Night-Column.

var_day = []

for i in range(length):
    fives = arr[i:i+5]
    
    variance = np.var(fives)
    variance = variance.tolist()
    #stadev = math.sqrt(variance)    
    
    variance = round(variance, 4)
    var_day.append(variance)
    
print('Sliding variance of the day:', var_day)




# Export of the results as in the first column.

df_var_day = pd.DataFrame(var_day)
df_var_day.to_excel('var_day.xlsx', index=False)




