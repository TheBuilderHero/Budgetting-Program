import pandas as pd
import numpy as np

def find_first_nonzero_reversed(df_temp):
    results = []
    for index_temp in reversed(df_temp.index):
        row_temp = df_temp.loc[index_temp]
        first_zero_index = -1
        for i_temp in range(len(row_temp)):
            if row_temp.iloc[i_temp] != 0:
                first_zero_index = i_temp
                break
        if first_zero_index != -1:
            results.append(first_zero_index)
        else:
            results.append(np.nan)
    return np.nanmax(results[::-1]) # Restore original order
            

# Example usage:
data = {'col1': [1, 2, 0, 4, 0], 
        'col2': [5, 0, 7, 8, 9], 
        'col3': [0, 11, 12, 0, 14]}
df = pd.DataFrame(data)

first_zeros = find_first_nonzero_reversed(df)
print(first_zeros)
# Expected Output: [0, 0, 0, 0, 0]

data_with_no_zeros = {'col1': [1, 2, 1, 4, 1], 
        'col2': [5, 1, 7, 8, 9], 
        'col3': [1, 11, 12, 1, 14]}
df_no_zeros = pd.DataFrame(data_with_no_zeros)

first_zeros_no_zeros = find_first_nonzero_reversed(df_no_zeros)
print(first_zeros_no_zeros)
# Expected Output: [nan, nan, nan, nan, nan]