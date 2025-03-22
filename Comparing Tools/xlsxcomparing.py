import pandas as pd

df1 = pd.read_excel('tasks.xlsx')
df2 = pd.read_excel('dif.xlsx')


merged_df = pd.merge(df1, df2, on='Görev Durumu', how='outer', indicator=True)

differences = merged_df[merged_df['_merge'] != 'both']

print(differences)
