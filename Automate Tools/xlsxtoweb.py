import pandas as pd
import webbrowser
import time


excel_dosyasi = 'updated-tasks.xlsx'
sheet_name = 'Sheet1'


df = pd.read_excel(excel_dosyasi, sheet_name, header=0)

filtered_df = df[df['Görev Durumu'] == 'Başlandı']

filtered_df = filtered_df[['Görev No', 'Görev Durumu']]

filtered_df = filtered_df.sort_values(by='Görev No', ascending=True)

for index, row in filtered_df.iterrows():
    numara = row.iloc[0]
    base_url = 'https://ipm.omegamuhendislik.com.tr/Task/Record/'
    full_url = f"{base_url}{numara}"
    webbrowser.open(full_url)
    print(f"Web sayfası açılıyor: {full_url}")
    time.sleep(2)


