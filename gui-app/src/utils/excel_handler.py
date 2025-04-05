import pandas as pd

def get_sheet_names(file_path):
    """Excel dosyasındaki sheet isimlerini döndürür."""
    try:
        excel_file = pd.ExcelFile(file_path, engine="openpyxl")  # Motoru açıkça belirt engine="openpyxl")  # Motoru açıkça belirt
        return excel_file.sheet_names
    except Exception as e:
        raise Exception(f"Failed to read sheet names: {e}")

def get_task_numbers(file_path, sheet_name):
    """Belirtilen sheet'ten görev numaralarını alır."""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")  # Motoru açıkça belirt engine="openpyxl")  # Motoru açıkça belirt
        task_numbers = df['Görev No'].dropna().astype(str).tolist()  # 'Task Number' sütununu al
        return task_numbers
    except Exception as e:
        raise Exception(f"Failed to read task numbers: {e}")