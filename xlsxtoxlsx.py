import os
import pandas as pd

INPUT_FOLDER = r"c:\Users\Hexecoder\Desktop\Omega Apps\input_excels"
OUTPUT_FILE = r"c:\Users\Hexecoder\Desktop\Omega Apps\output.xlsx"
TARGET_CELLS = ["A1", "A2"]

def read_excel_cells(file_path, target_cells):
    try:
        df = pd.read_excel(file_path, header=None)
        cell_values = []
        for cell in target_cells:
            col_index = ord(cell[0].upper()) - 65
            row_index = int(cell[1:]) - 1
            cell_values.append(df.iloc[row_index, col_index])
        return cell_values
    except Exception as e:
        print(f"Error processing file {file_path}: {e}")
        return [None] * len(target_cells)

def collect_data(input_folder, target_cells):
    data = []
    for file_name in os.listdir(input_folder):
        if file_name.endswith(".xlsx"):
            file_path = os.path.join(input_folder, file_name)
            cell_values = read_excel_cells(file_path, target_cells)
            data.append([file_name] + cell_values)
    return data

def save_to_excel(output_file, data, headers):
    df = pd.DataFrame(data, columns=headers)
    df.to_excel(output_file, index=False)
    print(f"Data collected and saved to {output_file}")

if __name__ == "__main__":
    headers = ["File Name"] + [f"Cell {cell}" for cell in TARGET_CELLS]
    collected_data = collect_data(INPUT_FOLDER, TARGET_CELLS)
    save_to_excel(OUTPUT_FILE, collected_data, headers)