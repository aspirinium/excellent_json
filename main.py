import pandas as pd
import geopandas as gpd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import platform
import sys

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def select_file():
    user_home = os.path.expanduser("~")
    username = os.getlogin()
    user_os = platform.system()

    if user_os == "Windows":
        default_directory = f"C:/Users/{username}/Develop/excellent_json"
    elif user_os == "Linux":
        default_directory = f"/home/{username}/Develop/excellent_json"
    else:
        default_directory = user_home 
    
    file_path = filedialog.askopenfilename(initialdir=default_directory)
    
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)


def replace_commas_with_dots_excel(data):
    for column in data.columns:
        if data[column].dtype in ['float64', 'int64']:
            data[column] = data[column].apply(lambda x: str(x).replace(',', '.') if isinstance(x, str) else x)
    return data

def replace_commas_with_dots(geo_data, numeric_columns):
    for column in numeric_columns:
        if column in geo_data.columns:
            geo_data[column] = geo_data[column].astype(str).replace({',': '.'}, regex=True)
    return geo_data

def convert_to_geojson():
    file_path = entry_file_path.get()
    if not file_path:
        messagebox.showerror("Error", "Please select an Excel file first.")
        return
    
    try:
        excel_data = pd.read_excel(file_path)
        numeric_columns = ['GesamthoeheM', "KEV-Liste", "Rotordurchmesser", 'TotalTurbinen', "MW", "GWhA"]
        for column in numeric_columns:
            excel_data[column] = pd.to_numeric(excel_data[column], errors='coerce')
    
        if "Kanton" in excel_data.columns:
            excel_data["Kanton"] = excel_data["Kanton"].fillna("").apply(
                lambda x: [k.strip() for k in x.split(",")] if isinstance(x, str) and x.strip() else []
            )

        excel_data = replace_commas_with_dots_excel(excel_data)
    
        excel_data["geometry"] = gpd.GeoSeries.from_wkt(excel_data["geometry"])
        geo_data = gpd.GeoDataFrame(excel_data, geometry=excel_data["geometry"])
        
        save_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        
        if save_path:
            geo_data.to_file(save_path, driver='GeoJSON')
            messagebox.showinfo("Success", "Excel file converted to JSON successfully!")
    except Exception as e:
        messagebox.showerror("Conversion Failed", f"Error: {str(e)}")

def convert_geojson_to_excel():
    file_path = entry_file_path.get()
    if not file_path:
        messagebox.showerror("Error", "Please select a GeoJSON file first.")
        return
    
    try:
        geo_data = gpd.read_file(file_path)

        if "Kanton" in geo_data.columns:
            geo_data["Kanton"] = geo_data["Kanton"].apply(
                lambda x: ", ".join(x) if isinstance(x, list) else x
            )
        
        numeric_columns = ['GesamthoeheM', "KEV-Liste", "Rotordurchmesser", 'TotalTurbinen', "MW", "GWhA"]
        geo_data = replace_commas_with_dots(geo_data, numeric_columns)
        
        
        geo_data['lat'] = geo_data['geometry'].apply(lambda x: x.y if x else None)
        geo_data['lon'] = geo_data['geometry'].apply(lambda x: x.x if x else None)
        
        for column in numeric_columns:
            geo_data[column] = pd.to_numeric(geo_data[column], errors='coerce')

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        
        if save_path:
            geo_data.to_excel(save_path, index=False)
            messagebox.showinfo("Success", "GeoJSON file converted to Excel successfully!")
    except Exception as e:
        messagebox.showerror("Conversion Failed", f"Error: {str(e)}")


root = tk.Tk()
root.title("File Converter")
root.geometry("500x400")

frame = tk.Frame(root)
frame.pack(pady=20)

entry_file_path = tk.Entry(frame, width=40)
entry_file_path.pack(side=tk.LEFT, padx=5)
btn_browse = tk.Button(frame, text="Browse", command=select_file, cursor="hand2")
btn_browse.pack(pady=10)

btn_convert_geojson = tk.Button(root, text="Convert Excel to JSON", command=convert_to_geojson, cursor="hand2")
btn_convert_geojson.pack(pady=10)

btn_convert_geojson_to_excel = tk.Button(root, text="Convert GeoJSON to Excel", command=convert_geojson_to_excel,  cursor="hand2")
btn_convert_geojson_to_excel.pack(pady=10)

root.mainloop()
