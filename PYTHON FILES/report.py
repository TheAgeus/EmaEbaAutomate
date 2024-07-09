import os

year = "2024"
month = "6"

rfc_folders = os.listdir(f"../DESCARGAS")
for rfc_folder in rfc_folders:
    full_route = f"../DESCARGAS/{rfc_folder}/{year}/{month}"
    try:
        number_files = len(os.listdir(full_route))
        print(f"{rfc_folder} => {number_files} files")
    except:
        print(f"{rfc_folder} => not even a rout")
    
x = input("Presiona cualquier tecla para terminar")