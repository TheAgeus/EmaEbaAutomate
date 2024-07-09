import folders
import automate
import db

def main():
    print("Hello world")

    # Get list of downloaded files (.zip)
    print("Estos son los archivos descargados =>")
    folders.setDownloadedList()
    folders.printDownloadedList()

    print("Obteniendo clientes =>")
    db.setClientes()

    automate.downloadClientes()

    input("Precione cualquier tecla paa terminar.")

main()