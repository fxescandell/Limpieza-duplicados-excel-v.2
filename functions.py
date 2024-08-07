# functions.py
import pandas as pd

def load_excel(file_path):
    """Carga un archivo Excel y devuelve un DataFrame de pandas.
    
    Args:
        file_path (str): La ruta del archivo Excel.

    Returns:
        DataFrame: Un DataFrame de pandas con los datos del archivo Excel.
    """
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error al cargar el archivo: {e}")
        return None
