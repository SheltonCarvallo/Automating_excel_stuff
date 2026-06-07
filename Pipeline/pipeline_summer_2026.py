#!C:\Users\Shelton\anaconda3\envs\DataScienceEnv\python.exe
import pandas as pd
import numpy as np
from pathlib import Path


PATH_SUMMER_2026 = "../raw_data/ReporteGeneralSummer2026.xls"
OUTPUT_PATH = "../data_cleaned/ReporteSUMMER2026Cleaned.csv"

COLUMNS_TO_DROP_REPORTE_GENERAL_SUMMER_2026 = [
    "ID", "Foto Perfil", "Traslado", "Informacion vacuna", "Telefono", "Comentario Pendiente",
    "Medio", "Nombre Amigo", "Con quien vieja",	"Amigo Viaja",
    "Entrevista", "DS160", "Tasa Consular", "Ademdum", 
    "Prueba pagada",	"Pago Inscripcion",	"Pago Entrevista",	"Pago DS2019", "Aprobacion Inscripcion", "Fecha emision visa",
    "Visa turista",	"Visa mision turista",	"Fecha expiracion visa", "Visa expiracion turista",	"Resultado psicometrico",	"Tipo Envio Formulario visa"
] 

COLUMNAS_FECHAS = ['Fecha Inicio Carrera', 'Fecha Final Carrera', 'Fecha Inscripcion', 'Fecha Nacimiento'] 

SKIP_ROWS = 6
SEPARATOR = "=" * 120

def pad_cedula_celular(numero) -> str:
    "Asegura que la cedula o número de celular tengan 10 dígito sin ingnorar el 0"
    return str(numero).zfill(10)

def find_datetime_columns_set_to_string(df) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame):
        raise TypeError(f"Se esperaba un DataFrame, se recibio un {type(df)}")
    
    mask = df.dtypes == "datetime64[ns]"

    time_columns = list(df.dtypes[mask].index)
    
    for column in time_columns:
        df[column] = df[column].apply(set_time_format)
    return df
    
def find_and_clean_str_columns(df) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame):
        raise TypeError(f"Se esperaba un DataFrame, se recibio un {type(df)}")
    
    mask = df.dtypes == "object"
    str_columns = list(df.dtypes[mask].index)
    for column in str_columns:
        df[column] = df[column].str.strip()
        df[column] = df[column].str.replace(r"\n|\r\n|\r", "", regex=True)
    return df

def set_time_format(time: pd.Timestamp) -> str:
    if not isinstance(time, pd.Timestamp):
        raise TypeError(f"Se esperaba un Timestamp, se recibió: {type(time)}")
    return time.strftime("%d/%m/%Y")


# ── I/O helpers ──────────────────────────────────────────────────────────────

def read_excel(excel_path: str, skip_rows: int, date_columns: list[str]) -> pd.DataFrame:
    "Load data from a excel file into a df"
    path = Path(excel_path)
    
    if not path.exists():
        raise FileNotFoundError(f"Reporte general no econtrado: {excel_path}")
    
    df = pd.read_excel(io=path, skiprows=skip_rows, parse_dates=date_columns)
    
    return df


def clean_df(df: pd.DataFrame, COLUMNS_TO_DROP_REPORTE_GENERAL: list[str] = None, celular: bool = True) -> pd.DataFrame:
    "It's main goal is to clean a dataframe"

    if not isinstance(df, pd.DataFrame):
        raise TypeError(f"Se esperaba un DataFrame, se recibio un {type(df)}")
    
    if df.empty:
        raise ValueError(f"El DataFrame está vacío, no se puede limpiar")
    
    # Clean
    print("Eliminando columnas innecesarias..")


    if isinstance(COLUMNS_TO_DROP_REPORTE_GENERAL, list):
        df = df.drop(columns=COLUMNS_TO_DROP_REPORTE_GENERAL)

    df = df.dropna(axis=1, how='all')
    df = df.dropna(axis=0, how='all')

    print("Seteando columnas numericas CI y Celular a columnas string, con el 0 por delante..")
    df['CI'] = df['CI'].apply(pad_cedula_celular)
    
    if celular:
        df['Celular'] = df['Celular'].apply(pad_cedula_celular)

    print(f"Seteando columnas de tipo datetime64[ns] a strings..")

    df = find_datetime_columns_set_to_string(df)     

    print("Eliminando espacios vacios por delante y detras de columnas string..")
    df = find_and_clean_str_columns(df)
    df['Nombre'] = df['Nombre'].str.lower()

    if df.empty:
        raise ValueError("El DataFrame quedó vacío después de la limpieza..")
    return df 


def save_csv(df: pd.DataFrame, path: str) -> None:
    if not isinstance(df, pd.DataFrame):
        raise TypeError(f"Se esperaba un DataFrame, se recibio un {type(df)}")    
    if df.empty:
        raise ValueError("El DataFrame se encuentra vacío")

    df.to_csv(path, index=False)
    print("Reporte guardado con éxito en formato csv")



def main() -> None:
    print(SEPARATOR)

    #Reporte General

    try:
        df = read_excel(PATH_SUMMER_2026, skip_rows=SKIP_ROWS, date_columns=COLUMNAS_FECHAS)
        print("Report general summer 2026 cargado")
    except FileNotFoundError as e:
        print(f"[ERROR] Ruta para leer excel file no econtrada => {e}")
        return
    except Exception as e:
        print(f"[ERROR] al cargar archivo de entrada => {e}")
        return
    
    print(SEPARATOR)
    
    try:
        print("Empezando limpieza df..")
        df = clean_df(df, COLUMNS_TO_DROP_REPORTE_GENERAL_SUMMER_2026)
        print("Dataframe 'limpio'")
        print(f"Número de filas {df.shape[0]}, Número de columnas {df.shape[1]}")
    except TypeError as e:
        print(f"[ERROR] al limpiar el dataframe => {e}")
        return
    except ValueError as e:
        print(f"[ERROR] al limpiar el dataframe => {e}")
        return
    except KeyError as e:
        print(f"[ERROR] columnas no existentes => {e}")
    except Exception as e:
        print(f"[ERROR] al limpiar el dataframe => {e}")
        return

    print(SEPARATOR)

    try:
        #print(df.loc[df['Nombre'] == "vivian carolina simbana catagna", ['Precio', 'Cualitativa']])
        save_csv(df, OUTPUT_PATH)
    except TypeError as e:
        print(f"[ERROR] al guardar el reporte excel en formato csv => {e}")
    except ValueError as e:
        print(f"[ERROR] al guardar el reporte excel en formato csv => {e}")
    except Exception as e:
        print(f"[ERROR] al guardar el reporte excel en formato csv => {e}")

if __name__ == "__main__":
    main()