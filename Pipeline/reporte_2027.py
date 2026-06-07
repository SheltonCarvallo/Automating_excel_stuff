#!C:\Users\Shelton\anaconda3\envs\DataScienceEnv\python.exe
import pandas as pd
import numpy as np
from pathlib import Path

PATH = "../raw_data/ReporteGeneral.xls"
NO_FIRMADO_PATH = "../raw_data/SinFirmarContrato.xls"
PRUEBA_INGLES_PATH = "../raw_data/Reporte prueba ingles.xlsx"

OUTPUT_PATH = "../data_cleaned/ReporteSpring2027Cleaned.csv"
OUTPUT_PATH_NO_FIRMADO = "../data_cleaned/ContratoNoFirmado2027Cleaned.csv"
OUTPUT_PATH_PRUEBAS_INGLES = "../data_cleaned/PruebasIngles2027Cleaned.csv"

SKIP_ROWS = 6
SEPARATOR = "=" * 120
COLUMNS_TO_DROP_REPORTE_GENERAL = ["Traslado", "Clave", "Nombre Amigo", "Con quien vieja", "Amigo Viaja",
                   "DS160", "Tasa Consular", "Ademdum", "Prueba pagada", "Pago Inscripcion", "Pago Entrevista",
                   "Aprobacion Inscripcion", "Cuenta convisa", "Fecha emision visa", "Fecha expiracion visa", 
                   "Visa mision turista"]

COLUMNS_TO_DROP_REPORTE_GENERAL_SIN_CONTRATO = ["Traslado", "Clave", "Nombre Amigo", "Con quien vieja", "Amigo Viaja",
                   "DS160", "Tasa Consular", "Ademdum", "Prueba pagada", "Pago Inscripcion", "Pago Entrevista",
                   "Aprobacion Inscripcion", "Cuenta convisa", "Fecha emision visa", "Fecha expiracion visa"]

# ── Helpers ──────────────────────────────────────────────────────────────────

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
    return df

def set_time_format(time: pd.Timestamp) -> str:
    if not isinstance(time, pd.Timestamp):
        raise TypeError(f"Se esperaba un Timestamp, se recibió: {type(time)}")
    return time.strftime("%d/%m/%Y")

# ── I/O helpers ──────────────────────────────────────────────────────────────

def read_excel(excel_path: str, skip_rows: int) -> pd.DataFrame:
    "Load data from a excel file into a df"
    path = Path(excel_path)
    
    if not path.exists():
        raise FileNotFoundError(f"Reporte general no econtrado: {excel_path}")
    
    df = pd.read_excel(io=path, skiprows=skip_rows)
    
    return df


def clean_df(df: pd.DataFrame, COLUMNS_TO_DROP_REPORTE_GENERAL: list[str] = None, celular: bool = True) -> pd.DataFrame:
    "It's main goal is to clean a dataframe"

    if not isinstance(df, pd.DataFrame):
        raise TypeError(f"Se esperaba un DataFrame, se recibio un {type(df)}")
    
    if df.empty:
        raise ValueError(f"El DataFrame está vacío, no se puede limpiar")
    
    # Clean
    print("Eliminando columns innecesarias")
    
    if isinstance(COLUMNS_TO_DROP_REPORTE_GENERAL, list):
        df = df.drop(columns=COLUMNS_TO_DROP_REPORTE_GENERAL)

    df = df.dropna(axis=1, how='all')
    df = df.dropna(axis=0, how='all')

    print("Seteando columnas numericas CI y Celular a columnas string, con el 0 por delante")
    df['CI'] = df['CI'].apply(pad_cedula_celular)
    
    if celular:
        df['Celular'] = df['Celular'].apply(pad_cedula_celular)

    print(f"Seteando columnas de tipo datetime64[ns] a strings")
    df = find_datetime_columns_set_to_string(df)
     
    print("Eliminando espacios vacios por delante y detras de columnas string")
    df = find_and_clean_str_columns(df)
    df['Nombre'] = df['Nombre'].str.lower()


    if df.empty:
        raise ValueError("El DataFrame quedó vacío después de la limpieza")
    
    print("Dataframe limpiado con éxito")
    print(SEPARATOR)
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
    # REPORTE GENERAL
    try:
        df = read_excel(PATH, SKIP_ROWS)
        print("Excel reporte cargado")
    except FileNotFoundError as e:
        print(f"[ERROR] al cargar archivo de entrada => {e}")
        return
    except Exception as e:
        print(f"[ERROR] al cargar archivo de entrada => {e}")
        return

    print(SEPARATOR)
    try:
        df = clean_df(df, COLUMNS_TO_DROP_REPORTE_GENERAL)
        print(df)
    except TypeError as e:
        print(f"[ERROR] al limpiar el dataframe => {e}")
        return
    except ValueError as e:
        print(f"[ERROR] al limpiar el dataframe => {e}")
        return
    except Exception as e:
        print(f"[ERROR] al limpiar el dataframe => {e}")
        return

    print(SEPARATOR)
    try:
        save_csv(df, OUTPUT_PATH)
    except TypeError as e:
        print(f"[ERROR] al guardar el reporte excel en formato csv => {e}")
    except ValueError as e:
        print(f"[ERROR] al guardar el reporte excel en formato csv => {e}")
    except Exception as e:
        print(f"[ERROR] al guardar el reporte excel en formato csv => {e}")

    # CONTRATO NO FIRMADO
    try:
        df = read_excel(NO_FIRMADO_PATH, SKIP_ROWS)
        print("Reporte de contratos no firmados cargado")
    except FileNotFoundError as e:
        print(f"[ERROR] al cargar archivo de entrada => {e}")
        return
    except Exception as e:
        print(f"[ERROR] al cargar archivo de entrada => {e}")
        return
    
    print(SEPARATOR)
    
    try:
        df = clean_df(df, COLUMNS_TO_DROP_REPORTE_GENERAL_SIN_CONTRATO)
        print(df)
    except TypeError as e:
        print(f"[ERROR] al limpiar el dataframe => {e}")
        return
    except ValueError as e:
        print(f"[ERROR] al limpiar el dataframe => {e}")
        return
    except Exception as e:
        print(f"[ERROR] al limpiar el dataframe => {e}")
        return

    print(SEPARATOR)

    try:
        save_csv(df, OUTPUT_PATH_NO_FIRMADO)
    except TypeError as e:
        print(f"[ERROR] al guardar el reporte excel en formato csv => {e}")
    except ValueError as e:
        print(f"[ERROR] al guardar el reporte excel en formato csv => {e}")
    except Exception as e:
        print(f"[ERROR] al guardar el reporte excel en formato csv => {e}")


    print(SEPARATOR)

    #PRUEBA INGLES
    try:
        df = read_excel(PRUEBA_INGLES_PATH, skip_rows=1)
        print("Reporte prueba de ingles cargado")
    except FileNotFoundError as e:
        print(f"[ERROR] al cargar archivo de entrada => {e}")
        return
    except Exception as e:
        print(f"[ERROR] al cargar archivo de entrada => {e}")
        return
    
    print(SEPARATOR)
    try:
        df = clean_df(df, celular=False)
        print(df)
    except TypeError as e:
        print(f"[ERROR] al limpiar el dataframe => {e}")
        return
    except ValueError as e:
        print(f"[ERROR] al limpiar el dataframe => {e}")
        return
    except Exception as e:
        print(f"[ERROR] al limpiar el dataframe => {e}")
        return

    print(SEPARATOR)

    try:
        save_csv(df, OUTPUT_PATH_PRUEBAS_INGLES)
    except TypeError as e:
        print(f"[ERROR] al guardar el reporte excel en formato csv => {e}")
    except ValueError as e:
        print(f"[ERROR] al guardar el reporte excel en formato csv => {e}")
    except Exception as e:
        print(f"[ERROR] al guardar el reporte excel en formato csv => {e}")  

if __name__ == "__main__":
    main()
