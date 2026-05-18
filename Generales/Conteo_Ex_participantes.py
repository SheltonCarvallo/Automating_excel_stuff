#!C:\Users\Shelton\anaconda3\envs\DataScienceEnv\python.exe
import pandas as pd
from openpyxl import load_workbook
from pathlib import Path

# ── Constants ────────────────────────────────────────────────────────────────
#INPUT_REPORT    = "NuevosParticipantes.xlsx"
INPUT_REPORT    = "ReporteGeneralSpring2027.xls"
INPUT_CI_FILE   = "CI_spring2026.txt"
OUTPUT_PREFIX   = "EX_participantes"

COLUMNAS = [
    'Sponsor', 'Estatus Participante', 'Universidad', 'Carrera',
    'Nombre', 'CI', 'Celular', 'Correo', 'Empleador', 'Posicion Laboral'
]

PERFILES_FAKE = {"Veronica Alejandra Ollage Velastegí", "CAMILA PRUEBA", "EMILY PRUEBA PRUEBA PRUEBA", "Ashly Prueba"}

SEPARATOR = "=" * 120

# ── Helpers ──────────────────────────────────────────────────────────────────
def pad_cedula(ci) -> str:
    """Ensure CI is always 10 digits by left-padding with a zero if needed."""
    return str(ci).zfill(10)


def autofit_columns(ws) -> None:
    """Set each column width to the length of its longest value."""
    for col in ws.columns:
        max_len = max(
            (len(str(cell.value)) for cell in col if cell.value is not None),
            default=0,
        )
        ws.column_dimensions[col[0].column_letter].width = max_len + 2


def save_excel(df: pd.DataFrame, filepath: str) -> None:
    """Save a DataFrame to Excel and auto-fit all column widths."""
    with pd.ExcelWriter(filepath) as writer:
        df.to_excel(writer, index=False)
    print(f"Excel guardado → {filepath}")

    wb = load_workbook(filepath)
    autofit_columns(wb.active)
    wb.save(filepath)


# ── I/O helpers ──────────────────────────────────────────────────────────────
def load_cedulas(filepath: str) -> set[str]:
    """Load CIs from a plain-text file into a set for O(1) look-ups."""
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"Archivo de CIs no encontrado: {filepath}")

    with open(path, "r", encoding="utf-8") as f:
        cedulas = {line.strip().lower() for line in f if line.strip()}

    print(f"{len(cedulas)} CI spring 2026 leídos")
    return cedulas


def load_reporte(filepath: str, columnas: list[str]) -> pd.DataFrame:
    """Load and clean the participants report."""
    path = Path(filepath.strip())           # .strip() guards against the trailing space in the filename
    if not path.exists():
        raise FileNotFoundError(f"Reporte no encontrado: {filepath}")

    df = pd.read_excel(path, usecols=columnas, skiprows=6)
    df['CI'] = df['CI'].apply(pad_cedula)

    mask_activos        = df['Estatus Participante'] == 'Activo'
    mask_perfiles_reales = ~df['Nombre'].isin(PERFILES_FAKE)
    df = df.loc[mask_activos & mask_perfiles_reales].copy()

    print(f"{len(df)} participantes activos y reales cargados")
    return df


# ── Core logic ───────────────────────────────────────────────────────────────
def classify_participants(
    df: pd.DataFrame,
    cedulas_anteriores: set[str],
) -> tuple[pd.DataFrame, list[str]]:
    """
    Split active participants into:
      - df_ex    : CIs that appear in the previous season
      - ci_nuevos: CIs that do NOT appear in the previous season
    """
    cedulas_actuales = set(df['CI'])                          # O(1) look-ups

    ci_ex_par  = cedulas_actuales & cedulas_anteriores        # set intersection
    ci_nuevos  = cedulas_actuales - cedulas_anteriores        # set difference

    df_ex = (
        df.loc[df['CI'].isin(ci_ex_par), COLUMNAS]
        .sort_values(by='Sponsor')
        .reset_index(drop=True)
    )

    print(f"Total de CI que NO son ex-participantes  => {len(ci_nuevos)}")
    print(f"Ex-participantes encontrados en reporte  => {len(df_ex)}")
    return df_ex, list(ci_nuevos)


def validate_totals(df_total: pd.DataFrame, df_ex: pd.DataFrame, ci_nuevos: list) -> None:
    """Warn if ex + new counts don't match the active participant total."""
    total_conteo = len(df_ex) + len(ci_nuevos)
    if total_conteo == len(df_total):
        print("✔ Sumatoria ex y nuevos participantes cuadra con el reporte general")
    else:
        print(
            f"⚠ No cuadran los totales: ex= {len(df_ex)} + nuevos= {len(ci_nuevos)} "
            f"= {total_conteo}, pero activos en reporte = {len(df_total)}"
        )


# ── Entry point ──────────────────────────────────────────────────────────────
def main() -> None:
    print(SEPARATOR)

    try:
        cedulas_2026 = load_cedulas(INPUT_CI_FILE)
        df           = load_reporte(INPUT_REPORT, COLUMNAS)
    except FileNotFoundError as e:
        print(f"[ERROR] {e}")
        return
    except Exception as e:
        print(f"[ERROR] Al cargar archivos de entrada: {e}")
        return

    print(SEPARATOR)

    try:
        df_ex, ci_nuevos = classify_participants(df, cedulas_2026)
    except KeyError as e:
        print(f"[ERROR] Columna esperada no encontrada: {e}")
        return
    except Exception as e:
        print(f"[ERROR] Al clasificar participantes: {e}")
        return

    validate_totals(df, df_ex, ci_nuevos)

    fecha_corte   = pd.Timestamp.now().normalize().strftime('%d-%m-%Y')
    output_path   = f"{OUTPUT_PREFIX}_{fecha_corte}.xlsx"

    try:
        save_excel(df_ex, output_path)
    except PermissionError:
        print(f"[ERROR] No se pudo guardar '{output_path}'. ¿Está abierto en Excel?")
    except Exception as e:
        print(f"[ERROR] Al guardar el archivo: {e}")


    print(len(ci_nuevos))
    df_test = df.loc[df['CI'].isin(ci_nuevos)]
    print(df_test)
    print(df['CI'].duplicated().sum())
    print(df.loc[df['CI'].duplicated()])
if __name__ == "__main__":
    main()