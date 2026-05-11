import pandas as pd
import numpy as np
from openpyxl import load_workbook
from pathlib import Path

# ── Constants ────────────────────────────────────────────────────────────────
INPUT_REPORT   = "ReporteGeneralSpring2027.xls"
INPUT_CI_FILE  = "CI_spring2026.txt"
OUTPUT_FILE    = "participantes_prueba_ingles_agendar.xlsx"
REPORT_SKIPROWS = 6

COLUMNAS_PROCESO = [
    'Nombre', 'CI', 'Sponsor', 'Empleador', 'Posicion Laboral',
    'Celular', 'Correo', 'Fecha Inscripcion'
]

# ── Helpers ──────────────────────────────────────────────────────────────────
def pad_cedula(ci) -> str:
    """Ensure CI is always 10 digits by left-padding with a zero if needed."""
    ci_str = str(ci)
    return ci_str.zfill(10)          # replaces the manual len-check


def get_fecha_corte() -> tuple[pd.Timestamp, pd.Timestamp]:
    """
    Return (start, end) window for yesterday's inscriptions.
    On Monday, look back 3 days to cover the weekend.
    """
    today = pd.Timestamp.now().normalize()
    is_monday = pd.Timestamp.today().day_name() == "Monday"
    offset = pd.to_timedelta("3 day") if is_monday else pd.to_timedelta("1 day")
    return today - offset, today


def autofit_columns(ws) -> None:
    """Set each column width to the length of its longest value."""
    for col in ws.columns:
        max_len = max(
            (len(str(cell.value)) for cell in col if cell.value is not None),
            default=0
        )
        ws.column_dimensions[col[0].column_letter].width = max_len + 2


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


def load_reporte(filepath: str, skiprows: int = REPORT_SKIPROWS) -> pd.DataFrame:
    """Load the general report Excel file."""
    if not Path(filepath).exists():
        raise FileNotFoundError(f"Reporte no encontrado: {filepath}")

    df = pd.read_excel(filepath, skiprows=skiprows)
    print(f"Número de registros en reporte general 2027 => {len(df)}")
    return df


def save_excel(df: pd.DataFrame, filepath: str) -> None:
    """Save DataFrame to Excel and auto-fit all column widths."""
    with pd.ExcelWriter(filepath) as writer:
        df.to_excel(writer, index=False)   # index=False keeps it clean
    print(f"Excel guardado → {filepath}")

    wb = load_workbook(filepath)
    autofit_columns(wb.active)
    wb.save(filepath)


# ── Core logic ───────────────────────────────────────────────────────────────
def filter_nuevos_participantes(
    df: pd.DataFrame,
    cedulas_anteriores: set[str]
) -> pd.DataFrame:
    """
    Keep only active participants inscribed in the cut-off window
    whose CI was NOT in the previous season.
    """
    fecha_inicio, fecha_fin = get_fecha_corte()
    print(f"Ventana de corte: {fecha_inicio} → {fecha_fin}")

    mask_activo      = df['Estatus Participante'] == "Activo"
    mask_fecha       = (df['Fecha Inscripcion'] >= fecha_inicio) & \
                       (df['Fecha Inscripcion'] <  fecha_fin)

    inscritos = df.loc[mask_activo & mask_fecha, COLUMNAS_PROCESO].copy()
    inscritos['CI'] = inscritos['CI'].apply(pad_cedula)

    print(f"Estudiantes inscritos en el período: {len(inscritos)}")

    mask_nuevos = ~inscritos['CI'].isin(cedulas_anteriores)  # ~ replaces the for-loop
    nuevos = inscritos.loc[mask_nuevos].sort_values(by='Fecha Inscripcion') 

    print(f"Participantes nuevos (no estaban en 2026): {len(nuevos)}")
    return nuevos


# ── Entry point ──────────────────────────────────────────────────────────────
def main() -> None:
    try:
        cedulas_2026 = load_cedulas(INPUT_CI_FILE)
        df_reporte   = load_reporte(INPUT_REPORT)
    except FileNotFoundError as e:
        print(f"[ERROR] {e}")
        return
    except Exception as e:
        print(f"[ERROR] Al cargar los archivos de entrada: {e}")
        return

    try:
        df_nuevos = filter_nuevos_participantes(df_reporte, cedulas_2026)
        print(df_nuevos)
    except KeyError as e:
        print(f"[ERROR] Columna esperada no encontrada en el reporte: {e}")
        return
    except Exception as e:
        print(f"[ERROR] Al filtrar participantes: {e}")
        return

    try:
        save_excel(df_nuevos, OUTPUT_FILE)
    except PermissionError:
        print(f"[ERROR] No se pudo guardar '{OUTPUT_FILE}'. ¿Está abierto en Excel?")
    except Exception as e:
        print(f"[ERROR] Al guardar el archivo de salida: {e}")


if __name__ == "__main__":
    main()