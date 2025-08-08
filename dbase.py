import configparser
import pandas as pd
import os

# === Load konfigurasi sekali saja ===
config = configparser.ConfigParser()
config.read("config.txt")

def read_dbase(section):
    """Fungsi umum untuk baca file parquet dari section tertentu"""
    try:
        db_name = config[section]["dbase"]
        db_path = config[section]["location"]
        full_path = os.path.join(db_path, db_name)
        # print(f"[INFO] Membaca {section} dari {full_path}")
        return pd.read_parquet(full_path)
    except Exception as e:
        print(f"[ERROR] Gagal membaca data {section}:", e)
        return None

# ===== Fungsi khusus per project (sudah dibersihkan / siap pakai) =====
def KSNI():
    df = read_dbase("KSNI")
    if df is not None:
        df = df[[
            "KD SLS2", "NAMA SLS2", "KODE OUTLET", "NAMA OUTLET", "KD_BRG", "NM_BRG", "QTY", "VALUE"
        ]]
        df = df[df["NAMA SLS2"].str.startswith(("AEP", "AEG", "TX2"))]
    return df


def MEIJI():
    df = read_dbase("MEIJI")
    if df is not None:
        df = df[[
            "KD SLS2", "NAMA SLS2", "KODE OUTLET", "NAMA OUTLET", "KD_BRG", "NM_BRG", "QTY", "VALUE"
        ]]
        df = df[df["NAMA SLS2"].str.startswith("AEP")]
    return df


def SIMBA():
    df = read_dbase("SIMBA")
    if df is not None:
        df = df[[
            "KD SLS2", "NAMA SLS2", "KODE OUTLET", "NAMA OUTLET", "KD_BRG", "NM_BRG", "QTY", "VALUE"
        ]]
        df = df[df["NAMA SLS2"].str.startswith("AEP")]
    return df
