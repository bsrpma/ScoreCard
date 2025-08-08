import configparser
import pandas as pd
import os

# === Load konfigurasi sekali saja ===
config = configparser.ConfigParser()
config.read("config.txt")

def read_dbase(section):
    """Fungsi umum untuk baca file parquet dari section tertentu"""
    try:
        db_name = config[section]["cb"]
        db_path = config[section]["location"]
        full_path = os.path.join(db_path, db_name)
        return pd.read_parquet(full_path)
    except Exception as e:
        print(f"[ERROR] Gagal membaca data {section}:", e)
        return None

# ===== Fungsi khusus untuk CB =====
def CB():
    df_cb = read_dbase("CB")
    if df_cb is not None:
        pass
        #print(f"✅ Data CB berhasil dibaca, jumlah baris: {len(df_cb)}")
    else:
        print("❌ Gagal membaca data CB")
    return df_cb
