def TB2K(df_cb):
        def ambil_target():
            file_path = os.path.join(os.getcwd(), "R09 PRIANGAN TIMUR.xlsm")
            try:
                df_target = pd.read_excel(
                    file_path,
                    sheet_name="SALESMAN LIST",
                    usecols=["KD_SLS", "TBK-E02K"]
                )
                df_target = df_target.rename(columns={"TBK-E02K": "TVALUE"})
                df_target["TVALUE"] = df_target["TVALUE"].fillna(0).astype(int)
                return df_target
            except Exception as e:
                print(f"❌ Gagal membaca file TARGET: {e}")
                return pd.DataFrame(columns=["KD_SLS", "TVALUE"])
        dfTB2K = dbase.KSNI()
        dfTB2K = dfTB2K[dfTB2K["NAMA SLS2"].str.startswith(startswith)].reset_index(drop=True)
        fdftb2k = [300196, 300812, 301924, 301938]
        dfTB2K = dfTB2K[dfTB2K["KD_BRG"].isin(fdfTB2K)]
        dfTB2K["TB2K"] = 1
        df_sum = dfTB2K.groupby("KD SLS2")[["TB2K", "VALUE"]].sum().reset_index()
        df_sum = df_sum.rename(columns={"KD SLS2": "KD_SLS"})
        df_merged = df_cb.merge(df_sum, on="KD_SLS", how="left")
        df_merged["TB2K"] = df_merged["TB2K"].fillna(0).astype(int)
        df_merged["VALUE"] = df_merged["VALUE"].fillna(0).astype(int)
        df_merged["TTB2K"] = 50
        df_target = ambil_target()
        df_merged = df_merged.merge(df_target, on="KD_SLS", how="left")
        df_merged["TVALUE"] = df_merged["TVALUE"].fillna(0).astype(int)
        df_merged["STB2K"] = ((df_merged["TB2K"] >= df_merged["TTB2K"]) &
                            (df_merged["VALUE"] >= df_merged["TVALUE"])).astype(int)
        df_merged["VALUE"] = df_merged["VALUE"].apply(lambda x: f"{x:,}")
        df_merged["TVALUE"] = df_merged["TVALUE"].apply(lambda x: f"{x:,}")
        kol_tambahan = ["TTB2K", "TB2K", "TVALUE", "VALUE", "STB2K"]
        return df_merged[kol_tambahan]