import pandas as pd

fichier_source = "etatroleavecprestation.xls"
feuille = 0

# Lecture du fichier
df = pd.read_excel(fichier_source, sheet_name=feuille)

# Nettoyage des colonnes
df.columns = df.columns.str.strip()
df["Prestation"] = df["Prestation"].astype(str).str.strip()

# Conditions de filtrage
condition_prestation = df["Prestation"].str.lower() == "redevance pour prélèvement sur la ressource en eau"
condition_regularisation = df["Prestation"].str.upper().str.contains("REGULARISATION")

# Filtrage
df_filtered = df[condition_prestation | condition_regularisation].copy()

# Gestion des lignes REGULARISATION
for i in df_filtered.index:
    if "REGULARISATION" in df_filtered.loc[i, "Prestation"].upper():
        if i - 1 in df_filtered.index:
            df_filtered.at[i, "N° facture"] = df_filtered.at[i - 1, "N° facture"]
            df_filtered.at[i, "Redevable"] = df_filtered.at[i - 1, "Redevable"]

# Recalcul de l'AVOIR HT à partir de la Quantité
df_filtered["AVOIR HT"] = df_filtered["Quantité"] * 0.0943

# Arrondi à 2 décimales et suppression des négatifs
df_filtered["AVOIR HT"] = df_filtered["AVOIR HT"].round(2).clip(lower=0)

# Calcul de la colonne "avoir final HT" si la colonne "HT" existe
if "HT" in df_filtered.columns:
    df_filtered["avoir final HT"] = df_filtered["AVOIR HT"] - df_filtered["HT"]
    # Remboursement uniquement si différence négative, sinon 0
    df_filtered["avoir final HT"] = df_filtered["avoir final HT"].apply(lambda x: x if x < 0 else 0)

    # Calcul TTC à 5.5% sur "avoir final HT"
    df_filtered["avoir final TTC"] = (df_filtered["avoir final HT"] * 1.055).round(2)
else:
    print("⚠️ La colonne 'HT' n'existe pas dans le fichier source. Impossible de calculer 'avoir final HT'.")
    df_filtered["avoir final HT"] = 0
    df_filtered["avoir final TTC"] = 0

# Colonnes finales
colonnes_voulues = ["N° facture", "Redevable", "AVOIR HT", "avoir final HT", "avoir final TTC"]
df_final = df_filtered[colonnes_voulues].copy()

# Export
df_final.to_excel("factures_filtrees_et_regularisations_sans_negatifs.xlsx", index=False)

print("✅ Export terminé : factures_filtrees_et_regularisations_sans_negatifs.xlsx")
