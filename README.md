# Traitement de factures Excel – Filtrage & Calculs

Ce script Python traite un fichier Excel de facturation, filtre certaines prestations, et génère un nouveau fichier enrichi avec des colonnes calculées.

## 🔧 Fonctionnalités

- Lecture et nettoyage du fichier Excel source  
- Filtrage des lignes contenant :  
  - la prestation "Redevance pour prélèvement sur la ressource en eau"  
  - ou une mention de "REGULARISATION"  
- Rattachement des lignes "REGULARISATION" à la facture précédente  
- Calcul de :  
  - `AVOIR HT` basé sur la quantité (tarif unitaire : 0,0943 €)  
  - `avoir final HT` : différence entre `AVOIR HT` et `HT` si négative  
  - `avoir final TTC` : ajout de 5,5 % de TVA sur `avoir final HT`  
- Export du résultat dans un fichier :  
  `factures_filtrees_et_regularisations_sans_negatifs.xlsx`

## ▶️ Utilisation

1. Place le fichier `etatroleavecprestation.xls` dans le dossier du script.  
2. Exécute le script avec Python :  
   ```bash
   python script.py
3. Le fichier final sera généré dans le même dossier.

## 🧩 Dépendances

Assure-toi d’avoir pandas et openpyxl installés :
```bash
pip install pandas openpyxl
```
## 📂 Exemple de colonnes en sortie
| N° facture | Redevable | AVOIR HT | avoir final HT | avoir final TTC |
| ---------- | --------- | -------- | -------------- | --------------- |
| F12345     | Dupont    | 14.15    | -3.20          | -3.38           |

## 📝 Auteurs
Script développé par [Hamza Meneceur](https://github.com/HamzaMeneceur)
