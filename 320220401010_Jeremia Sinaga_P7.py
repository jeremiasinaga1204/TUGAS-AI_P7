import numpy as np
import skfuzzy as fuzz
from skfuzzy import control as ctrl
import pandas as pd
import locale
import xlsxwriter

# Set locale untuk Indonesia
locale.setlocale(locale.LC_ALL, 'id_ID')

# Membaca data dari file Excel
data = pd.read_excel('data_perjalanan.xlsx')

# Input variabel
jarak_tempuh = ctrl.Antecedent(np.arange(0, 201, 1), 'jarak_tempuh')
konsumsi_bahan_bakar = ctrl.Antecedent(np.arange(0, 21, 1), 'konsumsi_bahan_bakar')
biaya_perjalanan = ctrl.Consequent(np.arange(0, 1000001, 1), 'biaya_perjalanan')

# Fungsi keanggotaan untuk variabel jarak tempuh
jarak_tempuh['dekat'] = fuzz.trimf(jarak_tempuh.universe, [0, 50, 100])
jarak_tempuh['sedang'] = fuzz.trimf(jarak_tempuh.universe, [75, 125, 175])
jarak_tempuh['jauh'] = fuzz.trimf(jarak_tempuh.universe, [150, 200, 200])

# Fungsi keanggotaan untuk variabel konsumsi bahan bakar
konsumsi_bahan_bakar['rendah'] = fuzz.trimf(konsumsi_bahan_bakar.universe, [0, 5, 10])
konsumsi_bahan_bakar['sedang'] = fuzz.trimf(konsumsi_bahan_bakar.universe, [8, 12, 16])
konsumsi_bahan_bakar['tinggi'] = fuzz.trimf(konsumsi_bahan_bakar.universe, [14, 20, 20])

# Fungsi keanggotaan untuk variabel biaya perjalanan
biaya_perjalanan['murah'] = fuzz.trimf(biaya_perjalanan.universe, [0, 250000, 500000])
biaya_perjalanan['sedang'] = fuzz.trimf(biaya_perjalanan.universe, [400000, 600000, 800000])
biaya_perjalanan['mahal'] = fuzz.trimf(biaya_perjalanan.universe, [700000, 850000, 1000000])

# Rules
rule1 = ctrl.Rule(jarak_tempuh['dekat'] & konsumsi_bahan_bakar['rendah'], biaya_perjalanan['murah'])
rule2 = ctrl.Rule(jarak_tempuh['dekat'] & konsumsi_bahan_bakar['sedang'], biaya_perjalanan['sedang'])
rule3 = ctrl.Rule(jarak_tempuh['dekat'] & konsumsi_bahan_bakar['tinggi'], biaya_perjalanan['mahal'])
rule4 = ctrl.Rule(jarak_tempuh['sedang'] & konsumsi_bahan_bakar['rendah'], biaya_perjalanan['murah'])
rule5 = ctrl.Rule(jarak_tempuh['sedang'] & konsumsi_bahan_bakar['sedang'], biaya_perjalanan['sedang'])
rule6 = ctrl.Rule(jarak_tempuh['sedang'] & konsumsi_bahan_bakar['tinggi'], biaya_perjalanan['mahal'])
rule7 = ctrl.Rule(jarak_tempuh['jauh'] & konsumsi_bahan_bakar['rendah'], biaya_perjalanan['sedang'])
rule8 = ctrl.Rule(jarak_tempuh['jauh'] & konsumsi_bahan_bakar['sedang'], biaya_perjalanan['mahal'])
rule9 = ctrl.Rule(jarak_tempuh['jauh'] & konsumsi_bahan_bakar['tinggi'], biaya_perjalanan['mahal'])

# Control System
biaya_perjalanan_ctrl = ctrl.ControlSystem([rule1, rule2, rule3, rule4, rule5, rule6, rule7, rule8, rule9])
biaya_perjalanan_sim = ctrl.ControlSystemSimulation(biaya_perjalanan_ctrl)

# Analisis data dari file Excel
hasil_perjalanan = []  # List untuk menyimpan hasil estimasi biaya perjalanan

for idx, row in data.iterrows():
    # Memasukkan nilai jarak tempuh dan konsumsi bahan bakar
    jarak_tempuh_value = row['Jarak Tempuh']
    konsumsi_bahan_bakar_value = row['Konsumsi Bahan Bakar']
    biaya_perjalanan_sim.input['jarak_tempuh'] = jarak_tempuh_value
    biaya_perjalanan_sim.input['konsumsi_bahan_bakar'] = konsumsi_bahan_bakar_value

    # Evaluasi sistem
    biaya_perjalanan_sim.compute()

    # Menyimpan hasil estimasi biaya perjalanan dalam format Rupiah dengan pemisah ribuan
    total_biaya_perjalanan = biaya_perjalanan_sim.output['biaya_perjalanan']
    hasil_perjalanan.append(total_biaya_perjalanan)
    print("Data ke-", idx + 1)
    print("Estimasi biaya perjalanan adalah: Rp", locale.format_string("%d", total_biaya_perjalanan, grouping=True))

# Menambahkan kolom hasil estimasi biaya perjalanan ke data awal
data['Estimasi Biaya Perjalanan (Rp)'] = hasil_perjalanan

# Menyimpan data dengan hasil estimasi biaya perjalanan ke dalam file Excel baru
data.to_excel('hasil_perjalanan.xlsx', index=False)
