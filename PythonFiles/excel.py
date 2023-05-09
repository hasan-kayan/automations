import os
import pandas as pd

# Klasör yolunu belirleyin
klasor_yolu = r'C:\Users\hasan\Desktop\year_excel\year'

# Tüm Excel dosyalarını seçin ve bir listeye aktarın
excel_dosyalari = [dosya for dosya in os.listdir(klasor_yolu) if dosya.endswith('.xlsx')]

# Excel dosyalarını birleştirin
birlesik_veri = pd.DataFrame()
for dosya in excel_dosyalari:
    dosya_yolu = os.path.join(klasor_yolu, dosya)
    veri = pd.read_excel(dosya_yolu)
    birlesik_veri = pd.concat([birlesik_veri, veri], ignore_index=True)

# Birleştirilmiş veriyi yeni bir Excel dosyasına kaydedin
hedef_dosya = os.path.join(klasor_yolu, 'tamami.xlsx')
birlesik_veri.to_excel(hedef_dosya, index=False)
