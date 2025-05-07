import pandas as pd
from find_emails import find_emails_in_file

def export_emails_to_excel(emails, output_file):
    try:
        # E-posta adreslerini DataFrame'e dönüştür
        df = pd.DataFrame(emails, columns=['E-posta Adresi'])
        
        # Excel dosyasına kaydet
        df.to_excel(output_file, index=False, sheet_name='E-posta Listesi')
        print(f"E-posta adresleri başarıyla '{output_file}' dosyasına kaydedildi.")
    except Exception as e:
        print(f"Hata oluştu: {str(e)}")

if __name__ == "__main__":
    # CSV dosyasından e-postaları al
    csv_file = "outlook (1).CSV"
    emails = find_emails_in_file(csv_file)
    
    if emails:
        # Excel dosyasının yolu
        excel_output = r"C:\Users\sezer.sevgin\Desktop\outlook\Kişi listesi.xlsx"
        
        # E-postaları Excel'e aktar
        export_emails_to_excel(sorted(emails), excel_output)  # Sıralı liste için sorted kullanıyoruz
        
        print(f"Toplam {len(emails)} e-posta adresi aktarıldı.")
    else:
        print("Aktarılacak e-posta adresi bulunamadı.") 