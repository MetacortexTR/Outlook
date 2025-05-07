import re

def clean_email_text(text):
    # HTML mailto: linklerini temizle
    text = re.sub(r'<mailto:[^>]+>', '', text)
    # İsim kısmını temizle (örn: "Kübra Binzat <kubra.binzat@yzf.com.tr>" -> "kubra.binzat@yzf.com.tr")
    text = re.sub(r'[^<]*<([^>]+)>', r'\1', text)
    return text

def get_names_from_text(text):
    # "Ad Soyad <email@domain.com>" formatından isimleri çıkar
    names = re.findall(r'([^<>]+)\s*<[^>]+>', text)
    return [name.strip() for name in names]

def analyze_csv():
    try:
        with open('outlook (1).CSV', 'r', encoding='utf-8-sig') as file:
            content = file.read()
            
            # E-posta adreslerini bul
            email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
            
            # Tüm e-posta gruplarını bul
            email_groups = re.finditer(r'(?:Kime:|Bilgi:|Gizli:)\s*([^"]+?)(?=(?:Kime:|Bilgi:|Gizli:)|$)', content)
            
            for match in email_groups:
                group_text = match.group(1)
                emails = re.findall(email_pattern, group_text)
                
                if len(emails) > 1:
                    print("\nBirden fazla alıcısı olan grup bulundu:")
                    print("-" * 80)
                    
                    # İsimleri çıkar
                    names = get_names_from_text(group_text)
                    
                    print(f"Alıcılar ({len(emails)} kişi):")
                    for i, (email, name) in enumerate(zip(emails, names + [''] * len(emails)), 1):
                        name_str = f" ({name})" if name else ""
                        print(f"  {i}. {email}{name_str}")
                    
                    print(f"\nOrijinal metin:")
                    print(f"  {group_text.strip()}")
                    print("-" * 80)
                    break  # İlk örneği bulduktan sonra dur
    
    except Exception as e:
        print(f"Hata: {str(e)}")
        import traceback
        print(traceback.format_exc())

if __name__ == "__main__":
    analyze_csv() 