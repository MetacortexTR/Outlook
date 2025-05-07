import csv
import logging
import re
import sys
import os

logger = logging.getLogger(__name__)

def is_valid_email(email):
    """E-posta adresinin geçerli olup olmadığını kontrol eder"""
    # Exchange/LDAP formatını filtrele
    if '/o=ExchangeLabs/' in email or '/ou=' in email:
        return False
    
    # Basit e-posta doğrulama
    email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(email_pattern, email))

def convert_email_to_name(email):
    """E-posta adresinden ad soyad oluşturur"""
    try:
        # E-posta adresinin @ işaretinden önceki kısmını al
        name_part = email.split('@')[0]
        
        # Nokta veya alt çizgi ile ayrılmış kısımları böl
        parts = re.split('[._]', name_part)
        
        # Her kelimenin ilk harfini büyük yap
        parts = [part.capitalize() for part in parts if part]
        
        # Parçaları birleştir
        full_name = ' '.join(parts)
        
        return full_name
    except:
        return email

def format_name(name):
    """İsmi düzgün formata getirir"""
    try:
        # E-posta adresi ise ada çevir
        if '@' in name and is_valid_email(name):
            return convert_email_to_name(name)
        
        # Noktalı virgül veya virgülle ayrılmış isimleri işle
        names = []
        for part in re.split('[;,]', name):
            part = part.strip()
            # Eğer parça e-posta ise ada çevir
            if '@' in part and is_valid_email(part):
                names.append(convert_email_to_name(part))
            elif part:  # Normal isim
                names.append(part)
        
        return ';'.join(names)
    except Exception as e:
        logger.error(f'İsim formatı hatası: {str(e)}')
        return name

def is_system_info(text):
    """Metnin sistem bilgisi olup olmadığını kontrol eder"""
    if not text:
        return False
        
    system_patterns = [
        'SMTP:', 
        'EX:', 
        ';EX:', 
        'Exchange',
        '/O=EXCHANGELABS/',
        '/OU=',
        'IMCEAEX-',
        'outlook_',
        'SPF',
        'DKIM',
        'DMARC',
        '/CN=',
        'X-MS-Exchange',
        'X-Microsoft',
        'Microsoft Exchange',
        'AutoDiscover',
        '/DC=',
        'smtp.mailfrom'
    ]
    return any(pattern.lower() in text.lower() for pattern in system_patterns)

def clean_email_text(text):
    """E-posta metnini temizler ve düzenler"""
    if not text:
        return text
    
    try:
        # Fazla boşlukları temizle
        text = ' '.join(text.split())
        
        # Prefix'leri temizle
        prefixes = [
            'From:', 'To:', 'Cc:', 'Bcc:', 
            'Kimden:', 'Kime:', 'Bilgi:', 'Gizli:',
            'From :', 'To :', 'Cc :', 'Bcc :',
            'Kimden :', 'Kime :', 'Bilgi :', 'Gizli :',
            'Sender:', 'Recipient:', 'Reply-To:',
            'Gönderen:', 'Alıcı:', 'Yanıtla:',
            'Sender :', 'Recipient :', 'Reply-To :',
            'Gönderen :', 'Alıcı :', 'Yanıtla :'
        ]
        
        # Prefix'leri baştan temizle
        for prefix in prefixes:
            if text.startswith(prefix):
                text = text[len(prefix):].strip()
        
        # Sistem bilgilerini temizle
        system_patterns = [
            r'SMTP:.*?(?=\s|$)',
            r'EX:.*?(?=\s|$)',
            r';EX:.*?(?=\s|$)',
            r'/O=EXCHANGELABS/.*?(?=\s|$)',
            r'/OU=.*?(?=\s|$)',
            r'IMCEAEX-.*?(?=\s|$)',
            r'outlook_.*?(?=\s|$)',
            r'SPF=.*?(?=\s|$)',
            r'DKIM=.*?(?=\s|$)',
            r'DMARC=.*?(?=\s|$)',
            r'/CN=.*?(?=\s|$)',
            r'X-MS-Exchange.*?(?=\s|$)',
            r'X-Microsoft.*?(?=\s|$)',
            r'/DC=.*?(?=\s|$)'
        ]
        
        for pattern in system_patterns:
            text = re.sub(pattern, '', text)
        
        # Ad Soyad <email@domain.com <mailto:email@domain.com>> formatını düzelt
        pattern = r'(.*?)\s*<([^>]+?)\s*<mailto:[^>]+>>'
        while re.search(pattern, text):
            text = re.sub(pattern, r'\1 <\2>', text)
        
        # Kalan mailto: etiketlerini temizle
        text = re.sub(r'<mailto:[^>]+>', '', text)
        
        # Aynı e-postanın tekrarını temizle (örn: email@domain.com <email@domain.com>)
        text = re.sub(r'(\S+@\S+\.\S+)\s*<\1>', r'\1', text)
        
        # Çift parantezli e-postaları temizle
        text = re.sub(r'<([^>]+)\s*<[^>]+>>', r'<\1>', text)
        
        # Gereksiz boşlukları temizle
        text = re.sub(r'\s+', ' ', text)
        text = text.strip()
        
        return text
    except Exception as e:
        logger.error(f'Metin temizleme hatası: {str(e)}')
        return text

def split_names(name_text):
    """Ad Soyad listesini ayırır"""
    if not name_text:
        return []
    
    try:
        # Noktalı virgül ve virgülle ayrılmış isimleri böl
        names = []
        for part in re.split('[;,]', name_text):
            name = part.strip()
            if name:
                names.append(name)
        
        return names
    except Exception as e:
        logger.error(f'İsim ayırma hatası: {str(e)}')
        return [name_text] if name_text else []

def extract_emails_from_text(text):
    """Metinden tüm e-posta adreslerini çıkarır"""
    if not text:
        return []
    
    try:
        emails = []
        # E-posta adreslerini bul
        for email in re.findall(r'[\w\.-]+@[\w\.-]+\.\w+', text):
            if is_valid_email(email):
                emails.append(email)
        
        # Eğer hiç e-posta bulunamadıysa, orijinal metni kontrol et
        if not emails and is_valid_email(text):
            emails.append(text)
        
        return emails
    except Exception as e:
        logger.error(f'E-posta çıkarma hatası: {str(e)}')
        return []

def find_categorized_emails_in_file(file_path):
    try:
        categorized_data = {
            'Kimden': [],
            'Kime': [],
            'Bilgi': [],
            'Gizli': []
        }
        
        with open(file_path, 'r', encoding='utf-8-sig') as file:
            reader = csv.DictReader(file)
            
            # CSV başlıklarını kontrol et
            headers = reader.fieldnames
            logger.debug(f'CSV başlıkları: {headers}')
            
            # Outlook CSV sütun isimleri
            columns = {
                'Kimden': {
                    'email': 'Kimden: (Adres)',
                    'name': 'Kimden: (Ad)'
                },
                'Kime': {
                    'email': 'Kime: (Adres)',
                    'name': 'Kime: (Ad)'
                },
                'Bilgi': {
                    'email': 'Bilgi: (Adres)',
                    'name': 'Bilgi: (Ad)'
                },
                'Gizli': {
                    'email': 'Gizli: (Adres)',
                    'name': 'Gizli: (Ad)'
                }
            }
            
            # Her satır için
            for row_num, row in enumerate(reader, start=1):
                try:
                    # Her kategori için
                    for category, column_info in columns.items():
                        email_column = column_info['email']
                        name_column = column_info['name']
                        
                        # E-posta ve isim bilgilerini al
                        email_text = row.get(email_column, '').strip()
                        name_text = row.get(name_column, '').strip()
                        
                        # İsmi formatla
                        name_text = format_name(name_text)
                        
                        if email_text and not is_system_info(email_text):
                            # E-posta metnini temizle
                            clean_text = clean_email_text(email_text)
                            logger.debug(f'Temizlenmiş metin: {clean_text}')
                            
                            # E-posta adreslerini bul
                            emails = extract_emails_from_text(clean_text)
                            logger.debug(f'Bulunan e-postalar: {emails}')
                            
                            # İsimleri ayır
                            names = split_names(name_text)
                            logger.debug(f'Ayırılan isimler: {names}')
                            
                            # Her e-posta için
                            for i, email in enumerate(emails):
                                # Eğer e-posta sayısı kadar isim varsa, eşleştir
                                # Yoksa mevcut isimleri tekrar kullan veya e-postadan isim oluştur
                                if i < len(names):
                                    name = names[i]
                                elif names:
                                    name = names[0]
                                else:
                                    name = convert_email_to_name(email)
                                
                                email_data = {
                                    'email': email,
                                    'original_text': f"{name} <{email}>",
                                    'name': name,
                                    'row': row_num
                                }
                                
                                # Tekrar kontrolü
                                if not any(d['email'] == email for d in categorized_data[category]):
                                    categorized_data[category].append(email_data)
                except Exception as e:
                    logger.error(f'Satır {row_num} işlenirken hata: {str(e)}')
                    continue
        
        # Boş kategorileri kaldır
        result = {k: v for k, v in categorized_data.items() if v}
        
        if not result:
            logger.warning('Hiç e-posta adresi bulunamadı')
        else:
            logger.info(f'Toplam {sum(len(v) for v in result.values())} e-posta adresi bulundu')
            for category, data in result.items():
                logger.info(f'{category}: {len(data)} e-posta')
        
        return result
        
    except Exception as e:
        logger.error(f'Dosya işleme hatası: {str(e)}')
        return {}

if __name__ == "__main__":
    # Dosya adını komut satırından al veya varsayılan kullan
    file_path = sys.argv[1] if len(sys.argv) > 1 else "outlook111.CSV"
    print(f"Dosya işleniyor: {file_path}")
    
    # Dosyanın varlığını kontrol et
    if not os.path.exists(file_path):
        print(f"HATA: {file_path} dosyası bulunamadı!")
        sys.exit(1)
        
    # Dosya boyutunu kontrol et
    size_mb = os.path.getsize(file_path) / (1024 * 1024)
    print(f"Dosya boyutu: {size_mb:.2f} MB")
    
    # İlk birkaç satırı göster
    print("\nDosyanın ilk birkaç satırı:")
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        for i, line in enumerate(f):
            if i < 5:  # İlk 5 satır
                print(f"Satır {i+1}: {line.strip()}")
            else:
                break
                
    print("\nE-posta adresleri aranıyor...")
    results = find_categorized_emails_in_file(file_path)
    
    if results:
        print("\nBulunan e-postalar:")
        total_emails = 0
        unique_emails = set()
        
        for category, data_list in results.items():
            print(f"\n{category}:")
            for data in data_list:
                print(f"  Satır {data['row']}: {data['original_text']} -> {data['email']} (İsim: {data['name']})")
                unique_emails.add(data['email'])
                total_emails += 1
        
        print(f"\nToplam {total_emails} e-posta adresi bulundu")
        print(f"Bunlardan {len(unique_emails)} tanesi benzersiz")
    else:
        print("Hiç e-posta adresi bulunamadı.") 