import pandas as pd
import re
import quopri
import codecs

def parse_vcf(vcf_file):
    """Parse VCF file and extract contacts."""
    contacts = []
    
    # Read as binary and try to decode properly
    with open(vcf_file, 'rb') as f:
        raw_content = f.read()
    
    # Try multiple encodings
    for encoding in ['utf-8', 'utf-8-sig', 'windows-1256', 'cp1256', 'iso-8859-6']:
        try:
            content = raw_content.decode(encoding)
            break
        except (UnicodeDecodeError, LookupError):
            continue
    else:
        content = raw_content.decode('utf-8', errors='replace')
    
    # Split by BEGIN:VCARD
    vcards = content.split('BEGIN:VCARD')[1:]
    
    for vcard in vcards:
        name = ''
        phone = ''
        
        # Extract name (FN or N field, handle quoted-printable)
        fn_match = re.search(r'FN[^:]*:(.*?)(?:\r?\n)', vcard, re.MULTILINE | re.DOTALL)
        if fn_match:
            name = fn_match.group(1).strip()
            # Handle quoted-printable encoding
            if '=' in name and re.search(r'=[0-9A-F]{2}', name):
                try:
                    name = quopri.decodestring(name.encode('latin1')).decode('utf-8')
                except:
                    pass
        else:
            n_match = re.search(r'N[^:]*:(.*?)(?:\r?\n)', vcard)
            if n_match:
                name = n_match.group(1).replace(';', ' ').strip()
                if '=' in name and re.search(r'=[0-9A-F]{2}', name):
                    try:
                        name = quopri.decodestring(name.encode('latin1')).decode('utf-8')
                    except:
                        pass
        
        # Extract phone number (first TEL field)
        tel_match = re.search(r'TEL[^:]*:(.*?)[\r\n]', vcard)
        if tel_match:
            phone = tel_match.group(1).strip()
            # Remove all country codes and non-digits
            phone = re.sub(r'\D', '', phone)
            # Remove common country codes from start
            phone = re.sub(r'^(962|1)', '', phone)
        
        # Skip if name is just a phone number or empty
        if name and phone and not name.replace('+', '').replace('0', '').isdigit():
            contacts.append({'Name': name, 'Nickname': '', 'Number': phone})
    
    return contacts

def vcf_to_excel(vcf_file, excel_file='contacts.xlsx'):
    """Convert VCF file to Excel format for whatsapp-bot."""
    contacts = parse_vcf(vcf_file)
    df = pd.DataFrame(contacts)
    df.to_excel(excel_file, index=False)
    print(f"Created {excel_file} with {len(contacts)} contacts")
    print("Please fill in the 'Nickname' column manually")

if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python vcf_to_excel.py <vcf_file> [output_excel]")
        print("Example: python vcf_to_excel.py contacts.vcf")
        sys.exit(1)
    
    vcf_file = sys.argv[1]
    excel_file = sys.argv[2] if len(sys.argv) > 2 else 'contacts.xlsx'
    
    vcf_to_excel(vcf_file, excel_file)
