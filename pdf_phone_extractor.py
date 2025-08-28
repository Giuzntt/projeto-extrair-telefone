import PyPDF2
import re
import os
import pandas as pd
import phonenumbers
from phonenumbers import NumberParseException, is_valid_number, format_number, PhoneNumberFormat
from datetime import datetime


class PDFPhoneExtractor:
    def __init__(self, debug_mode=False):
        self.debug_mode = debug_mode
        self.missed_patterns = []

        self.valid_ddds = {
            '11','12','13','14','15','16','17','18','19','21','22','24','27','28',
            '31','32','33','34','35','37','38','41','42','43','44','45','46','47',
            '48','49','51','53','54','55','61','62','63','64','65','66','67','68',
            '69','71','73','74','75','77','79','81','82','83','84','85','86','87',
            '88','89','91','92','93','94','95','96','97','98','99'
        }

    def normalize_phone(self, raw_number):
        try:
            cleaned_number = re.sub(r'\s+', '', raw_number)
            if not cleaned_number.startswith("+") and len(cleaned_number) in [10, 11]:
                cleaned_number = "+55" + cleaned_number

            parsed = phonenumbers.parse(cleaned_number, "BR")
            if is_valid_number(parsed):
                formatted = format_number(parsed, PhoneNumberFormat.NATIONAL)
                ddd = str(parsed.national_number)[:2]
                return formatted, ddd
            else:
                return None, None
        except NumberParseException:
            return None, None

    def extract_phones_from_pdf(self, pdf_path):
        phones_data = {}
        try:
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                for page_num, page in enumerate(reader.pages, 1):
                    text = page.extract_text()
                    if text:
                        phone_patterns = [
                            r'(?i)(?:tel|cel|whatsapp|telefone)[.:]?\s*\(?\d{2}\)?\s?9?\d{4}[- .]?\d{4}',
                            r'\+55\s?\(?\d{2}\)?\s?9?\d{4}[- .]?\d{4}',
                            r'\(?\d{2}\)?\s?9?\d{4}[- .]?\d{4}',
                            r'\d{2}\s?9?\d{4}[- .]?\d{4}',
                            r'\+55\d{10,11}',
                            r'\(?\d{2}\)?[ .\-/]?\d{4}[ .\-/]?\d{4}',
                            r'\(?\d{2}\)?[ .\-/]?9\d{4}[ .\-/]?\d{4}',
                            r'\(?\d{2}\)?\s?9?\d{4}/\d{4}',
                            r'\+55\.\d{2}\.\d{4}\.\d{4}',
                            r'\+55\.\d{2}\.9\d{4}\.\d{4}',
                            r'\d{10,11}',
                        ]

                        for pattern in phone_patterns:
                            matches = re.findall(pattern, text)
                            for match in matches:
                                match = re.sub(r'(^|\s)[0-9]{1,3}\s*(?=\()', '', match)
                                clean_number = re.sub(r'[^\d+]', '', match)

                                # Ignora padrões que são datas
                                if re.fullmatch(r'20\d{2}(-20\d{2})?', clean_number):
                                    if self.debug_mode:
                                        self.missed_patterns.append({
                                            'pdf': os.path.basename(pdf_path),
                                            'page': page_num,
                                            'pattern': match,
                                            'cleaned': clean_number,
                                            'reason': 'Ignorado por parecer ano ou intervalo de anos'
                                        })
                                    continue

                                if len(clean_number) < 10 or len(clean_number) > 13:
                                    possible = re.search(r'(\(?\d{2}\)?\s?\d{4,5}[- ]?\d{4})$', match)
                                    if possible:
                                        clean_number2 = re.sub(r'[^\d+]', '', possible.group(1))
                                        if len(clean_number2) >= 10 and len(clean_number2) <= 13:
                                            clean_number = clean_number2
                                        else:
                                            if self.debug_mode:
                                                self.missed_patterns.append({
                                                    'pdf': os.path.basename(pdf_path),
                                                    'page': page_num,
                                                    'pattern': match,
                                                    'cleaned': clean_number,
                                                    'reason': 'Ignorado por tamanho inválido'
                                                })
                                            continue
                                    else:
                                        if self.debug_mode:
                                            self.missed_patterns.append({
                                                'pdf': os.path.basename(pdf_path),
                                                'page': page_num,
                                                'pattern': match,
                                                'cleaned': clean_number,
                                                'reason': 'Ignorado por tamanho inválido'
                                            })
                                        continue

                                # DDD
                                ddd_candidate = None
                                if clean_number.startswith("55") and len(clean_number) >= 12:
                                    ddd_candidate = clean_number[2:4]
                                elif clean_number.startswith("+55") and len(clean_number) >= 13:
                                    ddd_candidate = clean_number[3:5]
                                else:
                                    ddd_candidate = clean_number[:2]

                                if ddd_candidate not in self.valid_ddds:
                                    if self.debug_mode:
                                        self.missed_patterns.append({
                                            'pdf': os.path.basename(pdf_path),
                                            'page': page_num,
                                            'pattern': match,
                                            'cleaned': clean_number,
                                            'reason': f'DDD inválido ({ddd_candidate})'
                                        })
                                    continue

                                formatted, ddd = self.normalize_phone(clean_number)
                                if formatted:
                                    if formatted not in phones_data:
                                        phones_data[formatted] = {
                                            'ddd': ddd,
                                            'pages': set([page_num]),
                                            'original': match
                                        }
                                    else:
                                        phones_data[formatted]['pages'].add(page_num)
                                else:
                                    if self.debug_mode:
                                        self.missed_patterns.append({
                                            'pdf': os.path.basename(pdf_path),
                                            'page': page_num,
                                            'pattern': match,
                                            'cleaned': clean_number,
                                            'reason': 'Número inválido pelo libphonenumber'
                                        })
        except Exception as e:
            print(f"❌ Erro ao processar {pdf_path}: {str(e)}")
        return phones_data

    def process_folder(self, folder_path):
        all_phones = {}
        pdf_files = []

        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(root, file))

        print(f"📄 Encontrados {len(pdf_files)} arquivos PDF")

        for i, pdf_file in enumerate(pdf_files, 1):
            print(f"🔍 Processando {i}/{len(pdf_files)}: {os.path.basename(pdf_file)}")
            phones = self.extract_phones_from_pdf(pdf_file)
            for phone, data in phones.items():
                source = os.path.basename(pdf_file)
                pages = ", ".join([f"Página {p}" for p in sorted(data['pages'])])
                if phone not in all_phones:
                    all_phones[phone] = {
                        'ddd': data['ddd'],
                        'sources': set([(source, pages)])
                    }
                else:
                    all_phones[phone]['sources'].add((source, pages))
        return all_phones


def save_to_excel_and_csv(results, excel_path, csv_path):
    try:
        data = []
        for phone, phone_data in results.items():
            for source, pages in phone_data['sources']:
                data.append({
                    'Telefone': phone,
                    'DDD': phone_data['ddd'],
                    'Arquivo de Origem': source,
                    'Páginas': pages,
                    'Data Extração': datetime.now().strftime('%d/%m/%Y %H:%M')
                })

        df = pd.DataFrame(data).sort_values(by=['DDD', 'Telefone'])

        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Telefones', index=False)
            worksheet = writer.sheets['Telefones']
            worksheet.auto_filter.ref = worksheet.dimensions
            col_widths = {'A': 20, 'B': 8, 'C': 30, 'D': 15, 'E': 20}
            for col, width in col_widths.items():
                worksheet.column_dimensions[col].width = width
        print(f"💾 Excel salvo: {excel_path}")

        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        print(f"💾 CSV salvo: {csv_path}")

        return True
    except Exception as e:
        print(f"❌ Erro ao salvar arquivos: {e}")
        return False


def save_failed_attempts(extractor, output_path):
    if extractor.debug_mode and extractor.missed_patterns:
        try:
            df = pd.DataFrame(extractor.missed_patterns)
            df.to_csv(output_path, index=False, encoding='utf-8-sig')
            print(f"\n📝 Telefones ignorados salvos em: {output_path}")
        except Exception as e:
            print(f"❌ Erro ao salvar CSV de erros: {e}")


def main():
    print("=== 📞 EXTRAITOR DE TELEFONES DE PDFs ===")
    print("Versão: 6.3 - Novas Regex + Salvamento\n")

    folder_path = os.getcwd()
    output_dir = os.path.join(folder_path, "output")
    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_output = os.path.join(output_dir, f"telefones_{timestamp}.xlsx")
    csv_output = os.path.join(output_dir, f"telefones_{timestamp}.csv")
    erros_output = os.path.join(output_dir, f"telefones_descartados_{timestamp}.csv")

    extractor = PDFPhoneExtractor(debug_mode=True)

    print(f"\n🚀 Iniciando processamento em: {folder_path}")
    results = extractor.process_folder(folder_path)

    if results:
        save_to_excel_and_csv(results, excel_output, csv_output)
        print(f"\n✅ Telefones encontrados: {len(results)}")
    else:
        print("❌ Nenhum telefone encontrado nos PDFs")

    if extractor.debug_mode:
        if extractor.missed_patterns:
            print("\n⚠️ Telefones ignorados durante a extração:")
            for i, item in enumerate(extractor.missed_patterns, 1):
                print(f"{i:02d}. [{item['pdf']}] Página {item['page']} - '{item['pattern']}' → {item['reason']}")
        else:
            print("\n✅ Nenhum telefone foi ignorado.")
        save_failed_attempts(extractor, erros_output)


if __name__ == "__main__":
    main()
