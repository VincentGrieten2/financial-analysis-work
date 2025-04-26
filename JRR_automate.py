import pdfplumber
import re
import os
from pathlib import Path
from decimal import Decimal
from collections import defaultdict
from openpyxl import load_workbook
from datetime import datetime, timedelta
import logging
import shutil
import openpyxl
from openpyxl.utils import get_column_letter
import sys

# Configureer logging
handlers = [logging.FileHandler('parser.log')]
# Only add StreamHandler if not running as executable
if not getattr(sys, 'frozen', False):
    handlers.append(logging.StreamHandler())

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=handlers
)

# ==============================================
# CATEGORIE DEFINITIES
# ==============================================

# RESULTATENREKENING CATEGORIEËN
resultaten_cat = [
    # Bedrijfsopbrengsten
    ("A", "70", "Omzet"),
    ("B", "71", "Voorraadwijzigingen (+) (-)"),
    ("C", "72", "Geproduceerde vaste activa"),
    ("D", "74", "Andere bedrijfsopbrengsten"),
    
    # Bedrijfskosten
    ("E", "60", "Handelsg., grond & hulpst."),
    ("F", "609", "Wijzigingen in de voorraad"),
    ("G", "61", "Diensten en diverse goed."),
    ("H", "62", "Bezoldig., soc. lasten en pensioen."),
    ("I", "630", "Afschrijvingen en waardevermi."),
    ("J", "631/4", "Waardevermind. op voorraden, b.i.u."),
    ("K", "635/7", "Voorz.vr risico's en kosten"),
    ("L", "640/8", "Andere bedrijfskosten"),
    ("M", "649", "Geactiveerde bedrijfskosten (-)"),
    
    # Financiële opbrengsten
    ("N", "750", "Opbrengsten uit fin. vaste activa"),
    ("O", "751", "Opbrengsten uit vlottende activa"),
    ("P", "752/9", "Andere financiële opbrengsten"),
    
    # Financiële kosten
    ("Q", "650", "Kosten van schulden"),
    ("R", "651", "Waardeverminderingen op vlottende act."),
    ("S", "652/9", "Andere financiële kosten"),
    
    # Uitzonderlijke opbrengsten
    ("T", "760", "Terugnem. afschr.en waardevermind. (im)mat.v.a."),
    ("U", "761", "Terugneming waardeverm.op fin.v.a."),
    ("V", "762", "Terugnem. voorz.vr uitz. kosten"),
    ("W", "763", "Meerwaarde bij realis. vaste act."),
    ("X", "764/9", "And. uitzonderlijke opbrengsten"),
    
    # Uitzonderlijke kosten
    ("Y", "660", "Uitz. afschrijvingen"),
    ("Z", "661", "Waardeverminder. op fin. vaste act."),
    ("AA", "662", "Voorz. vr uitz. risico's en kosten"),
    ("AB", "663", "Minderw. bij realisatie vaste act."),
    ("AC", "664/8", "Andere uitzonderlijke kosten"),
    ("AD", "669", "Geactiveerde uitzonderlijke kosten (-)"),
    
    # Belastingen
    ("AE", "670/3", "Belastingen (-)"),
    ("AF", "77", "Regulariser. vn belast. & terugnemingen vn voorziening. voor belast."),
]

# BALANS CATEGORIEËN
balans_cat = [
    # VASTE ACTIVA
    ("BA", "21", "Immateriële vaste activa"),
    ("BB", "22", "Terreinen en gebouwen"),
    ("BC", "23", "Installaties, machines en uitrust."),
    ("BD", "24", "Meubilair en rollend mat."),
    ("BE", "25", "Leasing en soortgelijke rechten"),
    ("BF", "26", "Overige materiele vaste activa"),
    ("BG", "27", "Activa in aanbouw en vooruitbetal."),
    
    # FINANCIËLE VASTE ACTIVA
    ("BH", "280", "Verbonden ondernemingen - Deelnemingen"),
    ("BI", "281", "Verbonden ondernemingen - Vorderingen"),
    ("BJ", "282", "Ondernemingen met deelnemingsverh. - Deelnemingen"),
    ("BK", "283", "Ondernemingen met deelnemingsverh. - Vorderingen"),
    ("BL", "284", "Andere financiele vaste activa - Aandelen"),
    ("BM", "285/8", "Andere financiele vaste activa - Vorder. & borgtochten"),
    
    # VLOTTENDE ACTIVA
    ("BN", "290", "Handelsvorderingen"),
    ("BO", "291", "Overige vorderingen"),
    
    # VOORRADEN
    ("BP", "30/31", "Grond- en hulpstoffen"),
    ("BQ", "32", "Goederen in bewerking"),
    ("BR", "33", "Gereed produkt"),
    ("BS", "34", "Handelsgoederen"),
    ("BT", "35", "Onroer. goed. bestemd vr. verkoop"),
    ("BU", "36", "Vooruitbetalingen"),
    ("BV", "37", "Bestellingen in uitvoering"),
    
    # OVERIGE VLOTTENDE ACTIVA
    ("BW", "40", "Handelsvorderingen"),
    ("BX", "41", "Overige vorderingen"),
    ("BY", "50", "Eigen aandelen"),
    ("BZ", "51/53", "Overige beleggingen"),
    ("CA", "54/58", "Liquide middelen"),
    ("CB", "490/1", "Overlopende rekeningen"),
    
    # EIGEN VERMOGEN
    ("CC", "100", "Geplaatst kapitaal"),
    ("CD", "101", "Niet opgevraagd kapitaal (-)"),
    ("CE", "11", "Uitgiftepremies"),
    ("CF", "12", "Herwaarderingsmeerwaarden"),
    
    # RESERVES
    ("CG", "130", "Wettelijke reserve"),
    ("CH", "131", "Onbeschikbare reserves"),
    ("CI", "132", "Belastingvrije reserves"),
    ("CJ", "133", "Beschikbare reserves"),
    ("CK", "140", "Overgedragen winst"),
    ("CL", "141", "Overgedragen verlies (-)"),
    ("CM", "15", "Kapitaalsubsidies"),
    
    # VOORZIENINGEN EN UITGESTELDE BELASTINGEN
    ("CN", "160", "Pensioen. & soortgel. verplichting"),
    ("CO", "161", "Belastingen"),
    ("CP", "162", "Grote herstel- en onderhoudswerken"),
    ("CQ", "163/9", "Overige risico's & kosten"),
    ("CR", "168", "Uitgestelde belastingen"),
    
    # VREEMD VERMOGEN LANG
    ("CS", "170", "Achtergestelde leningen"),
    ("CT", "171", "Niet achtergestelde obl. leningen"),
    ("CU", "172", "Leasingschulden"),
    ("CV", "173", "Kredietinstellingen"),
    ("CW", "174", "Overige leningen"),
    
    # HANDELSSCHULDEN
    ("CX", "1750", "Handelsschulden - Leveranciers"),
    ("CY", "1751", "Handelsschulden - Te betalen wissels"),
    ("CZ", "176", "Ontvangen vooruitbetal. op bestel."),
    ("DA", "178/9", "Overige schulden"),
    
    # VREEMD VERMOGEN KORT
    ("DB", "42", "Schuld + 1 jr,vervallen - 1 jr"),
    ("DC", "430/8", "Financiele schulden - Kredietinstellingen"),
    ("DD", "439", "Financiele schulden - Overige leningen"),
    ("DE", "440/4", "Handelsschulden - Leveranciers"),
    ("DF", "441", "Handelsschulden - Te betalen wissels"),
    ("DG", "46", "Ontvangen vooruitbetal. op bestell."),
    ("DH", "450/3", "Belastingen"),
    ("DI", "454/9", "Bezoldigingen & sociale lasten"),
    ("DJ", "47/48", "Overige schulden"),
    ("DK", "492/3", "Overlopende rekeningen")
]

# ==============================================
# HELPER FUNCTIES
# ==============================================

def maak_categorie_mapping(categorieen):
    """Maak een mapping van rekeningcodes naar categorieën"""
    try:
        mapping = defaultdict(list)
        for cat_letter, rek_codes, omschrijving in categorieen:
            if '/' in rek_codes:  # Voor bereiken zoals 631/4
                start, end = rek_codes.split('/')
                if len(start) != len(end):
                    end = start[:-len(end)] + end  # Voor 752/9 → 752-759
                base = start[:-(len(end)-len(start)+1)] if len(start) > len(end) else start[:-len(end)]
                for i in range(int(start[-len(end):]), int(end)+1):
                    full_code = f"{base}{i:0{len(end)}d}"
                    mapping[full_code].append((cat_letter, omschrijving))
            else:
                mapping[rek_codes].append((cat_letter, omschrijving))
        return mapping
    except Exception as e:
        logging.error(f"Fout bij maken categorie mapping: {str(e)}")
        raise

def parse_bedrag(waarde_str):
    """Parse een bedrag string naar een Decimal object"""
    try:
        if not waarde_str:
            return Decimal("0")
        
        # Verwijder alle spaties
        waarde_str = waarde_str.strip()
        
        # Verwerk negatieve bedragen tussen haakjes
        is_negatief = False
        if waarde_str.startswith('(') and waarde_str.endswith(')'):
            waarde_str = waarde_str[1:-1]
            is_negatief = True
        
        # Als er een punt is en daarna een komma, dan is het punt een duizendtalscheiding
        if '.' in waarde_str and ',' in waarde_str:
            waarde_str = waarde_str.replace('.', '')
        
        # Vervang komma door punt voor decimaalteken
        waarde_str = waarde_str.replace(',', '.')
        
        waarde = Decimal(waarde_str)
        if is_negatief:
            waarde = -waarde
        return waarde
    except Exception as e:
        logging.error(f"Fout bij parsen van bedrag '{waarde_str}': {str(e)}")
        return Decimal("0")

def vind_categorie(rekening, mapping):
    """Vind de categorie voor een gegeven rekeningcode"""
    try:
        # Remove any decimal points from the account number
        base_rekening = rekening.split('.')[0]
        
        # Try exact matches first (e.g., 630), then more general ones (e.g., 63)
        # Sort by code length in descending order to try longer matches first
        for code_len in sorted({len(code) for code in mapping.keys()}, reverse=True):
            # Take only the first digits up to code_len
            code = base_rekening[:code_len]
            if code in mapping:
                return mapping[code][0]  # return (letter, omschrijving)
        return None, None
    except Exception as e:
        logging.error(f"Fout bij vinden categorie voor rekening {rekening}: {str(e)}")
        return None, None

def vind_datum_in_pdf(pdf_pad):
    """Zoek naar een datum in het PDF bestand"""
    try:
        datum_formaten = [
            ('%d-%m-%Y', r'\b\d{2}-\d{2}-\d{4}\b'),  # DD-MM-YYYY
            ('%d/%m/%Y', r'\b\d{2}/\d{2}/\d{4}\b'),  # DD/MM/YYYY
            ('%Y-%m-%d', r'\b\d{4}-\d{2}-\d{2}\b'),  # YYYY-MM-DD
            ('%Y/%m/%d', r'\b\d{4}/\d{2}/\d{2}\b'),  # YYYY/MM/DD
            ('%d.%m.%Y', r'\b\d{2}\.\d{2}\.\d{4}\b'),  # DD.MM.YYYY
        ]
        
        gevonden_datums = []
        
        with pdfplumber.open(pdf_pad) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    for datum_format, regex_pattern in datum_formaten:
                        matches = re.findall(regex_pattern, text)
                        for match in matches:
                            try:
                                datum = datetime.strptime(match, datum_format)
                                # Controleer of het de laatste dag van de maand is
                                volgende_maand = datum.replace(day=28) + timedelta(days=4)
                                laatste_dag = volgende_maand - timedelta(days=volgende_maand.day)
                                if datum.day == laatste_dag.day:
                                    gevonden_datums.append(datum)
                            except ValueError:
                                continue
        
        if gevonden_datums:
            return max(gevonden_datums)
        return None
    except Exception as e:
        logging.error(f"Fout bij vinden datum in PDF {pdf_pad}: {str(e)}")
        return None

def verwerk_pdf_sectie(pdf_pad, mapping, categorieen):
    """Verwerk een PDF sectie en extraheer de bedragen per categorie"""
    try:
        sommen = {cat[0]: Decimal("0") for cat in categorieen}
        omschrijvingen = {cat[0]: cat[2] for cat in categorieen}
        rekening_codes = {cat[0]: cat[1] for cat in categorieen}
        
        # Dictionary to track account numbers and their values
        account_values = {}

        with pdfplumber.open(pdf_pad) as pdf:
            all_text = []
            for page in pdf.pages:
                try:
                    text = page.extract_text()
                    if text:
                        lines = [line.strip() for line in text.split('\n') if line.strip()]
                        all_text.extend(lines)
                except Exception as e:
                    logging.warning(f"Waarschuwing bij verwerken pagina: {str(e)}")
                    continue

            tekst = "\n".join(all_text)

        pattern = r"^(\d{2,12}(?:\.\d{1,3})?)\s+(.*?)\s*(-?\d[\d.]*,\d{2}|\(\d[\d.]*,\d{2}\))"

        
        stop_processing = False
        for lijn in tekst.splitlines():
            if stop_processing:
                break
                
            match = re.match(pattern, lijn.strip())
            if match:
                rekening = match.group(1)
                bedrag_str = match.group(3)
                
                if bedrag_str:
                    bedrag = parse_bedrag(bedrag_str)
                    
                    # Check if we've seen this account number before
                    if rekening in account_values:
                        # If the value is the same, we've hit an appendix section
                        if account_values[rekening] == bedrag:
                            logging.info(f"Stopped processing at account {rekening} due to duplicate value {bedrag_str}")
                            stop_processing = True
                            break
                    else:
                        # Store the account and its value
                        account_values[rekening] = bedrag
                        
                        # Only process if we haven't seen this account before
                        cat_letter, _ = vind_categorie(rekening, mapping)
                        if cat_letter:
                            sommen[cat_letter] += bedrag

        return sommen, omschrijvingen, rekening_codes
    except Exception as e:
        logging.error(f"Fout bij verwerken PDF sectie {pdf_pad}: {str(e)}")
        raise

def maak_dataframe(sommen, omschrijvingen, rekening_codes, prefix=""):
    """Maak een DataFrame van de verwerkte data"""
    try:
        data = []
        for cat_letter, omschrijving in omschrijvingen.items():
            totaal = sommen[cat_letter]
            
            if totaal < 0:
                bedrag_str = f"({abs(totaal):,.2f})".replace(".", ",").replace(",", ".", 1)
            else:
                bedrag_str = f"{totaal:,.2f}".replace(".", ",").replace(",", ".", 1)
            
            data.append({
                "Categorie": prefix + cat_letter,
                "Rekening": rekening_codes[cat_letter],
                "Omschrijving": omschrijving,
                "Totaal (€)": bedrag_str
            })
        return data
    except Exception as e:
        logging.error(f"Fout bij maken DataFrame: {str(e)}")
        raise

def convert_currency_to_float(currency_str):
    """Converteer Europese valuta string naar float"""
    try:
        # Verwijder alle spaties
        cleaned = currency_str.strip()
        
        # Verwerk negatieve bedragen tussen haakjes
        is_negatief = False
        if cleaned.startswith('(') and cleaned.endswith(')'):
            cleaned = cleaned[1:-1]
            is_negatief = True
            
        # Verwijder alle punten (duizendtalscheiding)
        cleaned = cleaned.replace('.', '')
        
        # Als er een komma is, gebruik deze als decimaalteken
        if ',' in cleaned:
            parts = cleaned.split(',')
            if len(parts) == 2:  # Normaal geval: één komma
                cleaned = parts[0] + '.' + parts[1]
            else:  # Meerdere komma's, gebruik de laatste als decimaalteken
                cleaned = ''.join(parts[:-1]) + '.' + parts[-1]
        else:
            # Als er geen komma is, veronderstel dat de laatste 2 cijfers centen zijn
            if len(cleaned) > 2:
                cleaned = cleaned[:-2] + '.' + cleaned[-2:]
                
        # Converteer naar float
        waarde = float(cleaned)
        
        # Pas het teken toe
        if is_negatief:
            waarde = -waarde
            
        return waarde
    except Exception as e:
        logging.error(f"Fout bij conversie van '{currency_str}': {str(e)}")
        raise

def extract_rekeningcodes_from_pdf(pdf_pad):
    """Extract unique rekeningcodes and their descriptions from a PDF file"""
    try:
        rekeningcodes = {}
        
        with pdfplumber.open(pdf_pad) as pdf:
            all_text = []
            for page in pdf.pages:
                try:
                    text = page.extract_text()
                    if text:
                        # Split into lines and clean them
                        lines = [line.strip() for line in text.split('\n') if line.strip()]
                        all_text.extend(lines)
                except Exception as e:
                    logging.warning(f"Waarschuwing bij verwerken pagina: {str(e)}")
                    continue

            tekst = "\n".join(all_text)

        # Modified pattern to be more inclusive with descriptions
        pattern = r'^\s*(\d{2,12}(?:\.\d{1,3})?)\s+([^0-9].*?)\s+(-?[\d.,]+(?:,\d{2}|\.\d{2})|\([\d.,]+(?:,\d{2}|\.\d{2})\))'
        
        # Keep track of unique codes to detect duplicates
        seen_codes = set()
        duplicate_found = False
        
        for lijn in tekst.splitlines():
            # Skip empty lines
            if not lijn.strip():
                continue
                
            # Stop if we've found a duplicate section
            if duplicate_found:
                break
                
            match = re.match(pattern, lijn.strip())
            if match:
                rekening = match.group(1)
                description = match.group(2).strip()
                bedrag_str = match.group(3)
                
                logging.debug(f"Found match in line: {lijn}")
                logging.debug(f"Rekening: {rekening}, Description: {description}, Bedrag: {bedrag_str}")
                
                if bedrag_str:
                    bedrag = parse_bedrag(bedrag_str)
                    
                    # Check for duplicates to detect appendix section
                    if rekening in seen_codes:
                        # If we see the same code with the same value, it's likely an appendix
                        if rekening in rekeningcodes and abs(rekeningcodes[rekening][1] - bedrag) < Decimal('0.01'):
                            logging.info(f"Found duplicate code with same value: {rekening}")
                            duplicate_found = True
                            break
                    
                    seen_codes.add(rekening)
                    
                    # Store or update the rekeningcode info
                    if rekening not in rekeningcodes or (description and not rekeningcodes[rekening][0]):
                        rekeningcodes[rekening] = (description, bedrag)
                        logging.debug(f"Stored/Updated code {rekening} with description: {description}")
            else:
                logging.debug(f"No match for line: {lijn}")

        logging.info(f"Extracted {len(rekeningcodes)} unique rekeningcodes from {pdf_pad}")
        return rekeningcodes
    except Exception as e:
        logging.error(f"Fout bij extractie rekeningcodes uit PDF {pdf_pad}: {str(e)}")
        return {}

def create_overview_sheet(wb, pdf_with_dates):
    """Create or update the Overview sheet with rekeningcodes and values from PDFs"""
    try:
        # Create new sheet or get existing one
        sheet_name = "Overview"
        counter = 1
        while sheet_name in wb.sheetnames:
            sheet_name = f"Overview_{counter}"
            counter += 1
        
        ws = wb.create_sheet(sheet_name)
        
        # Process PDFs in chronological order (oldest to newest)
        sorted_pdfs = sorted(pdf_with_dates, key=lambda x: x[1])
        
        # Set up headers
        headers = ["Prefix", "Rekeningcode", "Omschrijving"]
        # Add dates from PDFs as headers
        for _, date in sorted_pdfs:
            headers.append(date.strftime("%d-%m-%Y"))
            
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)
            # Make headers bold
            ws.cell(row=1, column=col).font = openpyxl.styles.Font(bold=True)
        
        # Extract all unique rekeningcodes and their descriptions
        all_codes = {}
        pdf_values = {}
        
        for pdf_file, date in sorted_pdfs:
            logging.info(f"Processing PDF {pdf_file} for Overview sheet")
            codes = extract_rekeningcodes_from_pdf(pdf_file)
            pdf_values[date] = codes
            logging.info(f"Found {len(codes)} codes in PDF {pdf_file}")
            
            for code, (desc, value) in codes.items():
                if code not in all_codes or (desc and not all_codes[code]):
                    all_codes[code] = desc
                    logging.debug(f"Added/Updated code {code} with description: {desc}")

        # Sort rekeningcodes
        sorted_codes = sorted(all_codes.keys(), key=lambda x: (x[:2], x))
        logging.info(f"Total unique codes to be added to Overview: {len(sorted_codes)}")
        
        # Populate the sheet
        for row, code in enumerate(sorted_codes, 2):
            # Prefix (first two digits)
            ws.cell(row=row, column=1, value=code[:2])
            
            # Full rekeningcode
            ws.cell(row=row, column=2, value=code)
            
            # Description
            ws.cell(row=row, column=3, value=all_codes[code])
            
            # Values from each PDF - store raw values without formatting
            for col, (pdf_file, date) in enumerate(sorted_pdfs, 4):
                if date in pdf_values and code in pdf_values[date]:
                    value = float(pdf_values[date][code][1])  # Convert Decimal to float
                    ws.cell(row=row, column=col, value=value)
                    logging.debug(f"Added value {value} for code {code} from PDF {pdf_file}")

        # Auto-adjust column widths
        for col in range(1, len(headers) + 1):
            max_length = 0
            for cell in ws[get_column_letter(col)]:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[get_column_letter(col)].width = adjusted_width

        logging.info(f"Successfully created {sheet_name} sheet with {len(sorted_codes)} rows")
        return True
    except Exception as e:
        logging.error(f"Error creating Overview sheet: {str(e)}")
        return False

def export_naar_template(resultaten_data, balans_data, template_pad, output_pad, kolom, datum=None, pdf_with_dates=None):
    """Exporteert alle categorieën naar template Excel met specifieke kolom, behoudt bestaande waarden en formules"""
    wb = None
    try:
        # Check if output file exists, if not copy template
        if not os.path.exists(output_pad):
            shutil.copy2(template_pad, output_pad)
            logging.info(f"Created new output file from template: {output_pad}")
        
        # Load the existing output file
        wb = openpyxl.load_workbook(
            filename=output_pad,
            data_only=False,
            read_only=False
        )
        
        if 'Historiek' not in wb.sheetnames:
            logging.error("Worksheet 'Historiek' niet gevonden in template")
            return False

        ws = wb['Historiek']
        
        # Vul de datum in als deze beschikbaar is
        if datum:
            datum_cel = f"{kolom}7"
            ws[datum_cel] = datum.strftime("%d-%m-%Y")
            logging.info(f"Datum {datum.strftime('%d-%m-%Y')} ingevuld in cel {datum_cel}")

        # Get rekening 60 value to determine sign conversion
        rekening_60_value = None
        for data in resultaten_data:
            if data['Rekening'] == '60':
                try:
                    rekening_60_value = convert_currency_to_float(data['Totaal (€)'])
                    break
                except:
                    pass

        # Verwerk alle data
        for data in resultaten_data + balans_data:
            cat = data['Categorie']
            rek = data['Rekening']
            waarde_str = data['Totaal (€)']
            
            # Bepaal de doelcel op basis van de rekeningcode
            if rek in categorie_naar_cel:
                base_cel = categorie_naar_cel[rek]
                doel_cel = base_cel.replace('C', kolom)
                
                # Controleer of de cel een formule bevat
                if ws[doel_cel].value is not None and isinstance(ws[doel_cel].value, str) and ws[doel_cel].value.startswith('='):
                    logging.warning(f"Cel {doel_cel} bevat een formule en wordt niet overschreven")
                    continue
                
                try:
                    # Convert the value
                    waarde = convert_currency_to_float(waarde_str)
                    
                    # Special case for '670/3' (Belastingen) - always make it negative
                    if rek == '670/3' and waarde > 0:
                        waarde = -waarde
                    
                    # Apply sign conversion logic for cost rows if rekening 60 is negative
                    row_num = int(doel_cel[1:]) if len(doel_cel) > 1 else 0
                    is_cost_row = (154 <= row_num <= 165) or (174 <= row_num <= 177) or (188 <= row_num <= 194)
                    
                    if rekening_60_value is not None and rekening_60_value < 0 and is_cost_row:
                        if waarde < 0:
                            waarde = abs(waarde)  # Make costs positive
                        else:
                            waarde = -waarde  # Make non-costs negative
                    
                    ws[doel_cel] = waarde
                    logging.info(f"Cel {doel_cel} bijgewerkt met waarde {waarde}")
                except Exception as e:
                    logging.error(f"Fout bij updaten cel {doel_cel}: {str(e)}")
                    continue

        # Create Overview sheet after processing all PDFs
        # Only create Overview sheet when processing the last PDF (most recent one, column F)
        if kolom == 'F' and pdf_with_dates:
            create_overview_sheet(wb, pdf_with_dates)

        # Save the workbook
        wb.save(output_pad)
        wb.close()
        logging.info(f"Workbook succesvol bijgewerkt: {output_pad}")
        return True
                
    except Exception as e:
        logging.error(f"Fout bij exporteren naar template: {str(e)}")
        return False

    finally:
        # Zorg ervoor dat alle workbooks gesloten zijn
        if wb:
            try:
                wb.close()
            except:
                pass

# Maak categorie mappings
try:
    resultaten_mapping = maak_categorie_mapping(resultaten_cat)
    balans_mapping = maak_categorie_mapping(balans_cat)
    logging.info("Categorie mappings succesvol aangemaakt")
except Exception as e:
    logging.error(f"Fout bij aanmaken categorie mappings: {str(e)}")
    raise

# Mapping van rekeningcodes naar Excel cellen
categorie_naar_cel = {
    # Resultatenrekening
    '70': 'C148',    # Omzet
    '71': 'C149',    # Voorraadwijzigingen
    '72': 'C150',    # Geproduceerde vaste activa
    '74': 'C151',    # Andere bedrijfsopbrengsten
    
    '60': 'C157',    # Handelsgoederen, grond- en hulpstoffen
    '609': 'C158',   # Wijzigingen in de voorraad
    '61': 'C159',    # Diensten en diverse goederen
    '62': 'C160',    # Bezoldigingen, sociale lasten en pensioenen
    '630': 'C161',   # Afschrijvingen en waardeverminderingen
    '631/4': 'C162', # Waardeverminderingen op voorraden
    '635/7': 'C163', # Voorzieningen voor risico's en kosten
    '640/8': 'C164', # Andere bedrijfskosten
    '649': 'C165',   # Geactiveerde bedrijfskosten
    
    '750': 'C170',   # Opbrengsten uit financiële vaste activa
    '751': 'C171',   # Opbrengsten uit vlottende activa
    '752/9': 'C172', # Andere financiële opbrengsten
    
    '650': 'C175',   # Kosten van schulden
    '651': 'C176',   # Waardeverminderingen op vlottende activa
    '652/9': 'C177', # Andere financiële kosten
    
    '760': 'C182',   # Terugneming van afschrijvingen
    '761': 'C183',   # Terugneming van waardeverminderingen op financiële vaste activa
    '762': 'C184',   # Terugneming van voorzieningen voor uitzonderlijke risico's
    '763': 'C185',   # Meerwaarden bij de realisatie van vaste activa
    '764/9': 'C186', # Andere uitzonderlijke opbrengsten
    
    '660': 'C189',   # Uitzonderlijke afschrijvingen
    '661': 'C190',   # Waardeverminderingen op financiële vaste activa
    '662': 'C191',   # Voorzieningen voor uitzonderlijke risico's
    '663': 'C192',   # Minderwaarden bij de realisatie van vaste activa
    '664/8': 'C193', # Andere uitzonderlijke kosten
    '669': 'C194',   # Geactiveerde uitzonderlijke kosten
    
    '670/3': 'C208', # Belastingen
    '77': 'C210',    # Regularisering van belastingen
    
    # Balans
    '21': 'C19',     # Immateriële vaste activa
    '22': 'C22',     # Terreinen en gebouwen
    '23': 'C23',     # Installaties, machines en uitrusting
    '24': 'C24',     # Meubilair en rollend materieel
    '25': 'C25',     # Leasing en soortgelijke rechten
    '26': 'C26',     # Overige materiële vaste activa
    '27': 'C27',     # Activa in aanbouw en vooruitbetalingen
    
    '280': 'C31',    # Verbonden ondernemingen - Deelnemingen
    '281': 'C32',    # Verbonden ondernemingen - Vorderingen
    '282': 'C35',    # Ondernemingen met deelnemingsverhouding - Deelnemingen
    '283': 'C36',    # Ondernemingen met deelnemingsverhouding - Vorderingen
    '284': 'C38',    # Andere financiële vaste activa - Aandelen
    '285/8': 'C39',  # Andere financiële vaste activa - Vorderingen en borgtochten
    
    '290': 'C44',    # Handelsvorderingen
    '291': 'C45',    # Overige vorderingen
    
    '30/31': 'C49',  # Grond- en hulpstoffen
    '32': 'C50',     # Goederen in bewerking
    '33': 'C51',     # Gereed product
    '34': 'C52',     # Handelsgoederen
    '35': 'C53',     # Onroerende goederen bestemd voor verkoop
    '36': 'C54',     # Vooruitbetalingen
    '37': 'C55',     # Bestellingen in uitvoering
    
    '40': 'C58',     # Handelsvorderingen
    '41': 'C59',     # Overige vorderingen
    '50': 'C62',     # Eigen aandelen
    '51/53': 'C63',  # Overige beleggingen
    '54/58': 'C65',  # Liquide middelen
    '490/1': 'C67',  # Overlopende rekeningen
    
    '100': 'C78',    # Geplaatst kapitaal
    '101': 'C79',    # Niet opgevraagd kapitaal
    '11': 'C81',     # Uitgiftepremies
    '12': 'C83',     # Herwaarderingsmeerwaarden
    
    '130': 'C86',    # Wettelijke reserve
    '131': 'C87',    # Onbeschikbare reserves
    '132': 'C88',    # Belastingvrije reserves
    '133': 'C89',    # Beschikbare reserves
    '140': 'C91',    # Overgedragen winst
    '141': 'C92',    # Overgedragen verlies
    '15': 'C94',     # Kapitaalsubsidies
    
    '160': 'C99',    # Pensioenen en soortgelijke verplichtingen
    '161': 'C100',   # Belastingen
    '162': 'C101',   # Grote herstellings- en onderhoudswerken
    '163/9': 'C102', # Overige risico's en kosten
    '168': 'C103',   # Uitgestelde belastingen
    
    '170': 'C109',   # Achtergestelde leningen
    '171': 'C110',   # Niet-achtergestelde obligatieleningen
    '172': 'C111',   # Leasingschulden en soortgelijke
    '173': 'C112',   # Kredietinstellingen
    '174': 'C113',   # Overige leningen
    
    '1750': 'C115',  # Handelsschulden - Leveranciers
    '1751': 'C116',  # Handelsschulden - Te betalen wissels
    '176': 'C117',   # Ontvangen vooruitbetalingen op bestellingen
    '178/9': 'C118', # Overige schulden
    
    '42': 'C121',    # Schulden op meer dan één jaar die binnen het jaar vervallen
    '430/8': 'C123', # Kredietinstellingen
    '439': 'C124',   # Overige leningen
    '440/4': 'C126', # Leveranciers
    '441': 'C127',   # Te betalen wissels
    '46': 'C128',    # Ontvangen vooruitbetalingen op bestellingen
    '450/3': 'C130', # Belastingen
    '454/9': 'C131', # Bezoldigingen en sociale lasten
    '47/48': 'C132', # Overige schulden
    '492/3': 'C134'  # Overlopende rekeningen
}

# Process PDF files
def process_pdfs(template_path, pdf_files):
    try:
        # First, get dates for all PDFs and sort them
        pdf_with_dates = []
        for pdf_file in pdf_files:
            try:
                datum = vind_datum_in_pdf(pdf_file)
                if datum:
                    pdf_with_dates.append((pdf_file, datum))
                else:
                    logging.error(f"No valid date found in {pdf_file}")
            except Exception as e:
                logging.error(f"Error processing date from {pdf_file}: {str(e)}")
        
        # Sort PDFs by date, most recent first
        pdf_with_dates.sort(key=lambda x: x[1], reverse=True)
        
        if len(pdf_with_dates) < 1:
            logging.error("No valid PDFs with dates found")
            return False
        
        # Create output file from template
        output_path = 'financial_analysis.xlsx'
        if os.path.exists(output_path):
            os.unlink(output_path)
        shutil.copy2(template_path, output_path)
        
        # Process each PDF and write to the appropriate column
        for i, (pdf_file, datum) in enumerate(pdf_with_dates):
            if i >= 3:  # Only process the three most recent PDFs
                break
                
            try:
                # Map the PDFs to specific columns (F for most recent, E for second, D for oldest)
                kolom = chr(ord('F') - i)  # F, E, D
                
                # Process the PDF for both result and balance data
                resultaten_sommen, resultaten_omschrijvingen, resultaten_codes = verwerk_pdf_sectie(pdf_file, resultaten_mapping, resultaten_cat)
                balans_sommen, balans_omschrijvingen, balans_codes = verwerk_pdf_sectie(pdf_file, balans_mapping, balans_cat)
                
                # Create DataFrames
                resultaten_data = maak_dataframe(resultaten_sommen, resultaten_omschrijvingen, resultaten_codes)
                balans_data = maak_dataframe(balans_sommen, balans_omschrijvingen, balans_codes, prefix="B")
                
                # Export to template
                success = export_naar_template(resultaten_data, balans_data, template_path, output_path, kolom, datum, pdf_with_dates[:3])
                if not success:
                    logging.error(f"Failed to export data to column {kolom}")
                    return False
                    
                logging.info(f"Successfully processed {pdf_file} for column {kolom}")
                
            except Exception as e:
                logging.error(f"Error processing {pdf_file}: {str(e)}")
                return False
        
        logging.info("Successfully processed all PDFs")
        return True
    
    except Exception as e:
        logging.error(f"Error in process_pdfs: {str(e)}")
        return False

if __name__ == "__main__":
    print("This module provides PDF financial statement processing functionality.")
    print("Please use run_financial_analysis.py to execute the processing.")