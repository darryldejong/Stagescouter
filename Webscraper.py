import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import Alignment
from colorama import init, Fore, Style
import os
import sys

init(autoreset=True)

def run_script():
    print("***********************************************")
    print(Style.BRIGHT + Fore.RED + "WAARSCHUWING: Dit is alleen voor educatieve doeleinden.")
    print(Style.BRIGHT + Fore.BLUE + "Dit script dient als leermiddel voor webscraping en moet met zorg worden gebruikt.")
    print(Style.BRIGHT + Fore.WHITE + "Zorg ervoor dat je voldoet aan de voorwaarden van de website.")
    print("***********************************************\n")

    print(Style.BRIGHT + Fore.YELLOW + "Dit script is bedoeld om het vinden van stages gemakkelijker te maken en ze in een Excel-bestand op te slaan.\n")

    print("Keuzeopties voor zoeken:")
    print("1. Zoek op trefwoord (bijv. 'Voetbal')")
    print("2. Zoek op plaats (bijv. 'Utrecht')")
    print("3. Zoek op trefwoord en plaats (bijv. 'Allround medewerker IT systems and devices' in 'Nieuwegein')")

    print()  

    while True:
        search_type = input("Kies een zoekoptie (1 voor trefwoord, 2 voor plaats, 3 voor trefwoord en plaats): ").strip()
        if search_type in ['1', '2', '3']:
            break
        else:
            print(Fore.RED + "Kies (1,2,3)")

    if search_type == '1':
        keyword = input("Voer een trefwoord in voor de zoekopdracht: ").strip()
        base_url = f"https://stagemarkt.nl/vacatures/?Termen={keyword}&PlaatsPostcode=&Straal=0&Land=e883076c-11d5-11d4-90d3-009027dcddb5&ZoekenIn=A&Page={{}}&Longitude=&Latitude=&Regio=&Plaats=&Niveau=&SBI=&Kwalificatie=&Sector=&RandomSeed=486&Leerweg=&Bedrijfsgrootte=&Opleidingsgebied=&Internationaal=&Beschikbaarheid=&AlleWerkprocessenUitvoerbaar=&LeerplaatsGewijzigd=&Sortering=0&Bron=STA&Focus=&LeerplaatsKenmerk=&OrganisatieKenmerk="
        print(f"Zoekopdracht aan het uitvoeren voor trefwoord: {keyword}...")
        output_name = keyword
    elif search_type == '2':
        plaats = input("Voer een plaats in voor de zoekopdracht: ").strip()
        base_url = f"https://stagemarkt.nl/vacatures/?Termen=&PlaatsPostcode={plaats}&Straal=0&Land=e883076c-11d5-11d4-90d3-009027dcddb5&ZoekenIn=A&Page={{}}&Longitude=&Latitude=&Regio=&Plaats=&Niveau=&SBI=&Kwalificatie=&Sector=&RandomSeed=486&Leerweg=&Bedrijfsgrootte=&Opleidingsgebied=&Internationaal=&Beschikbaarheid=&AlleWerkprocessenUitvoerbaar=&LeerplaatsGewijzigd=&Sortering=0&Bron=STA&Focus=&LeerplaatsKenmerk=&OrganisatieKenmerk="
        print(f"Zoekopdracht aan het uitvoeren voor plaats: {plaats}...")
        output_name = plaats
    elif search_type == '3':
        keyword = input("Voer een trefwoord in voor de zoekopdracht: ").strip()
        plaats = input("Voer een plaats in voor de zoekopdracht: ").strip()

        while True:
            straal = input("Kies een straal in kilometers (0, 5, 10, 15, 20, 25): ").strip()
            if straal in ['0', '5', '10', '15', '20', '25']:
                break
            else:
                print(Fore.RED + "Kies een geldige straal!")

        base_url = f"https://stagemarkt.nl/vacatures/?Termen={keyword}&PlaatsPostcode={plaats}&Straal={straal}&Land=e883076c-11d5-11d4-90d3-009027dcddb5&ZoekenIn=A&Page={{}}&Longitude=&Latitude=&Regio=&Plaats=&Niveau=&SBI=&Kwalificatie=&Sector=&RandomSeed=486&Leerweg=&Bedrijfsgrootte=&Opleidingsgebied=&Internationaal=&Beschikbaarheid=&AlleWerkprocessenUitvoerbaar=&LeerplaatsGewijzigd=&Sortering=0&Bron=STA&Focus=&LeerplaatsKenmerk=&OrganisatieKenmerk="
        
        print(f"Zoekopdracht aan het uitvoeren: {keyword}, {plaats} {straal} km.")
        output_name = keyword
    else:
        print(Fore.RED + "Kies een geldige keuze! Probeer het opnieuw.")
        return  # Terug naar het begin van de functie zonder het script af te sluiten

    job_data = []

    total_pages = 100  # Aantal pagina's
    for page_number in range(1, total_pages + 1):
        url = base_url.format(page_number)
        print(f"Pagina {page_number}/{total_pages} wordt verwerkt...", end='\r')

        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
            }

            response = requests.get(url, headers=headers)
            response.raise_for_status()

            soup = BeautifulSoup(response.text, 'html.parser')

            job_listings = soup.find_all('div', class_='c-link-blocks-single-info')

            if not job_listings:
                pass

            for job in job_listings:
                title = job.find('h2').get_text(strip=True) if job.find('h2') else 'Geen titel beschikbaar'
                company_and_location = job.find_all('h3')
                company = company_and_location[0].get_text(strip=True) if company_and_location else 'Geen bedrijfsnaam'
                start_date, place, learning_path, level = ('Geen startdatum', 'Geen locatie', 'Geen leerweg', 'Geen niveau')

                ul_items = job.find('ul', class_='u-reset-ul').find_all('li') if job.find('ul', class_='u-reset-ul') else []
                if len(ul_items) > 0:
                    start_date = ul_items[0].get_text(strip=True).split('Startdatum')[-1]
                if len(ul_items) > 1:
                    place = ul_items[1].get_text(strip=True).split('Plaats')[-1]
                if len(ul_items) > 2:
                    learning_path = ul_items[2].get_text(strip=True).split('Leerweg')[-1]
                if len(ul_items) > 3:
                    level = ul_items[3].get_text(strip=True).split('Niveau')[-1]

                job_data.append({
                    'Titel': title,
                    'Bedrijf': company,
                    'Startdatum': start_date,
                    'Plaats': place,
                    'Leerweg': learning_path,
                    'Niveau': level
                })

        except requests.exceptions.RequestException as e:
            print(Fore.RED + f"Er is een fout opgetreden bij het verwerken van pagina {page_number}: {e}")

    if len(job_data) == 0:
        print(Fore.RED + f"Er zijn 0 zoekresultaten gevonden voor jouw zoekterm: {output_name}")
    else:
        df = pd.DataFrame(job_data)

        excel_filename = f'Stagezoekopdracht_{output_name}.xlsx'
        df.to_excel(excel_filename, index=False)

        with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Vacatures")
            workbook = writer.book
            worksheet = writer.sheets["Vacatures"]

            column_widths = {
                'A': 50,  # Titel
                'B': 40,  # Bedrijf
                'C': 20,  # Startdatum
                'D': 20,  # Plaats
                'E': 20,  # Leerweg
                'F': 20   # Niveau
            }

            for col, width in column_widths.items():
                worksheet.column_dimensions[col].width = width

            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=6):
                for cell in row:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        print(Fore.GREEN + f"De zoekresultaten zijn succesvol verzameld en opgeslagen in '{excel_filename}'.")

    while True:
        run_again = input("Wil je het script opnieuw uitvoeren? (ja/nee): ").strip().lower()
        if run_again in ['ja', 'nee']:
            break
        else:
            print(Fore.RED + "Kies (ja, nee)")

    if run_again == "ja":
        os.system('cls' if os.name == 'nt' else 'clear')  # Terminal schoonvegen
        run_script()
    else:
        print()
        sys.exit()

# Start het script
run_script()
