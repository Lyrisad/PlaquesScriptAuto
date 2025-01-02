import asyncio
import datetime
import math
import re
import subprocess
import os

from playwright.async_api import async_playwright
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

BATCH_SIZE = 20
WAIT_BETWEEN_BATCHES = 31  
WAIT_BETWEEN_PLATES = 10  

def load_plates_from_excel(filename="BaseDePlaques.xlsx"):
    """
    Lit un fichier Excel contenant 3 colonnes:
      - IMMATRICULATION
      - Categorie vehicule
      - Proprietaire

    Retourne une liste de tuples (immatriculation, categorie, proprietaire).
    """
    wb = load_workbook(filename)
    ws = wb.active

    data = []
    # On part de row=2 pour ignorer la ligne d'en-tête
    for row in ws.iter_rows(min_row=2, values_only=True):
        # row sera un tuple du type (IMMATRICULATION, Categorie vehicule, Proprietaire)
        if not row or all(cell is None for cell in row):
            continue  # ignorer les lignes vides
        immat, categorie, proprietaire = row
        # On stocke ce trio pour usage ultérieur
        data.append((str(immat).strip(), str(categorie).strip(), str(proprietaire).strip()))
    return data

async def check_plate(browser, immat, categorie, proprietaire):
    """
    Lance la vérification pour 1 plaque.
    Retourne (immat, categorie, proprietaire, status, montant, date_passage).
    """
    page = await browser.new_page()
    status = "Erreur"
    montant = "N/A"
    date_passage = "N/A"  # Renamed field: "Date de passage"

    try:
        # 1) Goto
        await page.goto("https://www.sanef.com/client/index.html#basket",
                        wait_until="networkidle",
                        timeout=20000)
        
        # 2) Cookies
        try:
            await page.click(".tarteaucitronCTAButton", timeout=8000)
        except:
            pass

        # 3) Fill the plate
        await page.wait_for_selector("input.input-no-focus", timeout=12000)
        await page.fill("input.input-no-focus", immat)

        # 4) Click + short pause
        await page.click("text=Vérifier mes péages à payer")
        await page.wait_for_timeout(4000)

        # 5) Body text
        body_text = await page.text_content("body") or ""

        # 6) Determine status
        if "Aucun passage en attente de paiement" in body_text:
            status = "Rien à payer"
            montant = "0 €"
            print(f"{immat} ({categorie}, {proprietaire}): Rien à payer")
        else:
            print(f"{immat} ({categorie}, {proprietaire}): Péages dus")

            # Try to find total amount by selector
            total_elem = await page.query_selector("span.total-amount")
            if total_elem:
                extracted = (await total_elem.text_content() or "").strip()
                montant = extracted
                print(f" - Montant trouvé via span.total-amount: {montant}")
            else:
                # Fallback Regex for amounts
                pattern = r"\b(\d{1,3},\d{2}\s?€)\b"
                matches = re.findall(pattern, body_text)
                if matches:
                    # Convert all matches to float for comparison
                    amounts = []
                    for m in matches:
                        # Nettoyer la chaîne pour convertir en float
                        clean_m = m.replace('€', '').replace(',', '.').strip()
                        try:
                            amount = float(clean_m)
                            amounts.append(amount)
                        except ValueError:
                            pass  # Ignorer les valeurs qui ne peuvent pas être converties

                    if amounts:
                        # Sélectionner le montant le plus élevé
                        max_amount = max(amounts)
                        montant = f"{max_amount:.2f} €"
                        print(f" - Montant total trouvé: {montant}")
                    else:
                        montant = "Inconnu"
                        print(" - Impossible de trouver un montant valide.")
                else:
                    montant = "Inconnu"
                    print(" - Impossible de trouver le montant.")

            # 7) Try to parse all “date de passage” from the page
            #    Example regex for a French format DD/MM/YYYY
            date_passage_pattern = r"(\d{2}/\d{2}/\d{4})"
            matches_date = re.findall(date_passage_pattern, body_text)
            if matches_date:
                # Enlever les doublons et trier les dates (optionnel)
                unique_dates = sorted(set(matches_date), key=lambda date: datetime.datetime.strptime(date, "%d/%m/%Y"))
                date_passage = ", ".join(unique_dates)
                num_peages = len(unique_dates)
                status = f"Péages dus: {num_peages}"
                print(f" - Dates de passage trouvées : {date_passage}")
                print(f" - Nombre de péages dus : {num_peages}")
            else:
                date_passage = "N/A"
                num_peages = 0
                status = "Péages dus: 0"
                print(" - Impossible de trouver la date de passage.")

    except Exception as e:
        print(f"{immat}: Erreur - {e}")
    finally:
        await page.close()

    # Wait between each plate
    await asyncio.sleep(WAIT_BETWEEN_PLATES)

    # Return full info
    return (immat, categorie, proprietaire, status, montant, date_passage)

async def get_browser_instance(p):
    """
    Tente Chrome, puis Edge, puis Firefox.
    """
    try:
        print("Trying Chrome...")
        return await p.chromium.launch(channel="chrome", headless=False)
    except:
        pass
    try:
        print("Trying Edge...")
        return await p.chromium.launch(channel="msedge", headless=False)
    except:
        pass
    try:
        print("Trying Firefox...")
        return await p.firefox.launch(headless=False)
    except:
        pass
    raise RuntimeError("No supported browser found.")

async def process_batch_sequential(browser, plates_data):
    """
    'plates_data' est une liste de tuples (immat, categorie, proprietaire).
    On les traite un par un, en séquentiel.
    """
    results = []
    for (immat, cat, prop) in plates_data:
        result = await check_plate(browser, immat, cat, prop)
        results.append(result)
    return results

async def main():
    ############################################################################
    # 1) Close the Excel process if open (BRUTAL: kills all Excel instances)
    ############################################################################
    try:
        subprocess.run(["taskkill", "/IM", "EXCEL.EXE", "/F"], check=True)
        print("Closed any running Excel instance to free up 'resultats.xlsx'.")
    except subprocess.CalledProcessError:
        pass
    except FileNotFoundError:
        pass

    ############################################################################
    # 2) Load plates from Excel instead of .txt
    ############################################################################
    plates_data = load_plates_from_excel("BaseDePlaques.xlsx")
    if not plates_data:
        print("BaseDePlaques.xlsx est vide ou invalide.")
        return

    # Example: plates_data = [
    #   ("AB123CD", "Voiture", "Jean Dupont"),
    #   ("XY456ZA", "Camion",  "Entreprise X"),
    #    ...
    # ]

    ############################################################################
    # 3) Normal script logic (Playwright checks)
    ############################################################################
    async with async_playwright() as p:
        browser = await get_browser_instance(p)

        all_results = []
        total_plates = len(plates_data)
        num_batches = math.ceil(total_plates / BATCH_SIZE)

        for i in range(num_batches):
            # Extract the portion for this batch
            start = i * BATCH_SIZE
            end = start + BATCH_SIZE
            batch_plates = plates_data[start:end]
            print(f"\n=== Batch {i+1}/{num_batches}, {len(batch_plates)} plaques ===")
            # Process them sequentially
            batch_results = await process_batch_sequential(browser, batch_plates)
            all_results.extend(batch_results)

            # Wait between batches
            if i < num_batches - 1:
                print(f"Attente de {WAIT_BETWEEN_BATCHES} s avant le prochain batch...")
                await asyncio.sleep(WAIT_BETWEEN_BATCHES)

        await browser.close()

    # all_results is now a list of tuples:
    # (immat, categorie, proprietaire, status, montant, date_passage)

    # Sort by immatriculation (index=0)
    results_sorted = sorted(all_results, key=lambda x: x[0])
    print("\n=== Récapitulatif final ===")
    for immat, cat, prop, status, montant, date_passage in results_sorted:
        print(f"{immat} ({cat}, {prop}) -> {status} (Montant: {montant}), Date de passage: {date_passage}")

    ############################################################################
    # 4) Write results to resultats.xlsx with conditional formatting
    ############################################################################
    wb = Workbook()
    ws = wb.active
    ws.title = "Résultats péages"

    ws.column_dimensions["A"].width = 20  # Heure et Date
    ws.column_dimensions["B"].width = 20  # Immatriculation
    ws.column_dimensions["C"].width = 30  # Catégorie
    ws.column_dimensions["D"].width = 20  # Propriétaire
    ws.column_dimensions["E"].width = 30  # Statut
    ws.column_dimensions["F"].width = 15  # Montant
    ws.column_dimensions["G"].width = 30  # Date de passage

    # Define fills for conditional formatting
    peages_due_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")  # Light Red
    rien_a_payer_fill = PatternFill(start_color="99FF99", end_color="99FF99", fill_type="solid")  # Light Green

    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # Header row
    ws.append(["Heure et Date",
               "IMMATRICULATION",
               "Categorie vehicule",
               "Proprietaire",
               "Statut",
               "Montant",
               "Date de passage"])

    header_font = Font(bold=True, size=12)
    header_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill

    # Fill rows with conditional formatting
    for immat, cat, prop, status, montant, date_passage in results_sorted:
        ws.append([now_str, immat, cat, prop, status, montant, date_passage])
        current_row = ws.max_row
        statut_cell = ws.cell(row=current_row, column=5)  # Column E: Statut

        if statut_cell.value.startswith("Péages dus"):
            statut_cell.fill = peages_due_fill
        else:
            statut_cell.fill = rien_a_payer_fill

    excel_filename = "resultats.xlsx"
    wb.save(excel_filename)
    print(f"\nFichier Excel enregistré sous {excel_filename}")

    ############################################################################
    # 5) Open resultats.xlsx at the end (Windows only)
    ############################################################################
    try:
        os.startfile(excel_filename)
    except AttributeError:
        print("os.startfile() is not supported on this OS.")
    except Exception as e:
        print(f"Could not open {excel_filename}: {e}")

if __name__ == "__main__":
    asyncio.run(main())
