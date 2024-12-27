import asyncio
import datetime
import math
import re
import subprocess
import os

from playwright.async_api import async_playwright
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

BATCH_SIZE = 20
WAIT_BETWEEN_BATCHES = 30
WAIT_BETWEEN_PLATES = 8

def load_plates_from_file(filename="plaques.txt"):
    with open(filename, "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]

async def check_plate(browser, plate):
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
        await page.fill("input.input-no-focus", plate)

        # 4) Click + short pause
        await page.click("text=Vérifier mes péages à payer")
        await page.wait_for_timeout(3000)

        # 5) Body text
        body_text = await page.text_content("body") or ""

        # 6) Determine status
        if "Aucun passage en attente de paiement" in body_text:
            status = "Rien à payer"
            montant = "0 €"
            print(f"{plate}: Rien à payer")
        else:
            status = "Péages dus"
            print(f"{plate}: Péages dus")

            # Try to find total amount by selector
            total_elem = await page.query_selector("span.total-amount")
            if total_elem:
                extracted = (await total_elem.text_content() or "").strip()
                montant = extracted
                print(f" - Montant trouvé via span.total-amount: {montant}")
            else:
                # Fallback Regex for amount
                pattern = r"\b(\d{1,3},\d{2}\s?€)\b"
                match = re.search(pattern, body_text)
                if match:
                    montant = match.group(1)
                    print(f" - Montant trouvé via regex: {montant}")
                else:
                    montant = "Inconnu"
                    print(" - Impossible de trouver le montant.")

            # 7) Try to parse a “date de passage” from the page
            #    Example regex for a French format DD/MM/YYYY
            date_passage_pattern = r"(\d{2}/\d{2}/\d{4})"
            match_date = re.search(date_passage_pattern, body_text)
            if match_date:
                date_passage = match_date.group(1)
                print(f" - Date de passage trouvée : {date_passage}")
            else:
                # If there's no match, we leave it as “N/A”
                pass

    except Exception as e:
        print(f"{plate}: Erreur - {e}")
    finally:
        await page.close()

    # Wait between each plate
    await asyncio.sleep(WAIT_BETWEEN_PLATES)

    # Return the new date field in the tuple
    return (plate, status, montant, date_passage)

async def get_browser_instance(p):
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

async def process_batch_sequential(browser, plates):
    results = []
    for plate in plates:
        result = await check_plate(browser, plate)
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
    # 2) Normal script logic
    ############################################################################
    plates_to_check = load_plates_from_file("plaques.txt")
    if not plates_to_check:
        print("Aucune plaque trouvée.")
        return

    async with async_playwright() as p:
        browser = await get_browser_instance(p)

        all_results = []
        total_plates = len(plates_to_check)
        num_batches = math.ceil(total_plates / BATCH_SIZE)

        for i in range(num_batches):
            batch_plates = plates_to_check[i*BATCH_SIZE : (i+1)*BATCH_SIZE]
            print(f"\n=== Batch {i+1}/{num_batches}, {len(batch_plates)} plaques ===")
            batch_results = await process_batch_sequential(browser, batch_plates)
            all_results.extend(batch_results)

            # Wait between batches
            if i < num_batches - 1:
                print(f"Attente de {WAIT_BETWEEN_BATCHES} s avant le prochain batch...")
                await asyncio.sleep(WAIT_BETWEEN_BATCHES)

        await browser.close()

    # Sort + Excel
    # Now each result is (plate, status, montant, date_passage)
    results_sorted = sorted(all_results, key=lambda x: x[0])
    print("\n=== Récapitulatif final ===")
    for plate, status, montant, date_passage in results_sorted:
        print(f"{plate} -> {status} (Montant: {montant}), Date de passage: {date_passage}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Résultats péages"

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 20  # For the Date de passage column

    # Header row
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append(["Heure et Date", "Plaque", "Statut", "Montant", "Date de passage"])

    header_font = Font(bold=True, size=12)
    header_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill

    # Write data rows
    for plate, status, montant, date_passage in results_sorted:
        ws.append([now_str, plate, status, montant, date_passage])

    excel_filename = "resultats.xlsx"
    wb.save(excel_filename)
    print(f"\nFichier Excel enregistré sous {excel_filename}")

    ############################################################################
    # 3) Open resultats.xlsx at the end (Windows only)
    ############################################################################
    try:
        os.startfile(excel_filename)
    except AttributeError:
        print("os.startfile() is not supported on this OS.")
    except Exception as e:
        print(f"Could not open {excel_filename}: {e}")


if __name__ == "__main__":
    asyncio.run(main())
