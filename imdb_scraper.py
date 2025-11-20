import os
import platform
import subprocess
import pandas as pd
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

def scrape_imdb_top_n(headless=False):
    # Ask the user for number of movies
    try:
        n = int(input("Enter the number of top movies to scrape (1-250): "))
        if n < 1 or n > 250:
            print("Number must be between 1 and 250. Defaulting to 10 movies.")
            n = 10
    except ValueError:
        print("Invalid input. Defaulting to 10 movies.")
        n = 10

    # Install ChromeDriver automatically
    chromedriver_autoinstaller.install()

    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    driver = webdriver.Chrome(options=options)
    driver.get("https://www.imdb.com/chart/top/")

    # Wait for movies to load
    movies = WebDriverWait(driver, 30).until(
        EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "ul.ipc-metadata-list li.ipc-metadata-list-summary-item")
        )
    )

    movie_list = []
    for rank, movie in enumerate(movies[:n], start=1):
        try:
            # Basic details
            title_elem = movie.find_element(By.CSS_SELECTOR, "h3.ipc-title__text")
            title = title_elem.text

            year = movie.find_element(By.CSS_SELECTOR, "span.cli-title-metadata-item").text
            rating = movie.find_element(By.CSS_SELECTOR, "span.ipc-rating-star--rating").text

            # Movie URL
            link_elem = movie.find_element(By.CSS_SELECTOR, "a")
            movie_url = link_elem.get_attribute("href")

            # Poster URL
            poster_elem = movie.find_element(By.CSS_SELECTOR, "img")
            poster_url = poster_elem.get_attribute("loadlate") or poster_elem.get_attribute("src")

        except Exception as e:
            print(f"Error scraping movie {rank}: {e}")
            title = year = rating = movie_url = poster_url = "N/A"

        movie_list.append({
            "Rank": rank,
            "Title": title,
            "Year": year,
            "IMDb Rating": rating,
            "Movie URL": movie_url,
            "Poster URL": poster_url
        })

        print(f"Scraped: {rank}. {title}")

    driver.quit()

    # Save to Excel
    df = pd.DataFrame(movie_list)
    excel_file = f"IMDb_Top_{n}.xlsx"
    df.to_excel(excel_file, index=False)

    # Format Excel
    wb = load_workbook(excel_file)
    ws = wb.active

    # Set headers bold and center aligned
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Center align all cells and adjust column widths
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 5

    # Make Movie URL column clickable
    url_col_index = None
    for idx, cell in enumerate(ws[1], start=1):
        if cell.value == "Movie URL":
            url_col_index = idx
            break

    if url_col_index:
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=url_col_index)
            if cell.value and cell.value != "N/A":
                cell.hyperlink = cell.value
                cell.style = "Hyperlink"
                cell.alignment = Alignment(horizontal="center", vertical="center")

    wb.save(excel_file)

    print(f"\n🎉 Scraping Completed! {n} movies saved to {excel_file}")

    # Automatically open Excel
    try:
        if platform.system() == "Windows":
            os.startfile(excel_file)
        elif platform.system() == "Darwin":  # macOS
            subprocess.call(["open", excel_file])
        else:  # Linux
            subprocess.call(["xdg-open", excel_file])
    except Exception as e:
        print(f"Could not open the Excel file automatically: {e}")

if __name__ == "__main__":
    scrape_imdb_top_n(headless=False)

