import pandas as pd
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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

    # Install correct ChromeDriver
    chromedriver_autoinstaller.install()

    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    driver = webdriver.Chrome(options=options)
    driver.get("https://www.imdb.com/chart/top/")

    # ⭐ Updated selector for IMDb’s new layout
    movies = WebDriverWait(driver, 30).until(
        EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "ul.ipc-metadata-list li.ipc-metadata-list-summary-item")
        )
    )

    movie_list = []
    for rank, movie in enumerate(movies[:n], start=1):
        title = movie.find_element(By.CSS_SELECTOR, "h3.ipc-title__text").text
        year = movie.find_element(By.CSS_SELECTOR, "span.cli-title-metadata-item").text
        rating = movie.find_element(By.CSS_SELECTOR, "span.ipc-rating-star--rating").text

        movie_list.append({
            "Rank": rank,
            "Title": title,
            "Year": year,
            "IMDb Rating": rating
        })

    driver.quit()

    # Save to Excel
    df = pd.DataFrame(movie_list)
    excel_file = f"IMDb_Top_{n}.xlsx"
    df.to_excel(excel_file, index=False)

    print(f"\n🎉 Scraping Completed! {n} movies saved to {excel_file}")

if __name__ == "__main__":
    scrape_imdb_top_n(headless=False)
