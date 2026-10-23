from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
import json
import pandas as pd

count = int(input("How many movies do you want? (max 250): "))
if count > 250:
    count = 250

print("🚀 Starting...")

chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option("useAutomationExtension", False)
chrome_options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)

driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3]});
        Object.defineProperty(navigator, 'languages', {get: () => ['en-US','en']});
    """
})

try:
    driver.get("https://www.imdb.com/chart/top/")
    print("✅ IMDB Loaded!")
    time.sleep(8)

    script = driver.find_element(By.ID, "__NEXT_DATA__")
    data = json.loads(script.get_attribute("innerHTML"))
    movies = data["props"]["pageProps"]["pageData"]["chartTitles"]["edges"]

    # Store in list
    movie_list = []
    for i, edge in enumerate(movies[:count]):
        node = edge["node"]
        movie_list.append({
            "Rank"  : i + 1,
            "Movie Name": node["titleText"]["text"],
            "Year"  : node["releaseYear"]["year"],
            "Rating": node["ratingsSummary"]["aggregateRating"]
        })

    # Create DataFrame
    df = pd.DataFrame(movie_list)

    print(f"\n🎬 TOP {count} MOVIES 🎬")
    print("-" * 50)
    print(df.to_string(index=False))
    print("-" * 50)

    # Save to Excel
    df.to_excel("imdb_top_movies.xlsx", index=False)
    print(f"✅ Done! Saved to imdb_top_movies.xlsx")

except Exception as e:
    print(f"❌ Error: {e}")

finally:
    driver.quit()