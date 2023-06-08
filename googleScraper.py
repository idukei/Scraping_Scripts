import requests
from bs4 import BeautifulSoup
import openpyxl

def scrape_google_results(query, num_pages):
    results = []

    for page in range(num_pages):
        url = f"https://www.google.com/search?q={query}&start={page * 10}"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        search_results = soup.find_all("div", class_="g")

        for result in search_results:
            title = result.find("h3").text.strip()
            link = result.find("a")["href"]
            snippet = result.find("div", class_="VwiC3b").text.strip()

            results.append({"Title": title, "Link": link, "Snippet": snippet})

    return results

def save_to_excel(results, file_name):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    sheet["A1"] = "Title"
    sheet["B1"] = "Link"
    sheet["C1"] = "Snippet"

    for i, result in enumerate(results, start=2):
        sheet[f"A{i}"] = result["Title"]
        sheet[f"B{i}"] = result["Link"]
        sheet[f"C{i}"] = result["Snippet"]

    workbook.save(file_name)
    print(f"Results saved to {file_name}")

# Example usage
query = "web scraping"
num_pages = 3
file_name = "google_results.xlsx"

results = scrape_google_results(query, num_pages)
save_to_excel(results, file_name)
