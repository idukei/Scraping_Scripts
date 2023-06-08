import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

def scrape_bet365_odds():
    # URL of the odds page on Bet365
    url = "https://www.bet365.es/#/HO/"
    
    # Set up headers to mimic a web browser
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
    }
    
    # Send a GET request to the website
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        # Create a BeautifulSoup object to parse the HTML content
        soup = BeautifulSoup(response.content, "html.parser")
        
        # Find the table containing the odds
        table = soup.find("table", {"class": "gl-MarketGroup_Wrapper"})
        
        if table:
            # Create a new Excel workbook and select the active sheet
            workbook = Workbook()
            sheet = workbook.active
            
            # Find all the rows in the table
            rows = table.find_all("tr", {"class": "gl-MarketGroup_Wrapper"})
            
            for row in rows:
                # Find the team names
                team_elements = row.find_all("div", {"class": "fh-ParticipantFixtureTeamSoccer_TeamName "})
                
                if len(team_elements) >= 2:
                    team1 = team_elements[0].text.strip()
                    team2 = team_elements[1].text.strip()
                    
                    # Find the odds
                    odds_elements = row.find_all("span", {"class": "fh-ParticipantFixtureOdd_Odds"})
                    
                    if len(odds_elements) >= 2:
                        odds1 = odds_elements[0].text.strip()
                        odds2 = odds_elements[1].text.strip()
                        
                        # Write the data to the Excel sheet
                        sheet.append([team1, team2, odds1, odds2])
            
            # Save the workbook as an Excel file
            workbook.save("bet365_odds.xlsx")
            
            print("Scraping completed. Data saved in 'bet365_odds.xlsx'.")
        else:
            print("Table not found on the page.")
    else:
        print("Failed to retrieve the page.")

# Call the function to start the scraping process
scrape_bet365_odds()
