import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


ROOT_URL = "http://vr.co.lancaster.pa.us/ElectionReturns/November_5,_2019_-_Municipal_Election/Categories.html"
IGNORED_LINKS = ["RETURN", "QUESTIONS", "CATEGORIES"]

SAVE = False


def main():
    print("Downloading Lancaster County election results...")
    
    results = list()
    
    results += parse_page(ROOT_URL, 1)
    
    if SAVE:
        wb = Workbook()
        sheet = wb.active
        sheet_row = 1

        for row_number, row in enumerate(results, start=1):
            for col_number, value in enumerate(row, start=1):
                sheet.cell(row=row_number, column=col_number).value = value

        wb.save('lancaster_election_results.xlsx')
    
    
def parse_page(url, recursion_level):
    print(url, recursion_level)
    results = []
    
    page = requests.get(url)
    soup = BeautifulSoup(page.text, 'html.parser')
    
    # Page has a 'Results table' on it.
    if len(soup.find_all('br')) == 2:
        summary_table = soup.find_all('br')[0].next_sibling.next_sibling
        results_table = soup.find_all('br')[-1].previous_sibling.previous_sibling
       
            
        office = summary_table.find_all('tr')[0].text.strip()
        
        vote_for = [x.upper() for x in summary_table.find_all('tr')[-1].find_all('td')[-1].text.strip("()").split()]
        
        if "ONE" in vote_for:
            vote_for = 1
        elif "TWO" in vote_for:
            vote_for = 2
        elif "THREE" in vote_for:
            vote_for = 3
        elif "FOUR" in vote_for:
            vote_for = 4
        elif "FIVE" in vote_for:
            vote_for = 5
            
        
        row = [office, vote_for]
        
        for tr in results_table.find_all('tr'):
            
            if len(tr.find_all('td')) == 2:
                name = tr.find_all('td')[0].text
                votes = tr.find_all('td')[-1].text

                if name.upper() == "BY PRECINCT":
                    break
                    
                else:
                    results.append(row + [name, int(votes)])
                    
    # Click all the links on the page
    else:
        for a in soup.find_all('a'):
            if a.text.upper() not in IGNORED_LINKS:
                results += parse_page(a['href'], recursion_level + 1)
                
    return results

if __name__ == '__main__':
    main()