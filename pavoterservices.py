from datetime import date

import requests
from openpyxl import Workbook
from bs4 import BeautifulSoup

ROOT_URL = 'https://www.pavoterservices.pa.gov/ElectionInfo/ElectionInfo.aspx'

def main():
    """
    Scrape demographic and contact info for all candidates listed on the PA Voter Services website and save to an Excel file.
    """
    all_candidates = []
    
    with requests.Session() as session:
        
        page = session.get(ROOT_URL) # Start on page 1
        page_number = 1
        
        while page:
            soup = BeautifulSoup(page.content, 'html.parser')
            
            context = get_context(soup)
            
            for a in soup.find_all('a', href=True):
                link = a['href']
                
                # Find all of the Candidate Info links on the current page, not including the shortcuts to their Petition page
                if 'CandidateInfo.aspx' in link and '&Tab=PET' not in link:
                    
                    candidate_link = f"https://www.pavoterservices.pa.gov/ElectionInfo/{link}&Tab=DET" # create an absolute link to the Candidate's Detail page
                    
                    candidate_context = {
                        '__EVENTTARGET': candidate_link,
                        **context
                    }
                    
                    candidate_page = requests.post(candidate_link, data=candidate_context)
                    candidate_soup = BeautifulSoup(candidate_page.content, 'html.parser')          

                    candidate_info = parse_candidate_soup(candidate_soup)
                    all_candidates.append(candidate_info)
                    #print(candidate_info)

            page_number += 1
            print(f"Getting page {page_number}...")
            
            page = get_next_page(session, soup, context)
            
    save_as_excel(all_candidates)
    
    print('Done.')
            
            
def get_context(soup):
    """
    Pull all of the relevant VIEWSTATE and VALIDATION parameters for the current page into a dictionary and return it.
    """
    
    viewstate = soup.find('input', {'id': '__VIEWSTATE'})['value']
    eventvalidation = soup.find('input', {'id': '__EVENTVALIDATION'})['value']

    viewstatefieldcount = soup.find('input', {'id': '__VIEWSTATEFIELDCOUNT'})['value'] if soup.find('input', {'id': '__VIEWSTATEFIELDCOUNT'}) else None
    viewstate1 = soup.find('input', {'id': '__VIEWSTATE1'})['value'] if soup.find('input', {'id': '__VIEWSTATE1'}) else None
    viewstate2 = soup.find('input', {'id': '__VIEWSTATE2'})['value'] if soup.find('input', {'id': '__VIEWSTATE2'}) else None
    viewstate3 = soup.find('input', {'id': '__VIEWSTATE3'})['value'] if soup.find('input', {'id': '__VIEWSTATE3'}) else None
    
    context = {}
    
    context['__VIEWSTATE'] = viewstate
    context['__EVENTVALIDATION'] = eventvalidation
    
    if viewstatefieldcount:
        context['__VIEWSTATEFIELDCOUNT'] = viewstatefieldcount
    
    if viewstate1:
        context['__VIEWSTATE1'] = viewstate1
        
    if viewstate2:
        context['__VIEWSTATE2'] = viewstate2
        
    if viewstate3:
        context['__VIEWSTATE3'] = viewstate3
    
    
    return context


def get_next_page(session, soup, context):
    """
    Send a POST request to the page to click the 'Next Page' button, if it is not disabled (i.e. you are not on the last page).
    If you are on the last page, return None. Otherwise, return the response to the POST request.
    """
    
    if soup.find(class_='NextItemDisabled'):
        return None
    
    context = {
        '__EVENTTARGET': 'ctl00$ContentPlaceHolder1$GridPager1$ctl02$ctl00', # CSS Selector of the 'Next Page' button
        **context
    }
   
    next_page = session.post(ROOT_URL, data=context)
    
    return next_page


def parse_candidate_soup(soup):
    """
    Given the page of a candidate parsed by BeautifulSoup, extract the relevant fields.

    Return as a list.
    """
    
    # Find all of the relevant fields
    candidate_id = soup.find(id='ctl00_ContentPlaceHolder1_lblCandID').text
    name = soup.find(id='ctl00_ContentPlaceHolder1_VRSHeading1').text.replace('Candidate Information -', '')
    
    office = soup.find(id='ctl00_ContentPlaceHolder1_lblOffice').text
    district = soup.find(id='ctl00_ContentPlaceHolder1_lblDistrict').text
    party = soup.find(id='ctl00_ContentPlaceHolder1_lblParty').text
    
    address = soup.find(id='ctl00_ContentPlaceHolder1_tabs_TabPanel1_lblMailingAddress').text
    email = soup.find(id='ctl00_ContentPlaceHolder1_tabs_TabPanel1_lblEmail').text
    phone = soup.find(id='ctl00_ContentPlaceHolder1_tabs_TabPanel1_lblPhone').text
    
    municipality = soup.find(id='ctl00_ContentPlaceHolder1_tabs_TabPanel1_lblMunicipality').text
    county = soup.find(id='ctl00_ContentPlaceHolder1_tabs_TabPanel1_lblCounty').text

    result = [candidate_id, name, office, district, party, address, email, phone, municipality, county]
    
    return result
    
    
def save_as_excel(all_candidates):
    """
    Save the list of candidate_info to an Excel file
    """
    wb = Workbook()
    sheet = wb.active
    row = 1
    
    for candidate in all_candidates:
        for col, val in enumerate(candidate, start=1):
            sheet.cell(row, col).value = val
        row += 1
            
    file_name = f"candidates_{str(date.today()).replace('-', '')}.xlsx"
    
    print(f"Saving to file '{file_name}'")
            
    wb.save(file_name)

if __name__ == '__main__':
    main()
