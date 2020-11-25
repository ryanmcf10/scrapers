import json
from urllib.parse import quote

import requests
from openpyxl import Workbook


def main():
    precinct_names = get_precinct_names()
    contest_titles = ["Attorney General", "Auditor General", "Lower Frederick Township Question",
                      "Presidential Electors", "Representative in Congress 1st District",
                      "Representative in Congress 4th District", "Representative in Congress 5th District",
                      "Representative in the General Assembly 131st District",
                      "Representative in the General Assembly 146th District",
                      "Representative in the General Assembly 147th District",
                      "Representative in the General Assembly 148th District",
                      "Representative in the General Assembly 149th District",
                      "Representative in the General Assembly 150th District",
                      "Representative in the General Assembly 151st District",
                      "Representative in the General Assembly 152nd District",
                      "Representative in the General Assembly 153rd District",
                      "Representative in the General Assembly 154th District",
                      "Representative in the General Assembly 157th District",
                      "Representative in the General Assembly 166th District",
                      "Representative in the General Assembly 172nd District",
                      "Representative in the General Assembly 194th District",
                      "Representative in the General Assembly 26th District",
                      "Representative in the General Assembly 53rd District",
                      "Representative in the General Assembly 61st District",
                      "Representative in the General Assembly 70th District",
                      "Senator in the General Assembly 17th District", "Senator in the General Assembly 7th District",
                      "State Treasurer"]

    wb = Workbook()
    sheet = wb.active
    row = 1

    for contest_title in contest_titles:
        for precinct_name in precinct_names:
            url = create_url(precinct_name, contest_title)
            request = requests.get(url)

            status = request.status_code

            if status == 200:
                data = json.loads(request.content.decode('utf-8'))
                candidates = data['features']

                for candidate in candidates:
                    attrs = candidate['attributes']

                    votes = attrs['value']
                    candidate_name = attrs['candidate_name']
                    party_code = attrs['Party_Code']

                    current_row = [contest_title, precinct_name, candidate_name, party_code, votes]
                    print(row, current_row)

                    for col, val in enumerate(current_row, start=1):
                        sheet.cell(row, col).value = val
                    row += 1

    wb.save('montco_precinct_results.xlsx')

def get_precinct_names() -> list:
    """
    Parse the JSON response from this URL that contains all of the Precinct Names
    :return: list of precinct names
    """
    precinct_names = []

    url = "https://services1.arcgis.com/kOChldNuKsox8qZD/arcgis/rest/services/ElectionResults_GE20_dashboard/FeatureServer/1/query?f=json&where=Contest_title%3D%27Attorney%20General%27&returnGeometry=false&spatialRel=esriSpatialRelIntersects&outFields=*&groupByFieldsForStatistics=Precinct_Sort&orderByFields=Precinct_Sort%20asc&outStatistics=%5B%7B%22statisticType%22%3A%22count%22%2C%22onStatisticField%22%3A%22Precinct_Sort%22%2C%22outStatisticFieldName%22%3A%22count_result%22%7D%5D&resultType=standard&cacheHint=true"
    request = requests.get(url)
    data = json.loads(request.content.decode('utf-8'))

    for precinct in data['features']:
        attrs = precinct['attributes']
        precinct_names.append(attrs['Precinct_Sort'])
        
    return precinct_names


def create_url(precinct_name: str, contest_title: str) -> str:
    """
    Create the URL for getting the Precinct Levels result for a given precinct and contest
    :param precinct_name: str
    :param contest_title: str
    :return: str
    """
    return f"https://services1.arcgis.com/kOChldNuKsox8qZD/arcgis/rest/services/ElectionResults_GE20_dashboard/FeatureServer/1/query?f=json&where=(Vote_Type%3D%27Total%20Votes%27)%20AND%20(Precinct_Sort%3D%27{quote(precinct_name)}%27)%20AND%20(Contest_title%3D%27{quote(contest_title)}%27)&returnGeometry=false&spatialRel=esriSpatialRelIntersects&outFields=*&groupByFieldsForStatistics=candidate_name%2CParty_Code&orderByFields=value%20desc&outStatistics=%5B%7B%22statisticType%22%3A%22sum%22%2C%22onStatisticField%22%3A%22Votes%22%2C%22outStatisticFieldName%22%3A%22value%22%7D%5D&resultType=standard&cacheHint=true"


if __name__ == '__main__':
    main()