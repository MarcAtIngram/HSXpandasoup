from pandas import DataFrame, ExcelWriter
import requests
from bs4 import BeautifulSoup
from datetime import datetime


def excel_date(date1):
    """
    Returns Excel serialized date as float
        Parameters:
            date1 (datetime): a Python datetime object

        Returns:
            serial (float): serialization of datetime for Excel date format
    """
    temp = datetime(1899, 12, 30)    # Note, not 31st Dec but 30th!
    delta = date1 - temp
    serial = float(delta.days) + (float(delta.seconds) / 86400)
    return serial


def get_all_prices():
    """
    Returns Pandas dataframe of HSX symbols and their current box office earnings estimates
        Parameters:

        Returns:
            df (pandas.DataFrame): Box office estimates and daily deltas keyed by their HSX symbol
    """
    page_count = 1
    curr_page = 1
    films = {}

    while curr_page <= page_count:
        r = requests.get('http://www.hsx.com/security/list.php?id=1&sfield=name&sdir=asc&page={}'.format(curr_page))
        soup = BeautifulSoup(r.text, 'lxml')
        if page_count == 1:
            page_count = int(soup.text[soup.text.find('Page 1 of')+9:soup.text.find('Page 1 of')+12])
            print('Scraping {} pages from hsx.com/security/list.php'.format(page_count))
        film_list = soup.find('tbody').findAll('tr')[1:]
        for film in film_list:
            film = film.text.strip().split('\n')
            movement = film[3].replace('(', '').replace(')', '').split('\xa0')
            films[film[1]] = (film[0], film[2], movement[0], movement[1])
            # print('{0}: {1}'.format(film[1], films[film[1]]))
        curr_page += 1

    df = DataFrame(films, index=['Name', 'PriceToday', 'ChangeToday', 'ChangePercent']).T
    df['PriceToday'] = df['PriceToday'].map(lambda x: float(x.replace('H$', ''))*1000000)
    df['ChangeToday'] = df['ChangeToday'].map(lambda x: float(x.replace('H$', ''))*1000000)
    return df


def get_all_release_dates():
    """
    Returns Pandas dataframe of HSX symbols and their estimated theatrical release date
        Parameters:

        Returns:
            df (pandas.DataFrame): Theatrical release dates keyed by their HSX symbol
    """
    page_count = 1
    curr_page = 1
    films = {}

    while curr_page <= page_count:
        r = requests.get('https://www.hsx.com/security/feature.php?type=upcoming&page={}'.format(curr_page))
        soup = BeautifulSoup(r.text, 'lxml')
        if page_count == 1:
            page_count = int(soup.text[soup.text.find('Page 1 of')+9:soup.text.find('Page 1 of')+12])
            print('Scraping {} pages from hsx.com/security/feature.php'.format(page_count))
        film_list = soup.find('tbody').findAll('tr')[1:]
        for film in film_list:
            film = film.text.strip().split('\n')
            film[2] = excel_date(datetime.strptime(film[2][0:12], '%b %d, %Y'))
            films[film[1]] = (film[2])
        curr_page += 1

    df = DataFrame(films, index=['ReleaseDate']).T
    return df


if __name__ == '__main__':
    ''' scrape our data into pandas dataframes '''
    priceFrame = get_all_prices()
    dateFrame = get_all_release_dates()
    summaryFrame = priceFrame.join(other=dateFrame, how='left', rsuffix='_price', lsuffix='_date')

    ''' export those frames into some suitable Excel format '''
    with ExcelWriter(datetime.now().strftime("%Y-%m-%d") + ' HSX Summary.xlsx') as writer:
        summaryFrame.to_excel(writer, sheet_name='Summary')
        dateFrame.to_excel(writer, sheet_name='Release Dates')
        priceFrame.to_excel(writer, sheet_name='Movie Stocks')
    print("Done!")
