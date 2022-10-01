import requests
from bs4 import BeautifulSoup
import pandas
from datetime import datetime


def fetch_prices():
    """
    Fetches HSX symbols and their estimated box office results into a Pandas DataFrame
        Parameters:

        Returns:
            df (pandas.DataFrame): Box office estimates keyed by their HSX symbol
    """
    current_page = 1
    url = 'http://www.hsx.com/security/list.php?id=1&sfield=name&sdir=asc&page={}'

    r = requests.get(url.format(current_page))
    soup = BeautifulSoup(r.text, 'lxml')
    page_count = int(soup.text[soup.text.find('Page 1 of') + 9:soup.text.find('Page 1 of') + 12])

    print('Scraping {} box office pages '.format(page_count))

    df = pandas.read_html(url.format(current_page))[0]
    current_page = current_page + 1

    while current_page <= page_count:
        thispagedf = pandas.read_html(url.format(current_page))[0]
        df = pandas.concat([df, thispagedf])
        current_page = current_page + 1

    return df


def format_prices(df):
    """
    Returns Pandas dataframe of HSX symbols and their estimated box office results
        Parameters:
            df (pandas.Dataframe): A messy scrape of box office estimates, keyed by
            their HSX symbol
        Returns:
            df (pandas.DataFrame): A tidy dataframe of box office estimates, keyed by
            their HSX symbol
    """
    df.columns = ['Name', 'Symbol', 'Price', 'Change', 'Button']
    df.set_index('Symbol', inplace=True)
    df.drop(['Change', 'Button'], axis=1, inplace=True)
    df['Price'] = df['Price'].map(lambda x: float(x.replace('H$', ''))*1000000)

    return df


def fetch_release_dates():
    """
    Fetches HSX symbols and their estimated theatrical release date into a Pandas DataFrame
        Parameters:

        Returns:
            df (pandas.DataFrame): Theatrical release dates keyed by their HSX symbol
    """
    current_page = 1
    url = 'https://www.hsx.com/security/feature.php?type=upcoming&page={}'

    r = requests.get(url.format(current_page))
    soup = BeautifulSoup(r.text, 'lxml')
    page_count = int(soup.text[soup.text.find('Page 1 of') + 9:soup.text.find('Page 1 of') + 12])

    print('Scraping {} release date pages '.format(page_count))

    df = pandas.read_html(url.format(current_page))[0]
    current_page = current_page + 1

    while current_page <= page_count:
        thispagedf = pandas.read_html(url.format(current_page))[0]
        df = pandas.concat([df, thispagedf])
        current_page = current_page + 1

    return df


def format_release_dates(df):
    """
    Returns Pandas dataframe of HSX symbols and their estimated theatrical release date
        Parameters:
            df (pandas.Dataframe): A messy scrape of theatrical release dates, keyed by
            their HSX symbol
        Returns:
            df (pandas.DataFrame): A tidy dataframe of Excel-formatted theatrical
            release dates, keyed by their HSX symbol
    """
    df.columns = ['Name', 'Symbol', 'Release Date', 'Price', 'Change', 'Button']
    df.set_index('Symbol', inplace=True)
    df.drop(['Name', 'Price', 'Change', 'Button'], axis=1, inplace=True)
    df['Release Date'] = pandas.to_datetime(df['Release Date'])
    df['Release Date'] = df['Release Date'].apply(lambda x: excel_date(x))

    return df


def excel_date(date):
    """
    Returns Excel serialized date as float
        Parameters:
            date (datetime): a Python datetime object

        Returns:
            serial (float): serialization of datetime for Excel date format
    """
    temp = datetime(1899, 12, 30)    # Note, not 31st Dec but 30th!
    delta = date - temp
    serial = float(delta.days) + (float(delta.seconds) / 86400)
    return serial


def write_excel_report(reportframe, dateframe, priceframe):
    with pandas.ExcelWriter(datetime.now().strftime("%Y-%m-%d") + ' HSX Summary.xlsx') as writer:
        reportframe.to_excel(writer, sheet_name='Summary')
        dateframe.to_excel(writer, sheet_name='Release Dates')
        priceframe.to_excel(writer, sheet_name='Movie Stocks')
    print("Report saved as " + datetime.now().strftime("%Y-%m-%d") + " HSX Summary.xlsx")
    return


if __name__ == '__main__':
    dateFrame = format_release_dates(fetch_release_dates())
    priceFrame = format_prices(fetch_prices())
    reportFrame = priceFrame.join(other=dateFrame, how='left', rsuffix='_price', lsuffix='_date')
    write_excel_report(reportFrame, dateFrame, priceFrame)
