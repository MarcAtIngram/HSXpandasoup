from bs4 import BeautifulSoup
from datetime import datetime
import pandas as pd
import requests
import time


def get_hsx_prices():
    """
    Fetches HSX symbols and prices into a Pandas DataFrame
    Args:

    Returns:
        symbol_frame (pandas.DataFrame): A list of HSX symbols
    """
    current_page = 1
    url = 'http://www.hsx.com/security/list.php?id=1&sfield=name&sdir=asc&page={}'

    r = requests.get(url.format(current_page))
    soup = BeautifulSoup(r.text, 'lxml')
    page_count = int(soup.text[soup.text.find('Page 1 of') + 9:soup.text.find('Page 1 of') + 12])

    print('Scraping {} box office pages'.format(page_count))

    symbol_frame = pd.read_html(url.format(current_page))[0]
    current_page = current_page + 1

    while current_page <= page_count:
        thispagedf = pd.read_html(url.format(current_page))[0]
        symbol_frame = pd.concat([symbol_frame, thispagedf])
        current_page = current_page + 1

    return symbol_frame


def format_prices(symbol_frame):
    """
    Formats the output of get_hsx_prices() for export in Excel
    Args:
        symbol_frame (pandas.DataFrame): An unformatted scrape of HSX symbols and their current prices
    Returns:
        formatted_symbol_frame (pandas.DataFrame): A formatted table of HSX symbols and their current prices
    """
    formatted_symbol_frame = symbol_frame
    formatted_symbol_frame.columns = ['Name', 'Symbol', 'Price', 'Change', 'Button']
    formatted_symbol_frame.set_index('Symbol', inplace=True)
    formatted_symbol_frame.drop(['Change', 'Button'], axis=1, inplace=True)
    formatted_symbol_frame['Price'] = symbol_frame['Price'].map(lambda x: float(x.replace('H$', '')) * 1000000)

    return formatted_symbol_frame


def get_hsx_release_dates():
    """
    Fetches HSX symbols and their estimated theatrical release date into a Pandas DataFrame
    Args:

    Returns:
        release_date_frame (pandas.DataFrame): Theatrical release dates keyed by their HSX symbol
    """

    # scrape the first page
    current_page = 1
    url = 'https://www.hsx.com/security/feature.php?type=upcoming&page={}'
    r = requests.get(url.format(current_page))
    soup = BeautifulSoup(r.text, 'lxml')
    release_date_frame = pd.read_html(url.format(current_page))[0]

    # capture the number of pages to be scraped
    page_count = int(soup.text[soup.text.find('Page 1 of') + 9:soup.text.find('Page 1 of') + 12])
    print('Scraping {} release date pages '.format(page_count))

    # scrape all pages after the first
    current_page = current_page + 1

    while current_page <= page_count:
        this_page_df = pd.read_html(url.format(current_page))[0]
        release_date_frame = pd.concat([release_date_frame, this_page_df])
        current_page = current_page + 1

    return release_date_frame


def format_release_dates(release_date_frame):
    """
    Formats the output of get_hsx_release_dates() for export in Excel
    Args:
        release_date_frame (pandas.DataFrame): A messy scrape of theatrical release dates, keyed by their HSX symbol
    Returns:
        formatted_date_frame (pandas.DataFrame): A tidy dataframe of Excel-formatted theatrical release dates, keyed by
        their HSX symbol
    """
    formatted_date_frame = release_date_frame

    formatted_date_frame.columns = ['Name', 'Symbol', 'TheatricalRelease', 'Price', 'Change', 'Button']
    formatted_date_frame.set_index('Symbol', inplace=True)
    formatted_date_frame.drop(['Name', 'Price', 'Change', 'Button'], axis=1, inplace=True)
    formatted_date_frame['TheatricalRelease'] = pd.to_datetime(formatted_date_frame['TheatricalRelease'])
    formatted_date_frame['TheatricalRelease'] = formatted_date_frame['TheatricalRelease'].apply(lambda x:
                                                                                                format_excel_date(x))

    return formatted_date_frame


def get_symbol_detail(symbol_frame):
    """
    Fetches Movie details from each security's page based on an inputted list of movie Symbols
        Args:
            symbol_frame (pandas.DataFrame): A list of HSX movie security Symbols
        Returns:
            symbol_detail_frame (pandas.DataFrame): A table of IPO attributes, indexed by symbol
    """
    url = "https://www.hsx.com/security/view/{}"
    row_counter = 1
    row_count = len(symbol_frame.index)

    symbol_detail_frame = pd.DataFrame(data=None, columns=['Symbol', 'TheatricalRelease', 'HomeRelease', 'Status',
                                                           'IPODate', 'Genre', 'MPAARating', 'Phase', 'Theaters',
                                                           'ReleasePattern', 'Distributor', 'Director', 'Cast'])
    symbol_detail_frame.set_index('Symbol', inplace=True)

    for symbol in symbol_frame.index:
        time.sleep(2)
        symbol_url = url.format(symbol)

        # pandas process
        this_page_pandas = pd.read_html(symbol_url)
        unioned_pandas = pd.concat([this_page_pandas[0], this_page_pandas[1]], ignore_index=True)
        unioned_pandas.columns = ['key', 'value']
        unioned_pandas.set_index('key', inplace=True)

        if "Symbol:" in unioned_pandas.index:
            symbol = unioned_pandas.loc['Symbol:']['value']

        if "Status:" in unioned_pandas.index:
            symbol_detail_frame.at[symbol, 'Status'] = unioned_pandas.loc['Status:']['value']
        if "IPO\xa0Date:" in unioned_pandas.index:
            symbol_detail_frame.at[symbol, 'IPODate'] = unioned_pandas.loc['IPO\xa0Date:']['value']
        if "Genre:" in unioned_pandas.index:
            symbol_detail_frame.at[symbol, 'Genre'] = unioned_pandas.loc['Genre:']['value']
        if "MPAA Rating:" in unioned_pandas.index:
            symbol_detail_frame.at[symbol, 'MPAARating'] = unioned_pandas.loc['MPAA Rating:']['value']
        if "Phase:" in unioned_pandas.index:
            symbol_detail_frame.at[symbol, 'Phase'] = unioned_pandas.loc['Phase:']['value']
        if "Theaters:" in unioned_pandas.index:
            symbol_detail_frame.at[symbol, 'Theaters'] = unioned_pandas.loc['Theaters:']['value']
        if "Release\xa0Pattern:" in unioned_pandas.index:
            symbol_detail_frame.at[symbol, 'ReleasePattern'] = unioned_pandas.loc['Release\xa0Pattern:']['value']

        # beautiful soup process
        this_page_requests = requests.get(symbol_url)
        this_page_soup = BeautifulSoup(this_page_requests.text, 'html.parser')

        for s in this_page_soup.find_all(string="Distributor"):
            symbol_detail_frame.at[symbol, 'Distributor'] = s.findParent().findNextSibling().get_text()

        for s in this_page_soup.find_all(string="Director"):
            symbol_detail_frame.at[symbol, 'Director'] = s.findParent().findNextSibling().get_text()

        for s in this_page_soup.find_all(string="Cast"):
            symbol_detail_frame.at[symbol, 'Cast'] = s.findParent().findNextSibling().get_text()

        if row_counter % 10 == 0:
            print('Scraping symbol {} of {}'.format(row_counter, row_count))

        row_counter += 1

    return symbol_detail_frame


def format_combined_frames(formatted_prices, symbol_detail_frame, formatted_date_frame):
    """
    Merges pricing, release date, and movie detail frames together into a single DataFrame
    Args:
        formatted_prices (pandas.DataFrame):
        symbol_detail_frame (pandas.DataFrame):
        formatted_date_frame (pandas.DataFrame):

    Returns:
        combined_frame (pandas.DataFrame):
    """
    combined_frame = pd.merge(formatted_prices, symbol_detail_frame, left_index=True, right_index=True)
    combined_frame = combined_frame.merge(formatted_date_frame, on='Symbol', how='left')
    combined_frame['TheatricalRelease'] = \
        combined_frame['TheatricalRelease_y'].fillna(combined_frame['TheatricalRelease_x'])
    combined_frame.drop(['TheatricalRelease_x', 'TheatricalRelease_y'], axis=1)
    combined_frame = combined_frame[['Name', 'Price', 'TheatricalRelease', 'HomeRelease', 'Status', 'IPODate', 'Genre',
                                     'MPAARating', 'Phase', 'Theaters', 'ReleasePattern', 'Distributor', 'Director',
                                     'Cast']]
    return combined_frame


def format_excel_date(date):
    """
    Converts Python datetime object into a serialized Excel date
    Args:
        date (datetime): a Python datetime object
    Returns:
        serial (float): serialization of datetime for Excel date format
    """
    temp = datetime(1899, 12, 30)    # Note, not 31st Dec but 30th!
    delta = date - temp
    serial = float(delta.days) + (float(delta.seconds) / 86400)
    return serial


def write_excel_report(symbol_detail_frame):
    """
    Exports finalized dataframe into an Excel spreadsheet
    Args:
        symbol_detail_frame (pandas.DataFrame):
    Returns:
        none
    """
    with pd.ExcelWriter(datetime.now().strftime("%Y-%m-%d %H%M") + ' HSX Symbol detail beta.xlsx') as writer:
        symbol_detail_frame.to_excel(writer, sheet_name='symbol_details')
    print("Report saved as " + datetime.now().strftime("%Y-%m-%d %H%M") + " HSX Symbol detail beta.xlsx")
    return


def main():
    symbol_frame = get_hsx_prices()

    formatted_prices = format_prices(symbol_frame)

    release_date_frame = get_hsx_release_dates()

    formatted_date_frame = format_release_dates(release_date_frame)

    symbol_detail_frame = get_symbol_detail(symbol_frame)

    combined_frame = format_combined_frames(formatted_prices, symbol_detail_frame, formatted_date_frame)

    write_excel_report(combined_frame)


main()
