import pandas
from datetime import datetime


def fetch_box_office_results():
    current_year = datetime.now().year
    url = "https://www.boxofficemojo.com/year/{}/?releaseScale=all&grossesOption=totalGrosses"
    print("Polling movies from {}".format(current_year))
    df = pandas.read_html(url.format(current_year))[0]
    df['Year'] = current_year
    current_year -= 1

    while current_year >= 1977:
        print("Polling movies from {}".format(current_year))
        thispagedf = pandas.read_html(url.format(current_year))[0]
        thispagedf['Year'] = current_year
        df = pandas.concat([thispagedf,df])
        # df = df.append(thispagedf, ignore_index=True)
        current_year -= 1

    return df


def format_box_office_results(df):
    # TODO trim feature name column to 83 character max, because 83 was all Borat needed for it's full, joke title
    # TODO genre, budget, and running time columns aren't polling?
    # TODO assert data typing for numbers, percents, dates
    # TODO write proper, descriptive headers
    # TODO figure out the semantic meaning of the 'Estimated column'
    return df


def write_excel_report(reportframe):
    with pandas.ExcelWriter(datetime.now().strftime("%Y-%m-%d") + " BoxOfficeMojo Summary.xlsx") as writer:
        reportframe.to_excel(writer, sheet_name='Summary')
    print("Report saved as " + datetime.now().strftime("%Y-%m-%d") + " BoxOfficeMojo Summary.xlsx")
    return


if __name__ == '__main__':
    boxdf = format_box_office_results(fetch_box_office_results())
    write_excel_report(boxdf)
