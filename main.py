from bs4 import BeautifulSoup
from lxml import etree
import requests
import csv
from typing import List, Dict
import yfinance as yf


def scrape_html(url: str):
    """
    Download the html page and transform it into a dom.
    param url
    return: The dom of the html page
    """
    source = requests.get(url).text
    soup = BeautifulSoup(source, 'lxml')
    dom = etree.HTML(str(soup))
    return dom


def extract_codes(dom) -> List[str]:
    """
    Extract the ftse100 constituents codes from the summary page.
    param dom
    return: a list of ftse100 constituentscode
    """
    codes = dom.xpath('//*[@id="ftse-index-table"]/table/tbody/tr[*]/td[2]/app-link-or-dash/a/@href')
    unique_code = [i.split('/')[1] for i in codes]
    return unique_code


def extract_summary_page(dom) -> List[str]:
    """
    Extract the urls of ftse100 constituents from the summary page.
    param dom
    return: a list of url for ftse100 constituents
    """
    urls = dom.xpath('//*[@id="ftse-index-table"]/table/tbody/tr[*]/td[2]/app-link-or-dash/a/@href')
    url_list = []
    for url in urls:
        summary_page_url = "https://www.londonstockexchange.com/{}/company-page".format(url)
        url_list.append(summary_page_url)
    return url_list


def extract_sector_info(dom, stock_uniquecode) -> Dict[str, str]:
    """
    Extract the sector info from the stock sector page.
    param dom
    param stock_uniquecode
    return: a dict of sector info for ftse100 constituents
    """
    ftse_industry = dom.xpath('//*[@id="ccc-data-ftse-industry"]/div[2]/text()')
    ftse_sector = dom.xpath('//*[@id="ccc-data-ftse-sector"]/div[2]/text()')

    return {
        "stock_uniquecode": stock_uniquecode,
        "ftse_industry": ftse_industry[0].strip(),
        "ftse_sector": ftse_sector[0].strip()
    }


def extract_detailed_page(dom) -> Dict[str, str]:
    """
    Extract the market cap data from the stock detailed page.
    param dom
    return: a dict of market cap information for ftse100 constituents
    """
    market_cap = dom.xpath('//*[@id="chart-table"]/div[3]/div[2]/app-index-item[2]/div/text()')
    stock_code = dom.xpath('//*[@id="ticker"]/div/section/div[1]/div[2]/span[2]/span[1]/text()')

    return {
        "stock_code": stock_code[0],
        "market_cap": market_cap[0].replace(',', '')
    }


def write_to_csv(detailed_info: List[Dict[str, str]], filename: str, headers: List[str]) -> None:
    """
    Write the scraped data into csv with headers.
    param detailed_info: The info that we would like to write into the csv
    param filename: The filename of the csv
    param headers: Headers in the csv
    return:
    """
    with open(filename, mode='w') as csv_file:
        csv_writer = csv.writer(csv_file, delimiter='|')
        
        # header
        csv_writer.writerow(headers)

        for info in detailed_info:
            new_row = []
            for col_name in headers:
                new_row.append(info[col_name])
            csv_writer.writerow(new_row)

   
def get_yahoo_api_data(stock_code: str, start_date: str, end_date: str) -> None:
    """
    Get the historical data for ftse100 constituents between start_date and end_date. And save them into csv files.
    param stock_code
    param start_date
    param end_date
    return:
    """
    if stock_code[-1] == '.':
        stock_code = stock_code[:-1]
    stock_code = stock_code.replace('.', '-')
    new_stock_code = "{}.L".format(stock_code)

    data_df = yf.download(new_stock_code, start=start_date, end=end_date)
    data_df.to_csv('csv/{}.csv'.format(stock_code))


if __name__ == "__main__":

   # If for code testing purpose, comment out page 2 to page 6 for better readability of console output.
    
    pages = [
        "https://www.londonstockexchange.com/indices/ftse-100/constituents/table?lang=en",
        "https://www.londonstockexchange.com/indices/ftse-100/constituents/table?lang=en&page=2",
        "https://www.londonstockexchange.com/indices/ftse-100/constituents/table?lang=en&page=3",
        "https://www.londonstockexchange.com/indices/ftse-100/constituents/table?lang=en&page=4",
        "https://www.londonstockexchange.com/indices/ftse-100/constituents/table?lang=en&page=5",
        "https://www.londonstockexchange.com/indices/ftse-100/constituents/table?lang=en&page=6"
    ]

    print('Scrape stock codes and urls')
    stock_codes = []
    stock_urls = []
    for summary_url in pages:
        print("Scraping {}".format(summary_url))
        dom = scrape_html(summary_url)
        stock_codes += extract_codes(dom)
        stock_urls += extract_summary_page(dom)
    print('Scrape stock codes and urls: Completed.')
    print("Length of stock code list: {}".format(len(stock_codes)))
    print("Length of stock url list: {}".format(len(stock_urls)))

    print('==============')
    print('Go into the detailed page and get the market cap data')
    stock_infos = []  # [ {"market_cap": 123, "stock_code": "xxx"} ]
    for stock_url in stock_urls:
        print("Scraping: " + stock_url)
        detailed_dom = scrape_html(stock_url)
        detailed_info = extract_detailed_page(detailed_dom)  # {"market_cap": 123}
        stock_infos.append(detailed_info)
    print('Scrape stock market cap data: Completed.')
    print("Length of stock info: {}".format(len(stock_infos)))
    
    print('==============')
    print('Go into our-story to get sector info')
    sector_urls = [ i.replace("company-page", "our-story") for i in stock_urls ]
    sector_infos = []
    for sector_url in sector_urls:
        print("Scraping: " + sector_url)
        sector_dom = scrape_html(sector_url)
        stock_uniquecode = sector_url.split('/')[4]
        sector_info = extract_sector_info(sector_dom, stock_uniquecode)
        sector_infos.append(sector_info)
    print('Scrape stock sector info: Completed.')
    print("Length of sector info: {}".format(len(sector_infos)))

    print('==============')
    print("Write the stock info into csv")
    write_to_csv(stock_infos, "stock_info.csv", ["stock_code", "market_cap"])
    print("stock_info.csv is ready!")
    print("Write the sector info into csv")
    write_to_csv(sector_infos, "sector_info.csv", ["stock_uniquecode", "ftse_industry", "ftse_sector"])
    print("sector_info.csv is ready!")

    print('==============')
    print("Scrape historical pricing info")
    for s_code in stock_codes:
        start_date, end_date = '2020-4-23', '2021-4-23'
        print("Get yahoo data for stock code '{}' from {} to {}".format(s_code, start_date, end_date))
        get_yahoo_api_data(s_code, start_date, end_date)
    print('Scrape stock historical pricing info: Completed.')
