import pygsheets
from datetime import datetime

TABLE_HEAD = ['推送', '更新', '标题', '摘要', '作者']


def format_title_with_hyperlink(title, url):
    return '=HYPERLINK("{}","{}")'.format(url, title)


def format_article_list(article_list):
    rows = []
    for article in article_list:
        # to formatted string from milliseconds
        created_at = datetime.fromtimestamp(article['datetime']).strftime('%m/%d')
        updated_at = datetime.now().strftime("%m/%d %H:%M")

        title = str(article['title'])
        url = str(article['content_url'])
        title = format_title_with_hyperlink(title, url)

        row = [created_at, updated_at, title, str(article['abstract']), str(article['author'])]
        rows.append(row)
    return rows


class GSpreadSheet:
    def __init__(self, credentials_file_name, doc_name):
        # Go to Google Sheets and share your spreadsheet with an email you have in your json_key
        # ['client_email'].
        # Otherwise you’ll get a SpreadsheetNotFound exception when trying to open it.

        # Open spreadsheet and then worksheet
        gc = pygsheets.authorize(service_file=credentials_file_name)
        spreadsheet = gc.open(doc_name)
        worksheets = spreadsheet.worksheets()

        self.spreadsheet = spreadsheet
        self.worksheets = {}

        for worksheet in worksheets:
            self.worksheets[worksheet.title] = worksheet

    def create_new_sheet_for_new_gzh(self, gzh_name):
        new_sheet = self.spreadsheet.add_worksheet(gzh_name, 1, 6)
        new_sheet.insert_rows(0, 1, TABLE_HEAD)
        self.worksheets[gzh_name] = new_sheet

    # gzh_name and worksheet name must match
    def add_new_articles(self, gzh_name, articles):
        rows_of_articles = format_article_list(articles)
        worksheet = self.worksheets[gzh_name]
        worksheet.insert_rows(1, len(articles)-1, rows_of_articles)

    # 1. get cell with the article title
    # 2. update the cell with title with new url hyperlinked to it
    #    and the 'updated_at' cell to its left
    def update_article_url(self, gzh_name, recent_article):
        worksheet = self.worksheets[gzh_name]
        title = recent_article['title']
        cells_list = worksheet.find(title)
        if not cells_list:
            return

        cell = cells_list[0]
        article_row = cell.row
        updated_at = datetime.now().strftime("%m/%d %H:%M")
        hyperlinked_title = format_title_with_hyperlink(title, recent_article['content_url'])
        range_time_and_title = "B{}:C{}".format(article_row, article_row)

        worksheet.update_cells(range_time_and_title, [[updated_at, hyperlinked_title]])
