import xlrd, validators
from object_extract.config import *

def extract_articles(excel_file):
    """
    Extract links from excel file that is in correct format.
    Correct format: links appear in the second column.
    :param excel_file:
    :return: list of links that appear in the second column of the excel file
    """
    spread_sheet = xlrd.open_workbook(excel_file)
    results = list()
    for sheet in spread_sheet.sheets():
        col = sheet.col(ARTICLE)
        for cell in col:
            link = cell.value
            if validators.url(link):
                article = Article(link)
                article.download()
                article.parse()

                results.append(article)
    return results
