import xlrd
import validators
from newspaper import Article
from os import path
from object_extract.config import *


def extract(data_root):
    """

    :param data_root:
    :return:
    """
    articles = list()
    for tag in FILES:
        file = data_root + tag
        if path.exists(file):
            articles.append(extract_articles(data_root + tag, tag))
    return articles


def extract_articles(excel_file, tag, nlp=False):
    """
    Extract links from excel file that is in correct format.
    Correct format: links appear in the second column.
    :param excel_file:
    :param tag:
    :param nlp:
    :return:
    """
    spread_sheet = xlrd.open_workbook(excel_file)
    results = list()
    for sheet in spread_sheet.sheets():
        col = sheet.col(ARTICLE)
        counter = 0
        for cell in col:
            if counter == 10:
                break
            link = cell.value
            if validators.url(link):
                article = Article(link)
                try:
                    article.download()
                    article.parse()
                except:
                    continue
                if nlp:
                    article.nlp()
                article.tag = tag
                results.append(article)
                counter += 1
    return results
