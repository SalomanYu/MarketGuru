import config

from oauth2client.service_account import ServiceAccountCredentials
import gspread
import xlsxwriter, xlrd
import json
import os
from time import sleep

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from sys import platform


if platform == 'win32':
    import ctypes
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)

SUCCESS_MESSAGE = '\033[2;30;42m [SUCCESS] \033[0;0m' 
WARINING_MESSAGE = '\033[2;30;43m [WARNING] \033[0;0m'
ERROR_MESSAGE = '\033[2;30;41m [ ERROR ] \033[0;0m'


class GoogleSheet:
    def __init__(self, googlesheet_id):
        self.googlesheet_id = googlesheet_id 

    def auth_spread(self):
        """
        Подключение к GoogleSheet API
        OUTPUT: Таблица, готовая к работе
        """
        scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('Service Accounts/morbot-338716-b219142d9c70.json', scope)

        gc = gspread.authorize(credentials)
        spread = gc.open_by_key(self.googlesheet_id)

        return spread

    def download_sheet(self, worksheet):
        """
        Для снижения нагрузки на API и ускорения работы программы скачиваем содержимое таблицы
        INPUT: Конкретная страница таблицы
        OUTPUT: None
        """
        os.makedirs('FILES', exist_ok=True)
        filename = 'FILES/Dowloaded sheet' + '.xlsx'

        workbook = xlsxwriter.Workbook(filename)
        sheet = workbook.add_worksheet()
        all_values = worksheet.get_all_values()
        for row_num,row_data in enumerate(all_values):
            for col_num, col_data in enumerate(row_data):
                sheet.write(row_num, col_num, col_data)

        workbook.close()
    
    def save_articles_to_json(self):
        book = xlrd.open_workbook('FILES/Dowloaded sheet.xlsx')
        sheet = book.sheet_by_index(0)
        table_titles = sheet.row_values(0)

        for title_num in range(len(table_titles)):
            if table_titles[title_num] == 'Наша цена АЛ':
                article_col = title_num
                articles_col_values = sheet.col_values(article_col)[1:]
                break

        result = {} # Делаем словарь 'наш артикул': [список артикулов конкурентов]
        for row in range(len(articles_col_values)): # Тут надо прибавлять еще 2 
            if articles_col_values[row].isdigit() and len(articles_col_values[row]) == 8:
                concurent_articles = []
                col = article_col + 2
                concurent = sheet.cell(row+1, col).value
                
                while concurent:
                    if concurent.isdigit() and len(concurent) == 8:
                        concurent_articles.append(concurent)
                        col += 1
                        concurent = sheet.cell(row+1, col).value
                    else:
                        break

                # if concurent_articles:
                result[articles_col_values[row]] = concurent_articles
        
        self.save_to_json(result, 'Articles for MarketGuru search')

    def save_to_json(self, data, filename):
        with open(f'FILES/{filename}.json', 'w') as file:
            json.dump(data, file, ensure_ascii=False, indent=2)


class Guru:
    def __init__(self):
        self.auth()

    def auth(self):
        options = Options()
        options.add_argument("--headless") # ФОНОВЫЙ РЕЖИМ
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

        self.browser = webdriver.Chrome(options=options)
        self.browser.get('https://my.marketguru.io/auth/signin')
        sleep(2)
        email_button = self.browser.find_element(By.XPATH, "//div[@class='mail-wrap']//button")
        email_button.click()

        sleep(2)

        WebDriverWait(self.browser, 30).until(EC.element_to_be_clickable((By.XPATH, '//input[contains(@formcontrolname, "email")]'))).send_keys(config.LOGIN + Keys.ENTER)
        sleep(1)
        WebDriverWait(self.browser, 30).until(EC.element_to_be_clickable((By.XPATH, '//input[contains(@formcontrolname, "password")]'))).send_keys(config.PASSWORD + Keys.ENTER)

        sleep(3)
        
    def get_price_with_sales(self, article) -> tuple:
        
        self.browser.get(f'https://my.marketguru.io/wb/competitors/nomenclature/{article.strip()}/info')
        self.browser.implicitly_wait(10)

        try:
            first_sales = int(self.browser.find_element(By.XPATH, "(//div[@class='d-flex flex-column h-100'])[2]//div[@class='widget-item'][2]").text.split('шт')[0].replace(' ', ''))
        except:
            print(ERROR_MESSAGE+'Не удалось прочитать артикул ', article)
            first_sales = 0
            # first_sales = int(self.browser.find_element(By.XPATH, "(//div[@class='d-flex flex-column h-100'])[2]//div[@class='widget-item'][2]").text.split('шт')[0].replace(' ', ''))
        price = self.browser.find_element(By.XPATH, "//div[@class='d-flex flex-column h-100']//div[@class='widget-item'][3]//b").text

        
        self.browser.get(f'https://my.marketguru.io/wb/competitors/nomenclature/{article}/history') # Берем вторую инфу о продажах
        self.browser.implicitly_wait(10)


        self.browser.find_element(By.XPATH, "//div[@class='switch-type__switcher']").click()
        self.browser.implicitly_wait(10)
        try:
            second_sales = int(self.browser.find_element(By.XPATH, "(//div[@class='overflow-auto']//tr)[last()]//td[2]").text.replace(' ', ''))
            final_sales = round((first_sales+second_sales) / 2)

        except:
            return price, first_sales
        if first_sales != 0:
            return price, final_sales
        else:
            return price, second_sales

    def quit(self):
        self.browser.close()
        self.browser.quit()

def find_articles():
    """
    Запросы к marketGuru
    """
    articles = json.load(open('FILES/Articles for MarketGuru search.json'))
    try:
        market = Guru()
    except BaseException as error:
        print(ERROR_MESSAGE+ f'\t{error}')
        sleep(10)
        market = Guru()

    data = {}
    for item in articles:
        if len(item) == 8:
            price, sales = market.get_price_with_sales(item)
            data[item] = {
            'price': price,
            'sales': sales,
            'concurents': []
            }

        for concurent_article in articles[item]:
            try:
                price, sales = market.get_price_with_sales(concurent_article)
            except BaseException as error:
                print(ERROR_MESSAGE+ f'\t{error}', concurent_article)
                sleep(10)
                price, sales = market.get_price_with_sales(concurent_article)
            data[item]['concurents'].append({
                concurent_article:{
                    'price': price,
                    'sales': sales
                }
            })
    market.quit()
    return data

def update_table(worksheet):

    """
    Изменение данных в googlesheet
    """
    data = json.load(open('FILES/Articles data.json', 'r'))

    def update_article_cell(article):
        try:
            similar_articles = worksheet.findall(article)
            for item in similar_articles:
                row_article = item.row
                col_article = item.col
                worksheet.update_cell(row_article+1, col_article, data[article]['price'])  
                worksheet.update_cell(row_article+2, col_article, data[article]['sales']) 
        except gspread.exceptions.APIError:
            sleep(20)
            print('Превышение числа запросов')
            update_article_cell(article)

    def update_concurent_cell(concurent):
        try:
            key = list(concurent.keys())[0]
            similar_concurents = worksheet.findall(key)
            for item in similar_concurents:
                col_concurent = item.col
                row_concurent = item.row 
                worksheet.update_cell(row_concurent+1, col_concurent, concurent[key]['price'])
                worksheet.update_cell(row_concurent+2, col_concurent, concurent[key]['sales'])

        except gspread.exceptions.APIError:
            print('Превышение числа запросов')
            sleep(20)
            update_concurent_cell(concurent)


    for item in data:
        update_article_cell(item)
        for concurent in data[item]['concurents']:
            update_concurent_cell(concurent)

if __name__ == "__main__":
    """
    Сохранение вводных данных из таблицы
    """
    google = GoogleSheet(googlesheet_id='1lzGdBW4KJ6Tv6BOe8K3SuXCDyqZ-vsqSaiFf-pl1ZT4')  # Main
    spread = google.auth_spread()

    print(SUCCESS_MESSAGE, '\tПодключились к гугл таблицам')

    worksheet = spread.get_worksheet(0) 
    google.download_sheet(worksheet)

    print(SUCCESS_MESSAGE, '\tСкачали таблицу')

    google.save_articles_to_json()

    print(SUCCESS_MESSAGE, '\tПодготовили файл с артикулами для MarketGuru')

    data = find_articles()
    google.save_to_json(data, 'Articles data')

    print(SUCCESS_MESSAGE, '\tЗакончили искать информацию о артикулах')

    print(WARINING_MESSAGE, '\tОбновляем таблицу')

    update_table(worksheet)

    print(SUCCESS_MESSAGE, '\tЗакончили обновлять таблицу')
