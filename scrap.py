from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time

class Scraper():
    URL = "https://www.senado.cl/appsenado/index.php?mo=sesionessala&ac=votacionSala&legiini=462"

    def __init__(self):
        self.chrome_options = webdriver.ChromeOptions()
        #chrome_options.add_argument('--headless')
        self.chrome_options.add_argument('--no-sandbox')
        self.chrome_options.add_argument('--disable-dev-shm-usage')
        self.chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
        self.driver = webdriver.Chrome(ChromeDriverManager().install(),chrome_options=self.chrome_options)
        self.driver.set_page_load_timeout(30)
        self.wait = WebDriverWait(self.driver, 30)
        self.driver.get(self.URL)
        self.output = []

    def grabSessions(self):
        numOptions = len(self.driver.find_elements(By.XPATH, "//select[@name='sesionessala']/option"))
        i = 2
        self.links = {}
        while i < numOptions:
            x = 0
            self.driver.find_element(By.XPATH, f"//select[@name='sesionessala']/option[{i}]").click()
            time.sleep(5)
            try: 
                for _ in self.driver.find_elements(By.XPATH, "//table[@class='clase_tabla']//tr"):
                    links = self.driver.find_elements(By.XPATH, "//table[@class='clase_tabla']//tr//a")
                    link = links[x].get_attribute("href")
                    detalles = [detalle.text for detalle in self.driver.find_elements(By.XPATH, "//table[@class='clase_tabla']//tr//td") if "TEMA" in detalle.text]
                    detalle = detalles[x].split('\n')[0][6:]
                    tds =  self.driver.find_elements(By.XPATH, "//table[@class='clase_tabla']//tr/td")
                    dates = [td.text for td in tds if "/" in td.text]
                    date = dates[x]
                    try: 
                        boletins = [td.text for td in tds if "-" in td.text and len(td.text) < 20]
                        boletin = boletins[x]
                    except: boletin = ''
                    self.links[link] = [date, boletin, detalle]
                    x+=1
            except: pass
            i += 1

    def grabInfo(self):
        for link in list(self.links.keys()):
            self.driver.get(link)
            txt = self.driver.find_element(By.CLASS_NAME, "col1").text
            session = txt.split('/')[0][8:-1]
            date = self.links[link][0]
            tema = self.links[link][2]
            boletin = self.links[link][1]
            for tr in self.driver.find_elements(By.XPATH, "//tbody/tr"):
                if "Senador" in tr.text:
                    continue
                senador = tr.text
                vote = None
                try:
                    if tr.find_element(By.XPATH, ".//td[2]/img").get_attribute("src") != None:
                        vote = "1 - Si"
                except: pass
                try:
                    if tr.find_element(By.XPATH, ".//td[3]/img").get_attribute("src") != None:
                        vote = "2 - No"
                except: pass         
                try:
                    if tr.find_element(By.XPATH, ".//td[4]/img").get_attribute("src") != None:
                        vote = "3 - Abs"  
                except: pass
                try:
                    if tr.find_element(By.XPATH, ".//td[5]/img").get_attribute("src") != None:
                        vote = "4 - Par"
                except: pass
                if vote == None: 
                    vote = "5 - Sin info"
                    
                #print({"Sesión":session, "Fecha/Hora":date, "Boletin":boletin, "Detalle":tema, "Nombre":senador, "Votación":vote})
                self.output.append({"Sesión":session, "Fecha/Hora":date, "Boletin":boletin, "Detalle":tema, "Nombre":senador, "Votación":vote})


    def saveOutput(self):
        df = pd.DataFrame(self.output)
        writer = pd.ExcelWriter('Senado.xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name='output', index=False)
        writer.save()


    def scrap(self):
        self.grabSessions()
        self.grabInfo()
        self.saveOutput()

s = Scraper()
s.scrap()
