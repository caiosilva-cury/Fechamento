from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# baixa o driver mais recente e mostra o caminho
driver_path = ChromeDriverManager().install()
print(f"Driver baixado em: {driver_path}")

# configurações do Chrome
#options = Options()
#options.add_argument("--start-maximized")

# inicia o navegador
#driver = webdriver.Chrome(
 #   service=Service(driver_path),
 #   options=options
#)

#driver.get("https://www.google.com")
