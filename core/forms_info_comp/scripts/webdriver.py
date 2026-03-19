import os
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from dotenv import load_dotenv

load_dotenv()
ROOT = os.getenv("ROOT")

# Caminho do driver local
EDGE_DRIVER_PATH = os.path.join(ROOT, 'drivers', 'msedgedriver.exe')

def configurar_webdriver(pasta_download=None):
    options = webdriver.EdgeOptions()
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('start-maximized')
    options.add_argument('disable-infobars')
    options.add_argument('--disable-extensions')
    options.add_argument('log-level=3')
    options.add_argument('--remote-debugging-port=0')
    options.add_argument("--incognito")  # modo anônimo
    options.add_argument("--disable-cache")
    # options.add_argument("--headless=new")

    if pasta_download is None:
        pasta_download = os.path.join(os.getcwd(), "downloads")

    os.makedirs(pasta_download, exist_ok=True)

    prefs = {
        "download.default_directory": os.path.abspath(pasta_download),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)

    # usa o driver local
    edge_service = EdgeService(executable_path=EDGE_DRIVER_PATH)
    driver = webdriver.Edge(service=edge_service, options=options)

    return driver
