import customtkinter as ctk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, NoSuchElementException, NoAlertPresentException
from webdriver_manager.chrome import ChromeDriverManager
import threading
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from typing import List, Tuple

# --- Data ---
BOOK_URL_MAP = {
    "2488": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9788968488948/content",
    "4050": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647638/content",
    "4178": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647621/content",
    "4328": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647133/content",
    "4331": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156640684/content",
    "4428": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647263/content",
    "4476": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644767/content",
    "4491": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647737/content",
    "4493": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644934/content",
    "4497": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156640691/content",
    "4519": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648031/content",
    "4520": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647614/content",
    "4530": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648215/content",
    "4532": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648239/content",
    "4534": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648222/content",
    "4535": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648123/content",
    "4536": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647973/content",
    "4537": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648208/content",
    "4538": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648161/content",
    "4539": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648192/content",
    "4540": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648185/content",
    "4541": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648109/content",
    "4542": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648086/content",
    "4543": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648055/content",
    "4544": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648130/content",
    "4545": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648062/content",
    "4546": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648048/content",
    "4547": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648079/content",
    "4548": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647201/content",
    "4549": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647997/content",
    "4550": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648154/content",
    "4551": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648093/content",
    "4552": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648147/content",
    "4553": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648178/content",
    "4554": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648116/content",
    "4556": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648017/content",
    "4557": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648024/content",
    "4558": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647966/content",
    "4559": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647942/content",
    "4560": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647980/content",
    "4565": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156648000/content",
    "4567": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647959/content",
    "4568": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647843/content",
    "4569": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647850/content",
    "4570": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647935/content",
    "4572": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647799/content",
    "4574": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647928/content",
    "4575": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647812/content",
    "4576": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647829/content",
    "4577": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647775/content",
    "4578": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647867/content",
    "4579": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647744/content",
    "4580": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647751/content",
    "4581": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647768/content",
    "4583": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647782/content",
    "4586": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647911/content",
    "4587": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647898/content",
    "4589": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647720/content",
    "4590": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647874/content",
    "4591": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647690/content",
    "4592": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647652/content",
    "4594": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647706/content",
    "4597": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647683/content",
    "4598": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647591/content",
    "4599": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647478/content",
    "4600": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647669/content",
    "4601": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647607/content",
    "4602": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647577/content",
    "4603": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647515/content",
    "4604": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647560/content",
    "4605": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647492/content",
    "4606": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647539/content",
    "4607": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647508/content",
    "4608": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647485/content",
    "4609": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647461/content",
    "4610": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647522/content",
    "4611": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647546/content",
    "4612": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647553/content",
    "4613": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647393/content",
    "4615": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647430/content",
    "4616": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647454/content",
    "4617": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647423/content",
    "4618": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647447/content",
    "4619": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646198/content",
    "4620": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647379/content",
    "4622": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647041/content",
    "4623": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647324/content",
    "4624": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647386/content",
    "4625": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647416/content",
    "4626": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647409/content",
    "4627": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647348/content",
    "4628": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647331/content",
    "4629": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647362/content",
    "4632": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647256/content",
    "4635": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647096/content",
    "4636": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647119/content",
    "4637": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647195/content",
    "4639": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647232/content",
    "4640": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647218/content",
    "4641": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647249/content",
    "4642": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647188/content",
    "4643": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646990/content",
    "4644": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647140/content",
    "4645": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647058/content",
    "4646": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647164/content",
    "4647": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647089/content",
    "4649": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647126/content",
    "4650": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647065/content",
    "4651": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647072/content",
    "4652": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647102/content",
    "4655": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646846/content",
    "4657": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646945/content",
    "4658": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646891/content",
    "4659": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156640998/content",
    "4660": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646877/content",
    "4661": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646983/content",
    "4662": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646969/content",
    "4663": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646921/content",
    "4665": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646839/content",
    "4666": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646914/content",
    "4667": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646952/content",
    "4668": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646853/content",
    "4669": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646938/content",
    "4670": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156640639/content",
    "4671": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646860/content",
    "4672": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646884/content",
    "4673": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156640936/content",
    "4674": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156640820/content",
    "4675": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156646815/content",
    "4677": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156640967/content",
    "4678": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156640943/content",
    "4679": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156640981/content",
    "4680": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156640974/content",
    "4727": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647270/content",
    "4735": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647355/content",
    "4758": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647584/content",
    "4764": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645962/content",
    "4767": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647676/content",
    "4783": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647836/content",
    "4788": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647881/content",
    "4790": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156647904/content",
    "4824": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645313/content",
    "4825": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645290/content",
    "4826": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645115/content",
    "4827": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645122/content",
    "4828": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645085/content",
    "4829": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645153/content",
    "4830": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645160/content",
    "4831": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645108/content",
    "4832": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645238/content",
    "4833": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645252/content",
    "4834": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645283/content",
    "4835": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645245/content",
    "4836": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645061/content",
    "4837": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645030/content",
    "4838": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645078/content",
    "4839": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645092/content",
    "4840": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645221/content",
    "4841": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645214/content",
    "4842": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645047/content",
    "4843": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645139/content",
    "4844": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645054/content",
    "4845": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9788998756802/content",
    "4846": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644941/content",
    "4847": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645146/content",
    "4848": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644859/content",
    "4849": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645009/content",
    "4852": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644965/content",
    "4853": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156645023/content",
    "4855": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644927/content",
    "4856": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644958/content",
    "4857": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644897/content",
    "4859": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644453/content",
    "4861": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644651/content",
    "4862": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644637/content",
    "4863": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644576/content",
    "4866": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642640/content",
    "4868": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642978/content",
    "4869": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9788998756758/content",
    "4870": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642770/content",
    "4875": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644408/content",
    "4876": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643302/content",
    "4878": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644057/content",
    "4880": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644361/content",
    "4881": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644552/content",
    "4883": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643012/content",
    "4884": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643845/content",
    "4885": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644415/content",
    "4886": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644491/content",
    "4889": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644088/content",
    "4890": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644101/content",
    "4891": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644316/content",
    "4892": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644569/content",
    "4894": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156641490/content",
    "4895": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156641773/content",
    "4896": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642367/content",
    "4897": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642602/content",
    "4898": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642404/content",
    "4899": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642053/content",
    "4920": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644736/content",
    "4921": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644392/content",
    "4922": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643883/content",
    "4923": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643722/content",
    "4924": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643715/content",
    "4925": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643654/content",
    "4926": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643203/content",
    "4928": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643029/content",
    "4929": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642374/content",
    "4930": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642633/content",
    "4931": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642589/content",
    "4932": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642596/content",
    "4934": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156641452/content",
    "4936": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644835/content",
    "4939": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644781/content",
    "4940": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644682/content",
    "4943": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644750/content",
    "4944": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644699/content",
    "4945": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644705/content",
    "4946": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644606/content",
    "4947": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644583/content",
    "4948": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644590/content",
    "4949": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644521/content",
    "4950": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644538/content",
    "4951": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644507/content",
    "4952": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644514/content",
    "4953": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644477/content",
    "4954": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644460/content",
    "4955": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644446/content",
    "4956": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644354/content",
    "4957": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643920/content",
    "4958": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643937/content",
    "4959": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643708/content",
    "4960": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643388/content",
    "4961": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644095/content",
    "4962": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644064/content",
    "4963": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643906/content",
    "4964": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643814/content",
    "4965": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643821/content",
    "4966": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643753/content",
    "4968": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643760/content",
    "4969": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643647/content",
    "4970": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643609/content",
    "4971": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643630/content",
    "4973": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156641599/content",
    "4975": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643272/content",
    "4976": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643265/content",
    "4979": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643067/content",
    "4980": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643043/content",
    "4981": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642794/content",
    "4982": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642459/content",
    "4983": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642428/content",
    "4985": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642022/content",
    "4986": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156641889/content",
    "4987": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156641377/content",
    "4988": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156641322/content",
    "4990": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644712/content",
    "4991": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643876/content",
    "4992": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644347/content",
    "4993": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156642541/content",
    "4994": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644828/content",
    "4995": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644743/content",
    "4996": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644811/content",
    "4997": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156641582/content",
    "4998": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156644866/content",
    "4999": "https://play.google.com/books/publish/u/0/a/6028660063035609927#book/ISBN:9791156643913/content",
    "5144": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791185933993/content",
    "5323": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791190846837/content",
    "5338": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791190846790/content",
    "5343": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791190846776/content",
    "5369": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791190846691/content",
    "5370": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791190846707/content",
    "5404": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080924/content",
    "5412": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080795/content",
    "5413": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080818/content",
    "5415": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080832/content",
    "5416": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080825/content",
    "5417": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080740/content",
    "5418": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080771/content",
    "5419": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080764/content",
    "5420": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080801/content",
    "5423": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080788/content",
    "5425": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080757/content",
    "5426": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080733/content",
    "5427": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080726/content",
    "5428": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080702/content",
    "5429": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080719/content",
    "5430": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080696/content",
    "5431": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080689/content",
    "5465": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080658/content",
    "5467": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080672/content",
    "5485": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080856/content",
    "5486": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080863/content",
    "5487": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080870/content",
    "5488": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080887/content",
    "5489": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080894/content",
    "5490": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080900/content",
    "5491": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791193080917/content",
    "7006": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791198140869/content",
    "7007": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791198140876/content",
    "7008": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791198140883/content",
    "10020": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162249895/content",
    "10044": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162249772/content",
    "10061": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162249659/content",
    "10077": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162248805/content",
    "10082": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162249574/content",
    "10093": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162249017/content",
    "10102": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162249437/content",
    "10104": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162249420/content",
    "10126": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162249352/content",
    "10135": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162249239/content",
    "10145": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162249208/content",
    "10166": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162248706/content",
    "10187": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246665/content",
    "10197": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162248478/content",
    "10203": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162248195/content",
    "10242": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162248201/content",
    "10303": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162247600/content",
    "10309": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216104/content",
    "10329": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162247457/content",
    "10336": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217873/content",
    "10346": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162247778/content",
    "10375": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246986/content",
    "10377": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162247204/content",
    "10378": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162247273/content",
    "10385": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162247167/content",
    "10449": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246726/content",
    "10454": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246955/content",
    "10456": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246627/content",
    "10461": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246924/content",
    "10470": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246207/content",
    "10490": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246184/content",
    "10494": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246160/content",
    "10501": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246290/content",
    "10505": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246399/content",
    "10506": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246306/content",
    "10508": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246146/content",
    "10510": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246139/content",
    "10511": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246276/content",
    "10512": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246283/content",
    "10513": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245965/content",
    "10515": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246122/content",
    "10516": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246252/content",
    "10517": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245927/content",
    "10526": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245958/content",
    "10527": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246238/content",
    "10529": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246221/content",
    "10530": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246245/content",
    "10531": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246214/content",
    "10532": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245996/content",
    "10533": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246115/content",
    "10534": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246108/content",
    "10535": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245972/content",
    "10537": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246092/content",
    "10538": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245941/content",
    "10539": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245897/content",
    "10540": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245934/content",
    "10541": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245910/content",
    "10542": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245903/content",
    "10550": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245842/content",
    "10551": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245873/content",
    "10552": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245880/content",
    "10553": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245781/content",
    "10554": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245866/content",
    "10555": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245798/content",
    "10556": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245835/content",
    "10557": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245859/content",
    "10559": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245811/content",
    "10560": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245828/content",
    "10561": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246016/content",
    "10562": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245804/content",
    "10564": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245774/content",
    "10565": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245767/content",
    "10566": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162246023/content",
    "10567": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216029/content",
    "10568": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791162245750/content",
    "10572": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216043/content",
    "10573": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216012/content",
    "10574": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216050/content",
    "11000": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216319/content",
    "11002": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216173/content",
    "11003": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216081/content",
    "11004": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216067/content",
    "11005": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216111/content",
    "11006": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216074/content",
    "11008": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216098/content",
    "11010": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216166/content",
    "11011": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216135/content",
    "11013": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216159/content",
    "11015": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216357/content",
    "11029": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216203/content",
    "11030": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216197/content",
    "11031": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216180/content",
    "11033": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216210/content",
    "11034": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216227/content",
    "11041": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216241/content",
    "11042": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216258/content",
    "11043": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216234/content",
    "11044": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216296/content",
    "11045": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216289/content",
    "11046": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216265/content",
    "11047": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216371/content",
    "11056": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217002/content",
    "11057": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217019/content",
    "11059": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169218054/content",
    "11113": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216999/content",
    "11120": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216982/content",
    "11133": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217712/content",
    "11176": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217637/content",
    "11178": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217729/content",
    "11179": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217644/content",
    "11181": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217828/content",
    "11182": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217651/content",
    "11183": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217668/content",
    "11184": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217675/content",
    "11185": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217699/content",
    "11186": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217620/content",
    "11187": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217682/content",
    "11188": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217705/content",
    "11191": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217743/content",
    "11192": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217781/content",
    "11193": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217798/content",
    "11194": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217804/content",
    "11195": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217750/content",
    "11196": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217736/content",
    "11197": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217835/content",
    "11198": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217767/content",
    "11199": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217774/content",
    "11200": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217811/content",
    "11203": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217910/content",
    "11204": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217941/content",
    "11205": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217927/content",
    "11206": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217842/content",
    "11207": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217859/content",
    "11208": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217880/content",
    "11209": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217866/content",
    "11210": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217897/content",
    "11211": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217903/content",
    "11212": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217934/content",
    "11213": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169218030/content",
    "11214": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169218047/content",
    "11216": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169218023/content",
    "11217": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169218016/content",
    "11218": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217972/content",
    "11219": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217989/content",
    "11220": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169218009/content",
    "11221": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217958/content",
    "11222": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217965/content",
    "11225": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217996/content",
    "11231": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169218115/content",
    "11242": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169218139/content",
    "11243": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169218108/content",
    "11630": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216302/content",
    "11638": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216388/content",
    "11639": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216395/content",
    "11640": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216401/content",
    "11651": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216517/content",
    "11652": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216524/content",
    "11653": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216531/content",
    "11654": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216548/content",
    "11655": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216555/content",
    "11656": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216562/content",
    "11657": "https://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169216579/content",
    "11700": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217002/content",
    "11701": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217019/content",
    "11702": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217026/content",
    "11703": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217033/content",
    "11704": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217040/content",
    "11705": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217057/content",
    "11706": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217064/content",
    "11707": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217071/content",
    "11708": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217088/content",
    "11709": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217095/content",
    "11710": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217101/content",
    "11711": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217118/content",
    "11712": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217125/content",
    "11713": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217132/content",
    "11714": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217149/content",
    "11715": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217156/content",
    "11716": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217163/content",
    "11717": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217170/content",
    "11718": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217187/content",
    "11719": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217194/content",
    "11720": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217200/content",
    "11721": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217217/content",
    "11722": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217224/content",
    "11723": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217231/content",
    "11724": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217248/content",
    "11725": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217255/content",
    "11726": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217262/content",
    "11727": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217279/content",
    "11728": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217286/content",
    "11729": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217293/content",
    "11730": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217309/content",
    "11731": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217316/content",
    "11732": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217323/content",
    "11733": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217330/content",
    "11734": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217347/content",
    "11735": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217354/content",
    "11736": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217361/content",
    "11737": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217378/content",
    "11738": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217385/content",
    "11739": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217392/content",
    "11740": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217408/content",
    "11741": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217415/content",
    "11742": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217422/content",
    "11743": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217439/content",
    "11744": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217446/content",
    "11745": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217453/content",
    "11746": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217460/content",
    "11747": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217477/content",
    "11748": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217484/content",
    "11749": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217491/content",
    "11750": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217507/content",
    "11751": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217514/content",
    "11752": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217521/content",
    "11753": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217538/content",
    "11754": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217545/content",
    "11755": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217552/content",
    "11756": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217569/content",
    "11757": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217576/content",
    "11758": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217583/content",
    "11759": "http://play.google.com/books/publish/u/0/a/535699340789858766#book/ISBN:9791169217590/content",
...  truncated for brevity ...
}

# --- Google Play Book Deletion Tool ---
class BookDeletionApp:
    # Constants
    BASE_URL = 'https://play.google.com/books/publish/u/0/?hl=ko'
    WAIT_TIME = 10
    PAGE_LOAD_WAIT = 5

    def __init__(self, master):
        self.master = master

        # Variables
        self.excel_file_path = None
        self.driver = None
        self.current_url = None
        self.total_processed = 0
        self.total_success = 0
        self.total_errors = 0

        self.create_widgets()

    def create_widgets(self):
        main_container = ctk.CTkFrame(self.master, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=20, pady=20)

        self.create_header_section(main_container)
        self.create_file_section(main_container)
        self.create_driver_section(main_container)
        self.create_progress_section(main_container)
        self.create_log_section(main_container)
        self.create_button_section(main_container)

    def create_header_section(self, parent):
        header_frame = ctk.CTkFrame(parent, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 20))
        title_label = ctk.CTkLabel(header_frame, text="📚 전자책 삭제 프로그램", font=ctk.CTkFont(size=28, weight="bold"), text_color=("#1f538d", "#1f538d"))
        title_label.pack(pady=(0, 5))
        subtitle_label = ctk.CTkLabel(header_frame, text="구글 플레이 도서에서 특정 사용자의 전자책을 자동으로 삭제합니다", font=ctk.CTkFont(size=14), text_color=("#666666", "#cccccc"))
        subtitle_label.pack()

    def create_file_section(self, parent):
        file_frame = ctk.CTkFrame(parent, corner_radius=12)
        file_frame.pack(fill="x", pady=(0, 15))

        section_title = ctk.CTkLabel(file_frame, text="📁 파일 선택", font=ctk.CTkFont(size=18, weight="bold"), text_color=("#1f538d", "#1f538d"))
        section_title.pack(pady=(15, 10), padx=15, anchor="w")

        button_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        button_frame.pack(fill="x", padx=15, pady=(10, 10))

        select_btn = ctk.CTkButton(button_frame, text="📄 엑셀 파일 선택", command=self.select_excel, font=ctk.CTkFont(size=14, weight="bold"), height=40, corner_radius=8)
        select_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))

        template_btn = ctk.CTkButton(button_frame, text="📋 템플릿 생성", command=self.create_template, font=ctk.CTkFont(size=14, weight="bold"), height=40, corner_radius=8, fg_color=("#28a745", "#28a745"), hover_color=("#218838", "#218838"))
        template_btn.pack(side="right", fill="x", expand=True, padx=(5, 0))

        self.excel_label = ctk.CTkLabel(file_frame, text="선택된 엑셀: 없음", font=ctk.CTkFont(size=12), text_color=("#666666", "#cccccc"))
        self.excel_label.pack(pady=(0, 15), padx=15, anchor="w")

    def create_driver_section(self, parent):
        driver_frame = ctk.CTkFrame(parent, corner_radius=12)
        driver_frame.pack(fill="x", pady=(0, 15))
        section_title = ctk.CTkLabel(driver_frame, text="🚀 크롬드라이버 설정", font=ctk.CTkFont(size=18, weight="bold"), text_color=("#1f538d", "#1f538d"))
        section_title.pack(pady=(15, 10), padx=15, anchor="w")
        info_frame = ctk.CTkFrame(driver_frame, fg_color="transparent")
        info_frame.pack(fill="x", padx=15, pady=(0, 15))
        status_frame = ctk.CTkFrame(info_frame, fg_color="transparent")
        status_frame.pack(fill="x")
        status_icon = ctk.CTkLabel(status_frame, text="✅", font=ctk.CTkFont(size=16))
        status_icon.pack(side="left", padx=(0, 8))
        status_text = ctk.CTkLabel(status_frame, text="자동으로 관리됩니다", font=ctk.CTkFont(size=14, weight="bold"), text_color=("#28a745", "#28a745"))
        status_text.pack(side="left")
        desc_text = ctk.CTkLabel(info_frame, text="webdriver-manager가 자동으로 최신 버전을 다운로드하고 관리합니다", font=ctk.CTkFont(size=12), text_color=("#666666", "#cccccc"))
        desc_text.pack(anchor="w")

    def create_progress_section(self, parent):
        progress_frame = ctk.CTkFrame(parent, corner_radius=12)
        progress_frame.pack(fill="x", pady=(0, 15))
        section_title = ctk.CTkLabel(progress_frame, text="📊 진행률", font=ctk.CTkFont(size=18, weight="bold"), text_color=("#1f538d", "#1f538d"))
        section_title.pack(pady=(15, 10), padx=15, anchor="w")
        self.progress_var = ctk.DoubleVar()
        self.progress_bar = ctk.CTkProgressBar(progress_frame, variable=self.progress_var, height=8, corner_radius=4)
        self.progress_bar.pack(fill="x", padx=15, pady=(0, 15))
        self.progress_bar.set(0)

    def create_log_section(self, parent):
        log_frame = ctk.CTkFrame(parent, corner_radius=12)
        log_frame.pack(fill="both", expand=True, pady=(0, 15))
        section_title = ctk.CTkLabel(log_frame, text="📝 작업 로그", font=ctk.CTkFont(size=18, weight="bold"), text_color=("#1f538d", "#1f538d"))
        section_title.pack(pady=(15, 10), padx=15, anchor="w")
        self.log_text = ctk.CTkTextbox(log_frame, font=ctk.CTkFont(size=12), corner_radius=8)
        self.log_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))

    def create_button_section(self, parent):
        button_frame = ctk.CTkFrame(parent, fg_color="transparent")
        button_frame.pack(fill="x")
        self.start_button = ctk.CTkButton(button_frame, text="🚀 작업 시작", command=self.start_deletion_thread, font=ctk.CTkFont(size=16, weight="bold"), height=50, corner_radius=10, fg_color=("#007bff", "#007bff"), hover_color=("#0056b3", "#0056b3"))
        self.start_button.pack(side="left", fill="x", expand=True, padx=(0, 10))
        # quit_button 제거 - 메인 창에서 도구를 전환할 수 있으므로 불필요

    def log_message(self, message, level="INFO"):
        timestamp = time.strftime("%H:%M:%S")
        emoji_map = {"INFO": "ℹ️", "SUCCESS": "✅", "WARNING": "⚠️", "ERROR": "❌", "SUMMARY": "📊"}
        emoji = emoji_map.get(level, "ℹ️")
        log_entry = f"[{timestamp}] {emoji} {level}: {message}\n"
        self.log_text.insert("end", log_entry)
        self.log_text.see("end")
        self.master.update_idletasks()

    def handle_error(self, error, context=""):
        error_message = f"{context}: {str(error)}" if context else str(error)
        self.log_message(error_message, "ERROR")
        self.total_errors += 1

    def select_excel(self):
        file_path = filedialog.askopenfilename(title="엑셀 파일 선택", filetypes=[("Excel 파일", "*.xlsx *.xls")])
        if file_path:
            try:
                df = pd.read_excel(file_path)
                if len(df.columns) < 2:
                    messagebox.showerror("오류", "엑셀 파일에 최소 2개의 열(A, B)이 필요합니다.")
                    return
                self.excel_file_path = file_path
                filename = os.path.basename(file_path)
                self.excel_label.configure(text=f"선택된 엑셀: {filename}")
                self.log_message(f"엑셀 파일이 선택되었습니다: {filename}", "SUCCESS")
                self.log_message(f"총 {len(df)}개의 항목이 발견되었습니다", "INFO")
            except Exception as e:
                messagebox.showerror("오류", f"엑셀 파일을 읽는 중 오류가 발생했습니다:\n{str(e)}")

    def create_template(self):
        try:
            file_path = filedialog.asksaveasfilename(title="템플릿 파일 저장", defaultextension=".xlsx", filetypes=[("Excel 파일", "*.xlsx")])
            if file_path:
                template_data = {'구글플레이도서_URL': ['https://play.google.com/books/publish/u/0/book/123456789'], '삭제할_이메일': ['user1@example.com']}
                df = pd.DataFrame(template_data)
                df.to_excel(file_path, index=False)
                self.log_message(f"템플릿 파일이 생성되었습니다: {os.path.basename(file_path)}", "SUCCESS")
                messagebox.showinfo("완료", f"템플릿 파일이 생성되었습니다:\n{file_path}")
        except Exception as e:
            messagebox.showerror("오류", f"템플릿 생성 중 오류가 발생했습니다:\n{str(e)}")

    def check_email_exists(self, email):
        try:
            email_xpath = f"//h4[@class='mat-mdc-list-item-title mdc-list-item__primary-text' and contains(text(),'{email}')]"
            elements = self.driver.find_elements(By.XPATH, email_xpath)
            return len(elements) > 0
        except Exception as e:
            self.log_message(f"이메일 존재 확인 중 오류: {str(e)}", "ERROR")
            return False

    def click_delete_button(self, email):
        try:
            email_xpath = f"//h4[@class='mat-mdc-list-item-title mdc-list-item__primary-text' and contains(text(),'{email}')]"
            email_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, email_xpath)))
            item_element = email_element.find_element(By.XPATH, "./ancestor::mat-list-item")
            delete_button = item_element.find_element(By.XPATH, ".//button[@aria-label='삭제']")
            self.driver.execute_script("arguments[0].scrollIntoView(true);", delete_button)
            time.sleep(0.5)
            self.driver.execute_script("arguments[0].click();", delete_button)
            return True
        except Exception as e:
            self.log_message(f"삭제 버튼 클릭 실패: {str(e)}", "ERROR")
            return False

    def click_confirm_button(self):
        try:
            confirm_selectors = ["//button[contains(text(), '삭제')]", "//button[contains(text(), '확인')]"]
            confirm_button = None
            for selector in confirm_selectors:
                try:
                    confirm_button = WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable((By.XPATH, selector)))
                    break
                except:
                    continue
            if confirm_button:
                self.driver.execute_script("arguments[0].click();", confirm_button)
                return True
            else:
                self.log_message(f"삭제 확인 버튼을 찾을 수 없습니다", "WARNING")
                return False
        except Exception as e:
            self.log_message(f"삭제 확인 버튼 클릭 실패: {str(e)}", "ERROR")
            return False

    def verify_deletion(self, email):
        try:
            self.driver.refresh()
            time.sleep(3)
            email_xpath = f"//h4[@class='mat-mdc-list-item-title mdc-list-item__primary-text' and contains(text(),'{email}')]"
            remaining_elements = self.driver.find_elements(By.XPATH, email_xpath)
            return len(remaining_elements) == 0
        except Exception as e:
            self.log_message(f"삭제 완료 확인 중 오류: {str(e)}", "WARNING")
            return False

    def start_deletion_thread(self):
        if not self.excel_file_path:
            messagebox.showerror("오류", "엑셀 파일을 선택해주세요.")
            return
        self.start_button.configure(state="disabled", text="🔄 처리 중...")
        threading.Thread(target=self.start_deletion_logic, daemon=True).start()

    def start_deletion_logic(self):
        try:
            start_time = time.time()
            self.log_message("작업을 시작합니다...", "INFO")
            self.log_message("크롬드라이버를 자동으로 설정합니다...", "INFO")

            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service)

            df = pd.read_excel(self.excel_file_path)
            book_urls = df.iloc[:, 0].dropna().tolist()
            emails = df.iloc[:, 1].dropna().tolist()

            self.driver.get(self.BASE_URL)
            self.log_message("로그인이 필요합니다. 로그인을 완료한 후 확인 버튼을 눌러주세요.", "WARNING")
            messagebox.showinfo("안내", "로그인을 완료한 후 확인을 눌러주세요.")

            url_groups = {}
            for url, email in zip(book_urls, emails):
                if url not in url_groups:
                    url_groups[url] = []
                url_groups[url].append(email)

            actual_total_items = len(book_urls)
            processed_count = 0

            for url, email_list in url_groups.items():
                if url != self.current_url:
                    self.log_message(f"페이지 로딩: {url}", "INFO")
                    self.driver.get(url)
                    WebDriverWait(self.driver, self.WAIT_TIME).until(lambda d: d.execute_script("return document.readyState") == "complete")
                    time.sleep(self.PAGE_LOAD_WAIT)
                    self.current_url = url

                for email in email_list:
                    processed_count += 1
                    progress = (processed_count / actual_total_items)
                    self.progress_var.set(progress)

                    self.log_message(f"[{processed_count}/{actual_total_items}] {email} 처리 중...", "INFO")

                    if self.check_email_exists(email):
                        if self.click_delete_button(email):
                            time.sleep(1)
                            if self.click_confirm_button():
                                time.sleep(3)
                                if self.verify_deletion(email):
                                    self.log_message(f"{email} 삭제 성공", "SUCCESS")
                                    self.total_success += 1
                                else:
                                    self.log_message(f"{email} 삭제 실패 (삭제 후에도 확인됨)", "ERROR")
                                    self.total_errors += 1
                            else:
                                self.log_message(f"{email} 삭제 확인 실패", "ERROR")
                                self.total_errors += 1
                        else:
                            self.log_message(f"{email} 삭제 버튼 클릭 실패", "ERROR")
                            self.total_errors += 1
                    else:
                        self.log_message(f"{email} 해당 이메일을 찾을 수 없습니다", "WARNING")
                    self.total_processed += 1
                    time.sleep(1)

            end_time = time.time()
            self.log_message(f"총 처리 시간: {end_time - start_time:.1f}초", "INFO")
            self.show_summary()

        except Exception as e:
            self.handle_error(e, "전체 작업 실패")
        finally:
            if self.driver:
                self.driver.quit()
            self.start_button.configure(state="normal", text="🚀 작업 시작")

    def show_summary(self):
        summary = f"📊 작업 완료!\n\n총 처리: {self.total_processed}개\n성공: {self.total_success}개\n실패: {self.total_errors}개"
        self.log_message(summary, "SUMMARY")
        messagebox.showinfo("작업 완료", summary)


# --- Google Play Ebook Registration Tool ---

class GoogleEbookRegistrator:
    """Handles the Selenium automation logic for ebook registration."""
    def __init__(self, log_widget):
        self.driver = None
        self.wait = None
        self.log_widget = log_widget
        self.setup_driver()

    def log(self, message, level="INFO"):
        """Logs a message to the GUI widget with a level."""
        timestamp = time.strftime("%H:%M:%S")
        emoji_map = {"INFO": "ℹ️", "SUCCESS": "✅", "WARNING": "⚠️", "ERROR": "❌"}
        emoji = emoji_map.get(level, "ℹ️")
        log_entry = f"[{timestamp}] {emoji} {level}: {message}\n"
        self.log_widget.insert("end", log_entry)
        self.log_widget.see("end")

    def setup_driver(self):
        """Sets up the Chrome driver using webdriver-manager."""
        self.log("크롬 드라이버를 설정합니다...")
        try:
            chrome_options = Options()
            # chrome_options.add_argument("--headless") # Headless mode can be enabled here if needed
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.wait = WebDriverWait(self.driver, 20)
            self.log("크롬 드라이버 설정 완료.")
        except Exception as e:
            self.log(f"드라이버 설정 오류: {e}")
            messagebox.showerror("드라이버 오류", f"크롬 드라이버를 설정하는 중 오류가 발생했습니다: {e}")
            raise

    def login(self):
        """Navigates to login page and waits for user to log in."""
        self.log("로그인 페이지로 이동합니다. 로그인 후 확인을 눌러주세요.")
        self.driver.get("https://play.google.com/books/publish/u/0/?hl=ko")
        messagebox.showinfo("로그인 필요", "로그인 페이지가 열렸습니다.\n로그인을 완료한 후 확인을 눌러주세요.")

        self.wait.until(lambda driver: "books/publish" in driver.current_url)
        self.log("로그인 성공을 확인했습니다.")

    def process_google_play_registration(self, email: str) -> bool:
        """Adds a content reviewer to the currently loaded book page."""
        try:
            self.log(f"'{email}'에 대한 검토자 추가를 시작합니다.")

            # Check if email already exists to avoid errors
            try:
                # This selector is a guess, might need adjustment
                email_xpath = f"//div[contains(@class, 'reviewer-email') and contains(text(), '{email}')]"
                existing_elements = self.driver.find_elements(By.XPATH, email_xpath)
                if len(existing_elements) > 0:
                    self.log(f"'{email}'은(는) 이미 검토자로 등록되어 있습니다.", "WARNING")
                    return True # Success, as the goal is already met
            except NoSuchElementException:
                # This is expected if the user is not yet a reviewer
                pass

            # Click "Add content reviewer" button
            self.log("'콘텐츠 검토자 추가' 버튼을 찾습니다...")
            add_reviewer_button_xpath = "//button[.//span[contains(text(), '콘텐츠 검토자 추가')]]"
            add_reviewer_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, add_reviewer_button_xpath)))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", add_reviewer_button)
            time.sleep(0.5)
            add_reviewer_button.click()
            self.log("'콘텐츠 검토자 추가' 버튼 클릭 완료.")
            time.sleep(2)

            # Enter email address in the dialog's textarea
            self.log(f"'{email}'을 입력할 텍스트 영역을 찾습니다...")
            email_textarea_xpath = "//textarea[@aria-label='이메일 주소']"
            email_textarea = self.wait.until(EC.presence_of_element_located((By.XPATH, email_textarea_xpath)))
            email_textarea.send_keys(email)
            self.log("이메일 주소 입력 완료.")
            time.sleep(1)

            # Click the final "Add" button in the dialog
            self.log("'추가' 버튼을 찾습니다...")
            add_button_xpath = "//mat-dialog-actions//button[contains(., '추가')]"
            add_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, add_button_xpath)))
            add_button.click()
            self.log("'추가' 버튼 클릭 완료.")
            time.sleep(4) # Wait for action to complete

            self.log(f"'{email}' 검토자 추가 성공.", "SUCCESS")
            return True

        except (TimeoutException, NoSuchElementException):
            self.log(f"오류: UI 요소를 찾을 수 없어 '{email}'을(를) 추가할 수 없습니다. 페이지가 변경되었거나 로딩이 너무 오래 걸릴 수 있습니다.", "ERROR")
            try:
                cancel_button_xpath = "//mat-dialog-actions//button[contains(., '취소')]"
                cancel_button = self.driver.find_element(By.XPATH, cancel_button_xpath)
                cancel_button.click()
                self.log("열려있는 대화 상자를 닫았습니다.")
            except:
                pass
            return False
        except Exception as e:
            self.log(f"'{email}' 처리 중 예상치 못한 오류 발생: {e}", "ERROR")
            return False

    def close(self):
        if self.driver:
            self.driver.quit()

class GoogleEbookRegistrationApp:
    def __init__(self, master):
        self.master = master
        self.excel_file_path = ""
        self.create_widgets()

    def create_widgets(self):
        main_container = ctk.CTkFrame(self.master, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=20, pady=20)

        self.create_header_section(main_container)
        self.create_file_section(main_container)
        self.create_progress_section(main_container)
        self.create_log_section(main_container)
        self.create_button_section(main_container)

    def create_header_section(self, parent):
        header_frame = ctk.CTkFrame(parent, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 20))
        title_label = ctk.CTkLabel(header_frame, text="🚀 Google Play 전자책 등록", font=ctk.CTkFont(size=28, weight="bold"), text_color=("#1f538d", "#1f538d"))
        title_label.pack(pady=(0, 5))
        subtitle_label = ctk.CTkLabel(header_frame, text="Google Play 도서에 콘텐츠 검토자를 자동으로 등록합니다", font=ctk.CTkFont(size=14), text_color=("#666666", "#cccccc"))
        subtitle_label.pack()

    def create_file_section(self, parent):
        file_frame = ctk.CTkFrame(parent, corner_radius=12)
        file_frame.pack(fill="x", pady=(0, 15))

        section_title = ctk.CTkLabel(file_frame, text="📁 파일 선택", font=ctk.CTkFont(size=18, weight="bold"), text_color=("#1f538d", "#1f538d"))
        section_title.pack(pady=(15, 10), padx=15, anchor="w")

        button_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        button_frame.pack(fill="x", padx=15, pady=(10, 10))

        select_btn = ctk.CTkButton(button_frame, text="📄 엑셀 파일 선택", command=self.select_excel_file, font=ctk.CTkFont(size=14, weight="bold"), height=40, corner_radius=8)
        select_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))

        template_btn = ctk.CTkButton(button_frame, text="📋 템플릿 생성", command=self.create_template, font=ctk.CTkFont(size=14, weight="bold"), height=40, corner_radius=8, fg_color=("#28a745", "#28a745"), hover_color=("#218838", "#218838"))
        template_btn.pack(side="right", fill="x", expand=True, padx=(5, 0))

        self.excel_label = ctk.CTkLabel(file_frame, text="선택된 엑셀: 없음", font=ctk.CTkFont(size=12), text_color=("#666666", "#cccccc"))
        self.excel_label.pack(pady=(0, 15), padx=15, anchor="w")

    def create_progress_section(self, parent):
        progress_frame = ctk.CTkFrame(parent, corner_radius=12)
        progress_frame.pack(fill="x", pady=(0, 15))
        section_title = ctk.CTkLabel(progress_frame, text="📊 진행률", font=ctk.CTkFont(size=18, weight="bold"), text_color=("#1f538d", "#1f538d"))
        section_title.pack(pady=(15, 10), padx=15, anchor="w")
        self.progress_var = ctk.DoubleVar()
        self.progress_bar = ctk.CTkProgressBar(progress_frame, variable=self.progress_var, height=8, corner_radius=4)
        self.progress_bar.pack(fill="x", padx=15, pady=(0, 15))
        self.progress_bar.set(0)

    def create_log_section(self, parent):
        log_frame = ctk.CTkFrame(parent, corner_radius=12)
        log_frame.pack(fill="both", expand=True, pady=(0, 15))
        section_title = ctk.CTkLabel(log_frame, text="📝 작업 로그", font=ctk.CTkFont(size=18, weight="bold"), text_color=("#1f538d", "#1f538d"))
        section_title.pack(pady=(15, 10), padx=15, anchor="w")
        self.log_text = ctk.CTkTextbox(log_frame, font=ctk.CTkFont(size=12), corner_radius=8)
        self.log_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))

    def create_button_section(self, parent):
        button_frame = ctk.CTkFrame(parent, fg_color="transparent")
        button_frame.pack(fill="x")
        self.run_button = ctk.CTkButton(button_frame, text="🚀 작업 시작", command=self.run_script, font=ctk.CTkFont(size=16, weight="bold"), height=50, corner_radius=10, fg_color=("#007bff", "#007bff"), hover_color=("#0056b3", "#0056b3"))
        self.run_button.pack(fill="x")

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(title="엑셀 파일 선택", filetypes=[("Excel 파일", "*.xlsx *.xls")])
        if file_path:
            try:
                df = pd.read_excel(file_path, dtype=str)
                # Expecting 2 columns: Book Code and Gmail
                if len(df.columns) < 2:
                    messagebox.showerror("오류", "엑셀 파일에 최소 2개의 열(도서 코드, 지메일)이 필요합니다.")
                    return
                self.excel_file_path = file_path
                filename = os.path.basename(file_path)
                self.excel_label.configure(text=f"선택된 엑셀: {filename}")
                self.log_text.insert("end", f"엑셀 파일이 선택되었습니다: {filename}\n")
                self.log_text.insert("end", f"총 {len(df)}개의 항목이 발견되었습니다\n")
            except Exception as e:
                messagebox.showerror("오류", f"엑셀 파일을 읽는 중 오류가 발생했습니다:\n{str(e)}")

    def create_template(self):
        try:
            file_path = filedialog.asksaveasfilename(title="템플릿 파일 저장", defaultextension=".xlsx", filetypes=[("Excel 파일", "*.xlsx")])
            if file_path:
                template_data = {'도서 코드': ['4599'], '지메일': ['user@example.com']}
                df = pd.DataFrame(template_data)
                df.to_excel(file_path, index=False)
                self.log_text.insert("end", f"템플릿 파일이 생성되었습니다: {os.path.basename(file_path)}\n")
                messagebox.showinfo("완료", f"템플릿 파일이 생성되었습니다:\n{file_path}")
        except Exception as e:
            messagebox.showerror("오류", f"템플릿 생성 중 오류가 발생했습니다:\n{str(e)}")

    def run_script(self):
        if not self.excel_file_path:
            messagebox.showerror("오류", "엑셀 파일을 먼저 선택해주세요.")
            return

        self.run_button.configure(state="disabled", text="🔄 처리 중...")
        self.log_text.delete("1.0", "end")
        self.progress_bar.set(0)

        # Run the automation in a separate thread to keep the GUI responsive
        threading.Thread(target=self.run_script_threaded, daemon=True).start()

    def run_script_threaded(self):
        ebook_reg = None
        try:
            df = pd.read_excel(self.excel_file_path, dtype=str)
            # Make sure to read the first two columns, regardless of their names
            book_codes = df.iloc[:, 0].dropna().tolist()
            emails = df.iloc[:, 1].dropna().tolist()

            if not book_codes or not emails:
                messagebox.showerror("오류", "엑셀 파일에 필요한 데이터(도서 코드, 지메일)가 부족합니다.")
                self.run_button.configure(state="normal", text="🚀 작업 시작")
                return

            ebook_reg = GoogleEbookRegistrator(log_widget=self.log_text)
            ebook_reg.login()

            success_count = 0
            error_count = 0
            total_count = len(book_codes)
            ebook_reg.log(f"총 {total_count}건의 등록을 시작합니다.")

            current_url = None

            for i, (code, email) in enumerate(zip(book_codes, emails)):
                progress = (i + 1) / total_count
                self.progress_bar.set(progress)
                ebook_reg.log(f"--- [{i+1}/{total_count}] {email} 처리 중 ---")

                book_url = BOOK_URL_MAP.get(str(code))
                if not book_url:
                    ebook_reg.log(f"도서 코드 '{code}'에 해당하는 URL을 찾을 수 없습니다.", "ERROR")
                    error_count += 1
                    continue

                try:
                    # To optimize, only switch pages if the URL is different
                    if book_url != current_url:
                        ebook_reg.log(f"페이지로 이동 중: {book_url[:50]}...")
                        ebook_reg.driver.get(book_url)
                        # Wait for a known element on the book page to ensure it's loaded
                        WebDriverWait(ebook_reg.driver, 15).until(
                            EC.presence_of_element_located((By.XPATH, "//div[contains(text(), '콘텐츠')]"))
                        )
                        current_url = book_url
                        time.sleep(2) # Extra wait for safety

                    if ebook_reg.process_google_play_registration(email):
                        success_count += 1
                    else:
                        error_count += 1
                except Exception as e:
                    ebook_reg.log(f"'{email}' 처리 중 심각한 오류 발생: {e}", "ERROR")
                    error_count += 1

                time.sleep(1)

            self.progress_bar.set(1.0)
            summary = f"총 {total_count}건 중 {success_count}건 성공, {error_count}건 실패."
            ebook_reg.log(f"--- 작업 완료 ---", "SUMMARY")
            ebook_reg.log(summary, "SUMMARY")
            messagebox.showinfo("완료", summary)

        except Exception as e:
            messagebox.showerror("오류", f"전체 작업 중 오류 발생: {e}")
            if ebook_reg:
                ebook_reg.log(f"치명적인 오류로 작업을 중단합니다: {e}", "ERROR")
        finally:
            if ebook_reg:
                ebook_reg.close()
            self.run_button.configure(state="normal", text="🚀 작업 시작")


# --- Main Menu ---
class MainMenu:
    def __init__(self, root):
        self.root = root
        self.root.title("Ebook Tool Suite")
        self.root.geometry("1200x800")

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        # 현재 표시할 도구
        self.current_tool = None
        self.current_tool_instance = None

        self.create_widgets()

    def create_widgets(self):
        # 메인 컨테이너
        main_container = ctk.CTkFrame(self.root, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=20, pady=20)

        # 헤더 섹션
        self.create_header_section(main_container)
        
        # 도구 선택 버튼 섹션
        self.create_tool_selection_section(main_container)
        
        # 도구 컨테이너 (동적으로 변경됨)
        self.tool_container = ctk.CTkFrame(main_container, fg_color="transparent")
        self.tool_container.pack(fill="both", expand=True, pady=(20, 0))

    def create_header_section(self, parent):
        header_frame = ctk.CTkFrame(parent, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 20))
        
        title_label = ctk.CTkLabel(header_frame, text="Ebook Tool Suite", font=ctk.CTkFont(size=28, weight="bold"), text_color=("#1f538d", "#1f538d"))
        title_label.pack(pady=(0, 5))
        
        subtitle_label = ctk.CTkLabel(header_frame, text="전자책 관련 자동화 도구 모음", font=ctk.CTkFont(size=14), text_color=("#666666", "#cccccc"))
        subtitle_label.pack()

    def create_tool_selection_section(self, parent):
        button_frame = ctk.CTkFrame(parent, corner_radius=12)
        button_frame.pack(fill="x", pady=(0, 20))
        
        section_title = ctk.CTkLabel(button_frame, text="🛠️ 도구 선택", font=ctk.CTkFont(size=18, weight="bold"), text_color=("#1f538d", "#1f538d"))
        section_title.pack(pady=(15, 15), padx=15, anchor="w")
        
        button_container = ctk.CTkFrame(button_frame, fg_color="transparent")
        button_container.pack(fill="x", padx=15, pady=(0, 15))
        
        google_button = ctk.CTkButton(
            button_container, 
            text="📚 Google Play 도서 삭제", 
            command=self.open_google_tool, 
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=8
        )
        google_button.pack(side="left", fill="x", expand=True, padx=(0, 10))

        google_reg_button = ctk.CTkButton(
            button_container, 
            text="🚀 Google Play eBook 등록",
            command=self.open_google_registration_tool,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=8
        )
        google_reg_button.pack(side="right", fill="x", expand=True, padx=(10, 0))

    def clear_tool_container(self):
        """도구 컨테이너를 비웁니다."""
        for widget in self.tool_container.winfo_children():
            widget.destroy()
        if self.current_tool_instance:
            self.current_tool_instance = None

    def open_google_tool(self):
        """Google Play 도서 삭제 도구를 메인 창에 표시합니다."""
        self.clear_tool_container()
        self.current_tool = "google_delete"
        self.current_tool_instance = BookDeletionApp(self.tool_container)
        self.root.title("Ebook Tool Suite - Google Play 도서 삭제")

    def open_google_registration_tool(self):
        """Google Play eBook 등록 도구를 메인 창에 표시합니다."""
        self.clear_tool_container()
        self.current_tool = "google_register"
        self.current_tool_instance = GoogleEbookRegistrationApp(self.tool_container)
        self.root.title("Ebook Tool Suite - Google Play eBook 등록")

if __name__ == "__main__":
    root = ctk.CTk()
    app = MainMenu(root)
    root.mainloop()
