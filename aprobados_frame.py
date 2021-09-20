from difflib import SequenceMatcher
import re
from PyPDF2 import pdf
import pandas as pd

import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

from colorama import init, Fore, Back, Style
# Initializes Colorama
init(autoreset=True)

import PyPDF2
def match_string_percentage(str1: str, str2: str):
    return SequenceMatcher(None, str1, str2).ratio()
    
file = open(input("Nombre del archivo de los aprobados: ") or "aprobados.pdf", "rb")
pdfReader = PyPDF2.PdfFileReader(file)
n_pages = pdfReader.numPages

people_dict = {"nombre": [], "especialidad": [], "nota": []}

for k in range(n_pages):
    page = pdfReader.getPage(k).extractText()
    page = page.replace("\n", "")
    print()
    espec_search = re.search(r"Especialidad:\d+\D+\d+(.*)Apellidos", page)

    if espec_search != None:
        especialidad = espec_search.group(1)

    lines = re.compile(r"\D+\d{8}[a-zA-Z]\d[,]\d{4}").findall(page)

    for line in lines:
        dni = re.search(r"\d{8}[a-zA-Z]", line).group(0)
        name = re.search(r"\D+", line).group(0)
        mark = re.search(r"\d[,]\d{4}", line).group(0)

        print(dni, name, mark)

        people_dict["nombre"].append(dni + " - " + name)
        people_dict["especialidad"].append(especialidad)
        people_dict["nota"].append(mark)
df = pd.DataFrame(people_dict)
df.to_excel("aprobados.xlsx")
