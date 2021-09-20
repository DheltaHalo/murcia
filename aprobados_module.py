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
def create_aprobados_frame():
    file = open("aprobados.pdf", "rb")
    pdfReader = PyPDF2.PdfFileReader(file)
    n_pages = pdfReader.numPages

    people_dict = {"nombre": [], "especialidad": [], "nota": []}

    for k in range(n_pages):
        page = pdfReader.getPage(k).extractText()
        if k == 23:
            print(page)
        page = page.replace("\n", "")

        espec_search = re.search(r"Especialidad:\d+\D+\d+(.*)Acceso:", page)

        if espec_search != None:
            especialidad = espec_search.group(1)

        lines = re.compile(r"[*]+\d+[*]+\D+\d[,]\d{4}").findall(page)
        if k == 23:
            print(lines)
        for line in lines:
            name = re.search(r"[*]+\d+[*]\D+", line).group(0)
            mark = re.search(r"\d[,]\d{4}", line).group(0)

            people_dict["nombre"].append(name)
            people_dict["especialidad"].append(especialidad)
            people_dict["nota"].append(mark)
    df = pd.DataFrame(people_dict)
    df.to_excel("aprobados.xlsx")
    return df

def match_string_percentage(str1: str, str2: str):
    return SequenceMatcher(None, str1, str2).ratio()

def main():
    final_pdf = "final.xlsx"
    aprobados = "aprobados.xlsx"

    print(Fore.CYAN + "Leyendo el archivo " + Fore.YELLOW + final_pdf)
    final = pd.read_excel(final_pdf)
    print(Fore.CYAN + "Leyendo el archivo " + Fore.YELLOW + "aprobados")
    # aprobados = create_aprobados_frame()
    aprobados = pd.read_excel(aprobados)

    final["nota_consiguen_plaza"] = 0

    print(Fore.CYAN + "Creando listas.")
    final_names = list(final["Nombre y apellidos"])
    aprobados_names = list(aprobados["nombre"])
    aprobados_especialidad = list(aprobados["especialidad"])

    print(Fore.GREEN + "Iniciando bucle.")
    len_left = len(aprobados_names)
    for k, final_name in enumerate(final_names):
        for i, aprobado_name in enumerate(aprobados_names):
            if match_string_percentage(final_name, aprobado_name) > 0.75:
                print(aprobado_name)
                if match_string_percentage(final["Especialidad"][k], aprobados_especialidad[i]) > 0.75:
                    print(Fore.CYAN + f"{aprobado_name} ha sido " + Fore.YELLOW + "ENCONTRADO.")
                    final["nota_consiguen_plaza"][k] = aprobados["nota"][i]

                    aprobados_names.pop(i)
                    aprobados_especialidad.pop(i)
                    break
        if len(aprobados_names) != len_left:
            print(len(aprobados_names))
            len_left = len(aprobados_names)
    print(aprobados_names)
    final.to_excel("final_2.xlsx")

if __name__ == "__main__":
    main()