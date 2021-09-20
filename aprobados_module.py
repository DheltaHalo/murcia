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
    final_names = list(final["nombre"])
    aprobados_names = list(aprobados["nombre"])
    aprobados_especialidad = list(aprobados["especialidad"])
    aprobados_nota = list(aprobados["nota"])

    print(Fore.GREEN + "Iniciando bucle.")
    len_left = len(aprobados_names)
    for k, final_name in enumerate(final_names):
        for i, aprobado_name in enumerate(aprobados_names):
            if match_string_percentage(final_name, aprobado_name) > 0.75:
                print(aprobado_name)
                if match_string_percentage(final["especialidad"][k], aprobados_especialidad[i]) > 0.75:
                    print(Fore.CYAN + f"{aprobado_name} ha sido " + Fore.YELLOW + "ENCONTRADO." + aprobados_nota[i])
                    final["nota_consiguen_plaza"][k] = aprobados_nota[i]

                    aprobados_names.pop(i)
                    aprobados_especialidad.pop(i)
                    aprobados_nota.pop(i)
                    break
        if len(aprobados_names) != len_left:
            print(len(aprobados_names))
            len_left = len(aprobados_names)
    print(aprobados_names)
    final.to_excel("final_2.xlsx")

if __name__ == "__main__":
    main()