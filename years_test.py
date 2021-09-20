from difflib import SequenceMatcher
import re
import pandas as pd
from pandas.core import frame
import tabula
from math import nan
import numpy as np

import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

from colorama import init, Fore, Back, Style
# Initializes Colorama
init(autoreset=True)

def match_string_percentage(str1: str, str2: str):
    return SequenceMatcher(None, str1, str2).ratio()

def create_years_frame():
    # We create the list from the pdf PAGES
    pdf_name = "baremo.pdf"#input("Nombre del pdf: ")
    print(Fore.CYAN + "Creating list from " + Fore.YELLOW + pdf_name)

    tb = tabula.read_pdf(pdf_name, pages = "all")

    people_dict = {"nombre": [], "b1": [], "b2": [], "b3": [], "b4": [], "total": [], "opos_sup": []}
    people_keys = list(people_dict.keys())

    for p_index, page in enumerate(tb):
        print(Fore.CYAN + "Página " + Fore.YELLOW + str(p_index + 1))
        page = page.fillna("")
        frame_keys = page.keys()

        name_list = []
        marks_list = []
        opos_list = []

        dni_regex = r"[*]+\d+[*]+|[.]+\d+[a-zA-Z]|\d+[a-zA-Z]"
        dni_list = re.compile(dni_regex).findall(str(page[frame_keys[0]]))
        holder_int = 0

        for k in dni_list:
            index = page[page[frame_keys[0]]==k].index.values.astype(int)
           
            if len(index) > 1:
                index = np.array([index[holder_int]], np.int32)
            name = page.iloc[index][frame_keys[1]].item()
            marks = page.iloc[index][frame_keys[3]].item()
            opos = page.iloc[index][frame_keys[4]].item()
            
            if name == "":
                name = page.iloc[index - 1][frame_keys[1]].item() + " " + page.iloc[index + 1][frame_keys[1]].item()

            if len(marks.split(" ")) == 3:
                marks = page.iloc[index - 1][frame_keys[3]].item() + " " + page.iloc[index][frame_keys[3]].item()

            name_list.append(name)
            marks_list.append(marks)
            opos_list.append(opos)

            if len(index) > 1:
                holder_int += 1

        for i in range(len(dni_list)):
            people_dict[people_keys[0]].append(dni_list[i] + " - " + name_list[i])

            marks_split = re.compile(r"\d[,]\d{4}").findall(marks_list[i])
            marks_split = [float(x.replace(",", ".")) for x in marks_split]

            if len(marks_split) == 4:
                total = marks_split[0]/0.6 + marks_split[1]/0.3 + marks_split[2]/0.3 + marks_split[3]/0.15
            else:
                total = f"ERROR en página {p_index} con dni {dni_list[i]}."
                for k in range(4 - len(marks_split)):
                    marks_split.append("Error")

            people_dict[people_keys[1]].append(marks_split[0])
            people_dict[people_keys[2]].append(marks_split[1])
            people_dict[people_keys[3]].append(marks_split[2])
            people_dict[people_keys[4]].append(marks_split[3])
            people_dict[people_keys[5]].append(total)
            people_dict[people_keys[6]].append(float(opos_list[i].replace(",", ".")))

    return people_dict        

def main():
    frame = create_years_frame()
    df = pd.DataFrame(frame)
    df.to_excel("years.xlsx", index=False)

if __name__ == "__main__":
    main()