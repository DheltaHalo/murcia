import enum
import PyPDF2
import re
import numpy as np
import pandas as pd
import tabula

import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

import murcia_create_person as murc_functions

from colorama import init, Fore, Back, Style
# Initializes Colorama
init(autoreset=True)

class func(murc_functions.func):
    @staticmethod
    def match_percent(string1: str, string2: str):
        return float(len(set(string1) & set(string2 )) / len(set(string1) | set(string2)))
        
class Line_Builder(murc_functions.Line_builder):
    def build_header(self):
        self.especialidad = re.search(r"Especialidad(.*)Experiencia", self.header).group(1)
        self.especialidad = re.search(r"[a-zA-Z\s]+", self.especialidad).group(0)[1:]

        return self.especialidad

def create_years_frame():
    print(Fore.CYAN + "Started the creation of the years frame.")

    pdf_file_name = r"anos.pdf"
    people_dict = {"nombre": [], "especialidad": [], "b1": [], "b2": [], "b3": [], \
                   "b4": [], "tiempo_total": [], "opos_superadas": []}

    # We read the person_data pdf
    print(Fore.MAGENTA + "Reading the pdf file" + Fore.YELLOW + f" {pdf_file_name}.")

    # We read the pdf that contains the marks
    file = open(pdf_file_name, "rb")
    pdfReader = PyPDF2.PdfFileReader(file)
    n_pages = pdfReader.numPages

    people_dict = {"nombre": [], "especialidad": [], "b1": [], "b2": [], "b3": [], \
                   "b4": [], "tiempo_total": [], "opos_superadas": []}

    for n_page in range(n_pages):
        page_read = pdfReader.getPage(n_page)
        text = page_read.extractText()

        page = Line_Builder(text)

        if page.header == None:
            print(Fore.GREEN + f"Page: {n_page + 1}" + Fore.RED + " ERROR.")
        else:
            df = tabula.read_pdf("anos.pdf", pages = n_page + 1)[0]
            df_col_0 = str(list(df.iloc[:, 0]))

            if len(df.keys()) == 1:
                match = re.compile(r"(?<!:)\d,\d{4}").findall(df_col_0)
                match.insert(0, "0,0000")
                years = []
                opos_sup = []

                for k, val in enumerate(match):
                    if (k)%5 == 0:
                        opos_sup.append(val)
                    else:
                        years.append(val)
                opos_sup = opos_sup[1:]
                
            if len(df.keys()) == 2:
                years = re.compile(r"(?<!:)\d,\d{4}").findall(df_col_0)
                opos_sup = re.compile(r"\d,\d{4}").findall(str(df.iloc[:, 1]))
            elif len(df.keys()) == 3:
                years = re.compile(r"\d,\d{4}").findall(str(df.iloc[:, 1]))
                opos_sup = re.compile(r"\d,\d{4}").findall(str(df.iloc[:, 2]))

                if len(years) != len(opos_sup):
                    years2 = re.compile(r"(?<!:)\d,\d{4}").findall(df_col_0)
                    for k, val in enumerate(years2):
                        years.insert(k*4, val)

            for k, opos in enumerate(opos_sup):
                total = 0
                for i, val in enumerate(years[k*4: k*4 + 4]):
                    val = val.replace(",", ".")
                    if i == 0:
                        total += float(val)/0.6
                        people_dict["b1"].append(float(val))

                    if i == 1:
                        total += float(val)/0.3
                        people_dict["b2"].append(float(val))
                    if i == 2:
                        total += float(val)/0.3
                        people_dict["b3"].append(float(val))
                    if i == 3:
                        total += float(val)/0.15
                        people_dict["b4"].append(float(val))
                    
                people_dict["tiempo_total"].append(total)
                people_dict["opos_superadas"].append(float(opos.replace(",", ".")))

            dni = re.compile(r"[*]+\d+[*]+|\d{8}[a-zA-Z]").findall(df_col_0)
            names = re.compile(r"(?:[*]+\d+[*]+|\d{8}[A-zÀ-ÿñÑ])\s?(?:\s?[A-zÀ-ÿñÑ]+[ª]?[ -]?[.]?[,]?)+").findall(df_col_0)
            especialidad = re.search(r"([A-zÀ-ÿñÑ]+\s?)+", re.search(r"Especialidad(.*)", df_col_0).group(1)).group(0)

            if len(dni) != len(names):
                try:
                    for k in range(len(dni)):
                        if dni[k] not in names[k]:
                            index = df[df[df.keys()[0]].str.contains(dni[k], regex=False).fillna(False)].index.values[0]
                            name_before = re.search(r"\D+", df.iloc[index - 1, 0]).group(0)
                            dni_fix = re.search(r"[*]+\d+[*]+|\d{8}\w", df.iloc[index, 0]).group(0)
                            name_after = re.search(r"\D+\s{0,2}", df.iloc[index + 1, 0]).group(0)
                            new_name = dni_fix + " " + name_before + " " + name_after
                            names.insert(k, new_name)

                except IndexError:
                    for k in range(len(dni) - len(names)):
                        #names.insert(len(names) - k, f"ERROR en página: {n_page + 1} con dni: {dni[len(dni) - 1 - k]}")

                        index = df[df[df.keys()[0]].str.contains(dni[len(dni) - 1 - k], regex=False).fillna(False)].index.values[0]
                        name_before = re.search(r"\D+", df.iloc[index - 1, 0]).group(0)
                        dni_fix = re.search(r"[*]+\d+[*]+|\d{8}\w", df.iloc[index, 0]).group(0)
                        name_after = re.search(r"\D+\s{0,2}", df.iloc[index + 1, 0]).group(0)
                        new_name = dni_fix + " " + name_before + " " + name_after
                        names.insert(len(dni) - k, new_name)

            for k in names:
                people_dict["nombre"].append(k)
                people_dict["especialidad"].append(especialidad)
            print(Fore.GREEN + f"Page: {n_page + 1}" + Fore.YELLOW + " Complete.")

    return people_dict

def main():
    frame = create_years_frame()
    df = pd.DataFrame(frame)
    df.to_excel("years.xlsx", index = False)

if __name__ == "__main__":
    main()