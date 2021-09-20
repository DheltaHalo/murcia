import PyPDF2
import re
import numpy as np
import pandas as pd

from colorama import init, Fore, Back, Style
# Initializes Colorama
init(autoreset=True)

import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

class func(object):
    def element_count(string: str, substring: str):
        return len(string.split(substring)) - 1
    
    def remove_substring(string: str, substring: str):
        index_ini = string.find(substring)
        index_end = index_ini + len(substring)
        string = string[:index_ini] + string[index_end:]

        return string

    def create_dir_tree(folder_name: str):
        path = os.getcwd() + "\\" +folder_name
        dir_list = []
        file_list = []
        final_list = []

        for k in os.walk(path):
            dir_list.append(k[0])
            file_list.append(k[2])
            
        dir_list = dir_list[1:]
        file_list = file_list[1:]

        for i, dir in enumerate(dir_list):
            for file in file_list[i]:
                final_list.append(dir + "\\" + file)

        return final_list

class Academic_Info(object):
    def __init__(self, header: str):
        self.header = header

        self.build_header()
    
    def build_header(self):
        cuerpo_str = "CUERPO DE PROFESORES DE "
        especialidad = "ESPECIALIDAD "
        tribunal = " - TRIBUNAL"

        if cuerpo_str in self.header:
            pass
        else:
            cuerpo_str =  cuerpo_str[:-4]

        self.cuerpo = re.search(cuerpo_str + r"(.*)" + especialidad, self.header).group(1)
        self.especialidad = re.search(especialidad + r"(.*)" + tribunal, self.header).group(1)
        self.tribunal = int(re.compile(r"\d{2}").findall(self.header[self.header.find(self.especialidad):])[0])

        if self.tribunal == 0:
            self.tribunal = 1

# We build the person classes       
class Person(object):
    """
    Creates a "Person" object which contains the information
    of said person.

    Input:
        - DNI
        - Name
        - PartA
        - PartB
        - Average
    """
    def __init__(self, Cuerpo, Especialidad, Tribunal, Orden, Acceso, Nombre, PartA, PartB):
        self.cuerpo = Cuerpo
        self.especialidad = Especialidad
        self.tribunal = Tribunal
        self.orden = Orden
        self.acceso = Acceso
        self.nombre = Nombre
        self.parta = float(PartA)
        self.partb = float(PartB)

        self.create_average()
        self.create_dicot_var()
    
    def create_average(self):
        if self.parta < 1.25 or self.partb < 1.25:
            self.average = ""
        else:
            self.average = (self.parta + self.partb)
        
        if self.parta == -1:
            self.parta = ""
        if self.partb == -1:
            self.partb = ""

    def create_dicot_var(self):
        if self.average == "":
            self.aprobado = 0
        else:
            if self.average < 5:
                self.aprobado = 0
            else:
                self.aprobado = 1

class Person_builder(object):
    def __init__(self, line: str, info: Academic_Info):
        self.line = line
        self.info = info
        self.separator = " - "

        self.build_dni()
        self.build_name()
        self.build_numbers()

    def build_dni(self):
        self.dni = self.line[:9]
        self.line = self.line[len(self.dni) + len(self.separator):]
    
    def build_name(self):
        match = re.compile(r"[a-zA-Z]\d|[a-zA-Z][-]\d|[a-zA-Z][-][-]|" +
                           r"[,]\d|[,][-]\d|[,][-][-]|" +
                           r"[a-zA-Z][.]\d|[a-zA-Z][.][-]|" 
                           r"[a-zA-Z][ª]\d|[a-zA-Z][ª][-]").findall(self.line)[0]

        index = self.line.find(match)
        if "." in match:
            index += 1
        self.name = self.line[:index + 1]
        self.line = self.line[index + 1:]

        self.name = self.dni + " - " + self.name

    def build_numbers(self):
        self.line = self.line.replace(",", ".")
        # We remove the third mark which bugs when reading this kind of PDF
        if func.element_count(self.line, "-") + func.element_count(self.line, ".") > 2:
            error_index = len(self.line) - self.line[::-1].find(".") - 2
            self.line = self.line[:error_index]
        
        # We determine which case we have
        # CASE 1: We don't have any marks registered
        if func.element_count(self.line, "-") == 2:
            self.parta = -1
            self.partb = -1
            self.line = func.remove_substring(self.line, "--")

        # CASE 2: We have one mark registered
        elif func.element_count(self.line, "-") == 1:
            match = re.compile(r"\d[.]\d{4}").findall(self.line)[0]
            if self.line.find("-") < self.line.find("."):
                self.parta = -1
                self.partb = match
            else:
                self.parta = match
                self.partb = -1

            self.line = func.remove_substring(self.line, "-")
            self.line = func.remove_substring(self.line, match)

        # CASE 3: We have both marks registered
        else:
            match = re.compile(r"\d[.]\d{4}").findall(self.line)
            self.parta = match[0]
            self.partb = match[1]

            self.line = func.remove_substring(self.line, self.parta)
            self.line = func.remove_substring(self.line, self.partb)
        
        match = re.compile(r"\d{5}").findall(self.line)[0]
        self.orden = int(match)

        self.line = func.remove_substring(self.line, match)
        self.acceso = re.compile(r"\d").findall(self.line)[0]

    def create_person(self):
        return Person(self.info.cuerpo, self.info.especialidad, self.info.tribunal, self.orden, self.acceso, \
                      self.name, self.parta, self.partb)

# And the reading lines classes
class Line_builder(object):
    def __init__(self, text: str):
        self.text = text
        try:
            self.build_name_line()
        except IndexError:
            self.error_handling()
    
    def build_name_line(self):
        full_dni = False
        if "***" not in self.text:
            full_dni = True
        
        if full_dni:
            if "VºBº" in self.text:
                self.text = self.text.split("VºBº")[0]

            pattern = re.compile(r"\d{8}[a-zA-Z]{1}[ ][-]|[a-zA-Z]{1}\d{7}[a-zA-Z]{1}[ ][-]")
            match = pattern.findall(self.text)
            for k, v in enumerate(match):
                match[k] = v[:-2]

            self.header = self.text.split(match[0])[0]
            self.people_list = []

            for k, v in enumerate(match):
                if v != match[-1]:
                    result = re.search(v + "(.*)" + match[k + 1], self.text)
                    self.people_list.append(v + result.group(1))
                else:
                    self.people_list.append(v + self.text.split(v)[1].split("/")[0])

        else:
            line_list = self.text.split("***")
            self.header = line_list[0]
            self.people_list = []
            for i, line in enumerate(line_list):
                if i == 0:
                    pass
                elif line == line_list[-1]:
                    self.people_list.append("***" + line.split("/")[0][:-2])
                else:
                    self.people_list.append("***" + line)

    def error_handling(self):
        self.header = None
        self.people_list = None

def create_marks_frame():
    # Message
    print(Fore.CYAN + "Started the creation of the marks frame.")

    # Variables
    people_list = []
    file_list = func.create_dir_tree("PRIMERA FASE FUSIONADOS")

    for files in file_list:
        try:
            # We read and store important info from our PDF
            file = open(files, "rb")
            pdfReader = PyPDF2.PdfFileReader(file)
            n_pages = pdfReader.numPages
            name_file = files.split(os.path.dirname(files))[1][1:]

            for page in range(n_pages):
                # We start reading pages and extracting the text
                page_read = pdfReader.getPage(page)
                text = page_read.extractText()

                # We build the people list
                page = Line_builder(text)
                if page.header == None:
                    pass

                else:
                    # We build the academic info
                    opos_info = Academic_Info(page.header)

                    for individual in page.people_list:
                        person = Person_builder(individual, opos_info).create_person()
                        people_list.append(person)

            # We close the file
            file.close()

            print(Fore.GREEN + "COMPLETED: " + Fore.YELLOW + name_file)
        except:
            print(Fore.RED + "ERROR: " + Fore.YELLOW + name_file)

    try:
        print(Fore.CYAN + "Frame created successfully.")
        
        # We create the neccesary dictionaries
        people_keys = Person(0, 0, 0, 0, 0, 0, 0, 0).__dict__.keys()
        people_dict = {key: [] for key in people_keys}
        
        for p in people_list:
            for key in people_keys:
                people_dict[key].append(p.__dict__[key])

        return people_dict

    except UnboundLocalError:
        print(Fore.RED + "No se ha encontrado ningún archivo.")

def main():
    frame = create_marks_frame()
    df = pd.DataFrame(frame)
    df.to_excel("marks.xlsx", index = False)

if __name__ == "__main__":
    main()