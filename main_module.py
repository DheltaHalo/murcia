from difflib import SequenceMatcher
from re import findall
import pandas as pd

import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

from colorama import init, Fore, Back, Style
# Initializes Colorama
init(autoreset=True)

def match_string_percentage(str1: str, str2: str):
    return SequenceMatcher(None, str1, str2).ratio()
                
def main():
    marks = pd.read_excel("marks.xlsx")
    years = pd.read_excel("years.xlsx")

    diff_keys = [x for x in years.keys() if x not in marks.keys()]
    for key in diff_keys:
        marks[key] = 0

    for k, m_name in enumerate(marks["nombre"]):
        for i, y_name in enumerate(years["nombre"]):
            cond = True
            if match_string_percentage(m_name, y_name) > 0.85:
                print(Fore.GREEN + "Match")
                print(m_name, y_name)
                print()
                for key in diff_keys:
                    marks[key][k] = str(years[key][i])
                
                cond = False
                break
     
        if cond:
            print(Fore.RED + "No Match")
        
    marks.to_excel("final.xlsx")

if __name__ == "__main__":
    main()