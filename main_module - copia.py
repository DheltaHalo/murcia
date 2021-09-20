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
                print(f"\nMarks name and especialidad: {m_name} " + Fore.YELLOW + marks["especialidad"][k])
                print(f"Trying to match: {y_name}" + Fore.YELLOW + years["especialidad"][i])

                if match_string_percentage(marks["especialidad"][k], years["especialidad"][i]) > 0.75:
                    print(Fore.GREEN + "\nMatch")
                    for key in diff_keys:
                        marks[key][k] = str(years[key][i])
                    
                    cond = False
                    break
        if cond:
            print(Fore.RED + "\nNo Match")
        
    marks.to_excel("final.xlsx")

if __name__ == "__main__":
    main()