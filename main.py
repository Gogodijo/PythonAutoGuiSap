import pyautogui as pya
from tilaus import Tilaus
import tkinter as tk
import pandas as pd
import time


def main():
    tilaukset = []
    pya.PAUSE = 0.03
    file = 'pakasteet.xlsx'
    pvm = '5.1.2020'
    sheet = pd.read_excel(file)
    sheetstring = sheet.to_string()
    wordlist = sheetstring.split()
    for i in range(len(wordlist)):
        if wordlist[i].startswith("s"):
            asnro = wordlist[i]
            eurlava = wordlist[i+2]
            rullakko = wordlist[i+3]
            finlava = wordlist[i+4]
            if eurlava == "NaN":
                eurlava = "0"
            if rullakko == "NaN":
                rullakko = "0"
            if finlava == "NaN":
                finlava = "0"
            eurlava = int(eurlava)
            rullakko = int(rullakko)
            finlava = int(finlava)
            tilaus = Tilaus(asnro, rullakko, eurlava, finlava)
            if(tilaus.onkoPakkasia()):
                tilaukset.append(tilaus)
    
    x, y = VerifyStartPosition()
    pya.click(x, y)
    
    while len(tilaukset) > 0:

        if pya.getActiveWindowTitle() == "Nimetön – Muistio":
            tilaus = tilaukset[0]
            print(tilaus)
            time.sleep(1)
            if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                continue
            pya.typewrite("0018")
            if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                continue
            pya.press('tab')
            if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                continue
            pya.typewrite("5018")
            if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                continue
            pya.press('tab')
            if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                continue
            pya.typewrite(tilaus.asiakasnumero)
            if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                continue
            pya.press('tab')
            if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                continue
            pya.press('tab')
            if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                continue
            pya.press('tab')
            if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                continue
            pya.press('tab')
            if tilaus.rullakot > 0:
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.typewrite('rll')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.typewrite(str(tilaus.rullakot))
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.typewrite('0004')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.typewrite('PAKASTE')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')

            if tilaus.eurolavat > 0:
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.typewrite('eur')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.typewrite(str(tilaus.eurolavat))
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.typewrite('0004')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.typewrite('PAKASTE')

            if tilaus.finlavat > 0:
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.typewrite('fin')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.typewrite(str(tilaus.finlavat))
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.typewrite('0004')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.press('tab')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
                pya.typewrite('PAKASTE')
                if pya.getActiveWindowTitle() != "Nimetön – Muistio":
                    continue
            pya.press('enter')
            tilaukset.pop(0)
        else:
            print("Muistio ei ole aktiivinen")


def VerifyStartPosition():
    imageFound = False
    while not imageFound:
        im2 = pya.locateCenterOnScreen('button.png')
        if im2 != None:
            x, y = im2
            imageFound = True
    return x, y


main()
