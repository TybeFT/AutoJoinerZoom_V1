import webbrowser
import datetime
import pyautogui as pyto
import time
from pandas import *

schedule_file = ExcelFile("horario.xlsx")
schedule_dataframe = schedule_file.parse(schedule_file.sheet_names[0]).fillna(0)
schedule_dict = schedule_dataframe.to_dict()

print("[+] Esperando Clase")

#print(schedule_dict)

days = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]

def open_class():
    started = False
    i = 0

    while i < len(schedule_dict["Hora"]):
        indexDay = datetime.datetime.now().weekday()
        weekDay = days[indexDay]
        if schedule_dict[weekDay][i] == 0:
            pass
        else:
            for day in schedule_dict.keys():
                if day == weekDay:
                        hour = schedule_dict["Hora"][i].strftime("%H: %M: %S")
                        link_key = "Enlace" + weekDay
                        link = schedule_dict[link_key][i]

                        while started != True:
                            current_time = takeCurrentTime()
                            if current_time == hour:
                                print("[+] La clase incio correctamente")
                                webbrowser.open(link)
                                started = True
                                #Tiempo para esperar que inicie la clase
                                time.sleep(10)
                                pyto.press('k')
                                time.sleep(2)
                                pyto.write("Presente profe")
                                time.sleep(2)
                                pyto.press("enter")
                            else:
                                continue

        i += 1
        started = False

def takeCurrentTime():
    print("[-] Esperando la siguiente clase")
    return datetime.datetime.now().strftime("%H: %M: %S")



if __name__ == '__main__':
    open_class()

