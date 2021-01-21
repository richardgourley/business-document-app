from searchspreadsheet import SearchSpreadsheet
from countcolumninstances import CountColumnInstances
from createcertificates import CreateCertificates
from mainmenu import MainMenu

main_menu = MainMenu()

if main_menu.choice == "1":
    SearchSpreadsheet()

if main_menu.choice == "2":
    CountColumnInstances()

if main_menu.choice == "3":
    CreateCertificates()


