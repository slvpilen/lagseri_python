"""   lagseriereultat-calculater   """


from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
import os 


#  pip install openpyxl
#  pip install pillow   # for filer med bilder


def excel_files_in_folder(direction):
    folder = os.listdir(direction)
    excel_files = []
    for file in folder:
        if file.endswith(".xlsx"):
            excel_files.append(file)
    return excel_files


def check_dato_in_qualification(dato_cell, start_date, end_date):  # denne er trøbbel!!! feil med format dato fra excel OG sjekk om dato er minder større enn qualicperiode
    correctDate = None
    try:
        dato_cell = dato_cell[8:10] + dato_cell[5:7] + dato_cell[0:4]
        newDate = datetime(int(dato_cell[6:11]),int(dato_cell[3:5]),int(dato_cell[0:2]))
        start = datetime(int(start_date[6:11]),int(start_date[3:5]),int(start_date[0:2]))
        end = datetime(int(end_date[6:11]),int(end_date[3:5]),int(end_date[0:2]))
        correctDate = True
    except ValueError:
        correctDate = False
        #present = datetime.now()
    if correctDate and start <= newDate <= end:
            return True
    else:
        if not correctDate:
            return "unvalid dato in sheet"
        else:
            return "date not in qualification periode"


def check_sheet(wb, filename, start_date, end_date):
    ok_sheets = []
    bad_sheets =  []

    mal_nvf = ['Vekt-', 'Kropps-', ' Kate-', 'Fødsels-', 'St', 'Navn', 'Lag', None, 'Rykk', None, None, 'Støt', None, '    Beste forsøk i', None, 'Sammen-', 'Poeng', 'Poeng', 'Pl.', 'Rek.', 'Sinclair Coeff.']

    sheets_names = wb.sheetnames

    for sheet in sheets_names:
        ws = wb[sheet] 
        added = 0

        seventh_row = [ws[str(get_column_letter(char)) + "7"].value for char in range(1, 22)]  # 1-21/ A-U

        if seventh_row == mal_nvf:  # check if the 7th row is excatly same as the "mal"
            dato_cell = str(ws['R5'].value)
            print(dato_cell)
            dato_in_ok_periode = check_dato_in_qualification(dato_cell, start_date, end_date)
            if dato_in_ok_periode != True:
                bad_sheets.append([sheet, f"right format, but {dato_in_ok_periode}"])
                continue

            for rad in range(9, 30):
                row = [ws[str(get_column_letter(char)) + str(rad)].value for char in range(1, 3)]

                try:
                    if str(row[0]).replace(".", "").replace(",","").replace("+","").isdigit() and str(row[1]).replace(".", "").replace(",","").replace("+","").isdigit():
                        ok_sheets.append(sheet)
                        added = True
                        break
                except:
                    continue


            if not added:
                bad_sheets.append([sheet, "right format, but no data"])
                
        else:
            bad_sheets.append([sheet, "wrong formate"])

        """
    Printing info about sheets passed and not passed check
    """        
    print( f"This sheets in file {filename} is ok:")
    for sheet in ok_sheets:
        print(sheet)
    print("\nThis is the bad ones:")
    for sheet in bad_sheets:
        print(sheet[0].ljust(15), "\t", "(", sheet[1], ")")
    print("-"*46)
    return ok_sheets, bad_sheets




    


def every_result(filename, direction, start_date, end_date, club):
    dic = {}

    """
    open workbook and find sheets
    """
    wb = load_workbook(filename)  # her må filen være i aktuell mappe, legg til direction slik at åpner folder
    sheets_names = wb.sheetnames
    ws = wb[sheets_names[0]]  #    wb[sheets[1]]

    ok_sheets, bad_sheets = check_sheet(wb, filename, start_date, end_date)
    """
    decide whats lifting data and append it to a list
    """
    #this is just copy of another work
    # while  dato:  

    #     if row == 1:
    #         liste.append(["Dato", f"Close-{filename}"])
    #     else:
    #         cell_dato = char_dato + str(row)
    #         cell_close = char_close + str(row)

    #         dato = ws[cell_dato].value
    #         close = ws[cell_close].value
            
    #         liste.append([dato, close])

    #     row+= 1
    
    # return liste[:-1]

    
    # print(ws['A10'].value)
    # print(ws['A11'].value)
    # print(ws['A12'].value)
    # print(ws['A13'].value)
    # print(ws['A14'].value)


"""
    char_dato = "A"  #column_finder("Dato", ws)
    char_close = "B"  #column_finder("Siste", ws)

    cell_dato = "A1"
    cell_close = "A1"
    close = 100000
    dato = "21.01.20"
    row = 1

    while  dato:  

        if row == 1:
            liste.append(["Dato", f"Close-{filename}"])
        else:
            cell_dato = char_dato + str(row)
            cell_close = char_close + str(row)

            dato = ws[cell_dato].value
            close = ws[cell_close].value
            
            liste.append([dato, close])

        row+= 1
    
    return liste[:-1]
"""



def info_from_user():
    def check_date(spm):
        svar = input(spm)
        if len(svar) != 10:
            print(f"{svar} is not a valid input. Try again! 1")
            check_date(spm)
        elif len(svar.replace(".","")) != 8:
            print(f"{svar} is not a valid input. Try again! 2")
            check_date(spm)
        elif svar[0:2].isdigit and svar[3:5].isdigit and svar[6:9].isdigit:
            # Check if input date is in the past:
            correctDate = None
            try:
                newDate = datetime(int(svar[6:11]),int(svar[3:5]),int(svar[0:2]))
                correctDate = True
            except ValueError:
                correctDate = False
            present = datetime.now()
            if correctDate and newDate <= present:
                return svar
            elif correctDate and newDate > present:
                print(f"{svar} is in the future. If you want to calculato until today, use todays dato!")
                check_date(spm)
            else:
                print(f"{svar} is not a valid input. Try again! 2")
                check_date(spm)
        else:
            print(f"{svar} is not a valid input. Try again! 3")
            check_date(spm)
            

    start_date = check_date("Write start date (dd.mm.yyyy): ")
    end_date =check_date("Write end date (dd.mm.yy): ")
    club = input("Enter the team/club you want to check. This need to be correct, no check!: ")

    print("-"*46)
    return start_date, end_date, club



def main():
    # Intro text
    print("This is a program that calculate team results for a periode","\n", "-"*46)

    start_date, end_date, club = info_from_user()
    #print(start_date, end_date, club)

    direction = os.path.dirname(os.path.realpath(__file__)) #r"C:\Users\oskar\Documents\lagserie_python"
    filenames = excel_files_in_folder(direction)

    #every_results er tenkt å ta inn en og en excel fil og legge beste res til hver pers i dic. 
    #Denne må derfor loopes igjennom alle excel-filene og dic blir oppdater (IKKE returnert ny)
    all_results_dic = every_result(filenames[0], direction, start_date, end_date, club)
    #print(liste)


if __name__ == "__main__":
    main()