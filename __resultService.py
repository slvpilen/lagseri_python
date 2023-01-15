# Her lages classer og funksjoner, som ikke er ferdi laget og brukes IKKE i lagserie.py (main()) enda.


from math import log10


class resultService:
    def __init__(self):
        # if: gir vanlig, else for 5-kamp
        self.allLifter = {}


    def addNewResult(self, data):
        if data[7].replace("-","0").split(".",1)[0].isdigit():   # Vanlig protokoll
            self.name = data[5]

        elif not data[7].replace("-","0").split(".",1)[0].isdigit():   # 5-kamp
            self.name = data[6]
           

        # dersom sette er tomt
        if len(self.allLifter) == 0:
            l1 = lifter(data)
            self.allLifter = {l1}

        # dersom løfteren har resultat fra før
        elif any(lifter_.name == self.name for lifter_ in self.allLifter):
            # Dersom løfteren har et resultat fra før kommer man hit og man er avhengig av å finne objektet
            found_lifter = next((lifter for lifter in self.allLifter if lifter.name == self.name), None)
            found_lifter.addResult(data)
        
        # ellers (tenkt for dersom sette ikke er tomt og løfteren ikke har resultat fra før)
        else: 
            self.allLifter.add(lifter(data))

        

class lifter:
    def __init__(self, data):
        self.stevner = []
        self.stevner.append(stevne(data))


        if "k" in data[2].lower():
            self.gender = "f" 
        else:
            self.gender = "m"

        # if: gir vanlig, else for 5-kamp
        if data[7].replace("-","0").split(".",1)[0].isdigit():   # Vanlig protokoll
            self.born = data[3]
            self.name = data[5]
            self.club = data[6].lower()

        elif not data[7].replace("-","0").split(".",1)[0].isdigit():   # 5-kamp
            self.born = data[4]
            self.name = data[6]
            self.club = data[7].lower()

    
    def addResult(self, data):
        self.stevner.append(stevne(data))
        

    def getBestSinclaire(self, men_poeng=False):
        self.best_sinclair = float()
        for self.stevne in self.stevner:
            if self.stevne.sinclair_point() > self.best_sinclair:
                self.best_sinclair = self.stevne.sinclair_point(men_poeng)
        return self.best_sinclair
        

    


class stevne:
    # Burde få inn stevne dato her, må hentes fra når excelsheeten lastes
    def __init__(self, data):
        self.bw = float(data[1])
        self.category = data[2]

        if "k" in self.category.lower():
            self.gender = "f" 
        else:
            self.gender = "m"

        # if: gir vanlig, (elif for 5-kamp)
        if data[7].replace("-","0").split(".",1)[0].isdigit():   # Vanlig protokoll
            self.attempts = [int(elem.replace("-","0").split(".",1)[0]) for elem in data[7:13]]

        #elif: for 5-kamp
        elif not data[7].replace("-","0").split(".",1)[0].isdigit():   # 5-kamp
            self.attempts = [int(str(elem).replace("-","0").split(".",1)[0]) for elem in data[8:14]]

    def get_total(self):
        self.snatch = self.attempts[0:3]
        self.cnj = self.attempts[3:6]

        if len(self.snatch) == 0 or len(self.cnj) == 0:
            return False
        
        self.good_snatch = False
        self.good_cnj = False

        self.best_snatch = 0
        self.best_cnj = 0
        for self.elem in self.snatch:
            if self.elem >= 1:
                self.good_snatch = True
                self.best_snatch = float(self.elem)
        for self.elem in self.cnj:
            if self.elem >= 1:
                self.good_cnj = True
                self.best_cnj = float(self.elem)

        if not self.good_snatch or not self.good_cnj:
            return False

        return self.best_snatch + self.best_cnj

    def sinclair_point(self, men_poeng=False):  # dersom True i parameter: blir det herrepoeng, uansett kjønn!
        self.men_poeng = men_poeng
        if self.gender.lower() == "m" or self.men_poeng == True and self.get_total()!=False:
            poeng = self.get_total()*(10**(0.751945030*((log10(self.bw/175.508))**2)))
            return poeng

        elif self.gender.lower() == "f" and self.get_total()!=False:
            poeng = self.get_total()*(10**(0.783497476*((log10(self.bw/153.655))**2)))
            return poeng
        else:
            return -1  # No valid result (need 1 good attempt in snatch and cnj to get a total)!
    




"""
Det er mye mer lesbart å bruke classene over, enn slik lagserie.py er idag
Det er også mye nklere å få ut annen data, dersom nye behov oppstår i fremtiden! 

Merk: Kun de resultatene soim lastes inn er i resultater, altså vil kun stvner i kvallikperiode bli med

Her er eksempelkjøring med classene. 
"""
data1 = ['76', '71.94', 'JK', '2004-03-26 00:00:00', 'None', 'Marte Walseth', 'Nidelv IL', '54.0', '-57.0', '57.0', '70.0', '74.0', '-77.0', '=IF(MAX(H10:J10)<0,0,TRUNC(MAX(H10:J10)/1)*1)']
data2 = ['76', '71.94', 'JK', '2004-03-26 00:00:00', 'None', 'Fredd fobs', 'Nidelv IL', '60.0', '-77.0', '87.0', '90.0', '104.0', '-117.0', '=IF(MAX(H10:J10)<0,0,TRUNC(MAX(H10:J10)/1)*1)']
data3 = ['76', '74.94', 'JK', '2004-03-26 00:00:00', 'None', 'Marte Walseth', 'Nidelv IL', '100.0', '-101', '102.0', '130.0', '160.0', '-150.0', '=IF(MAX(H10:J10)<0,0,TRUNC(MAX(H10:J10)/1)*1)']


resultater = resultService()
resultater.addNewResult(data1)
print(resultater.allLifter)

resultater.addNewResult(data2)
resultater.addNewResult(data3)



alle_loftere = resultater.allLifter
for lofter in alle_loftere:
    print(lofter.getBestSinclaire(), lofter.name)
    


