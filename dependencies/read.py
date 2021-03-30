import csv

def storeData(path, auto):
     dta_stat = []
     dta_date = []
     dta_desc = []
     dta_debt = []
     dta_crdt = []
     dta_memb = []
     dta_catg = []
     dta_warn = []
     dta_comm = []
     data = []

     with open(path, newline='') as csvfile:
         reader = csv.reader(csvfile, delimiter=',', quotechar='"')
         for row in reader:
              
               dta_stat.append(row[0])
               dta_date.append(row[1])
               dta_desc.append(row[2])
               if(row[3] == ""):
                    dta_debt.append('0')
               else:
                    dta_debt.append(row[3])

               if(row[4] == ""):
                    dta_crdt.append('0')
               else:
                    dta_crdt.append(row[4])
                    
               dta_memb.append(row[5])
               dta_catg.append('NONE')
               dta_warn.append('NONE')
               dta_comm.append('NONE')

     data.append(dta_stat)
     data.append(dta_date)
     data.append(dta_desc)
     data.append(dta_debt)
     data.append(dta_crdt)
     data.append(dta_memb)
     data.append(dta_catg)
     data.append(dta_warn)
     data.append(dta_comm)
     data[7][0] = auto
     return data
     
