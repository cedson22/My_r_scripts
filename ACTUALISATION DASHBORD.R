setwd("C:/Users/mmala/Downloads/Docs EPC 14 août 2023/Suivi collecte/DASH BORD EPC RGPH-5")

library(pacman)
p_load(tidyverse,expss,readxl,googlesheets4,googledrive,datefixR,expss, questionr, haven, labelled,googlesheets4)


gs4_auth(email = "malandamway@gmail.com")


 BD_Dashbord_Menages <- read_excel("BD_Dashbord.xlsx", 
                    sheet = "CIMEN_GRPMEN")
 # gs4_create(
 # "BD_Dashbord_Menages",
   sheets = BD_Dashbord_Menages
   #)
 
 ss <- "1jFg9R3jw0OcsczxaGVBxJTwMyfL-ftQ3lYWuIj9iDyU"
 sheet_write(BD_Dashbord_Menages, ss = ss, #"18hlbIZx0tojEKSgGYORFn8Jj3FUCcB6T8_kcZoE2Y4M",
             sheet = "BD_Dashbord_Menages")
 
  write.csv(BD_Dashbord_Menages, file = "MenagesComplet.csv")
 drive_upload("MenagesComplet.csv")
 #ss <- "1Dg417km_N6cfr6JbG5wT9qbEqCf7hb3DdYRUYFZ3zvU"

 #drive_trash("BD_Dashbord_Menages")
#write.csv(BD_Dashbord_Menages, file = "BD_Dashbord_Menages.csv"
 
#drive_create("BD_Dashbord_Menages",type = "spreadsheet" )

 
 
 BD_Dashbord_Individus_vf <- read_excel("BD_Dashbord.xlsx", 
                                    sheet = "CIMMO")
#gs4_create(
#  "BD_Dashbord_Individus",
#  sheets = BD_Dashbord_Individus
#) 

ss ="1TSgJ-WX298DexDcZa3EmWyYoj3nJwXvz_2v632V5r9Q"
sheet_write(BD_Dashbord_Individus_vf, ss = ss, #"18hlbIZx0tojEKSgGYORFn8Jj3FUCcB6T8_kcZoE2Y4M",
            sheet = "BD_Dashbord_Individus_vf")





COMPA <- read_excel("BD_Dashbord.xlsx",  sheet = "COMPARAISON")

#gs4_create(
#  "COMPARAISON",
#  sheets = COMPARAISON
#) 

ss = "1pVNeXPjQRzJwNYg7M-YYJlpkKzFRoO6xYKnAaqL6lsE"
  
  sheet_write(COMPA, ss = ss,  
              sheet = "COMPA")




#drive_auth()

#driv_fin()
# drive_rm(wordstar, b4xl, execuvision)

