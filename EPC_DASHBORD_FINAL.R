setwd("C:/Users/mmala/Downloads/Docs EPC 14 août 2023/Suivi collecte/DASH BORD EPC RGPH-5")

library(pacman)
p_load(tidyverse,expss,readxl,googlesheets4,googledrive,datefixR,expss, questionr, haven, labelled,googlesheets4)


gs4_auth(email = "malandamway@gmail.com")

ANALYSEMENAGE <- read_excel("EPC_DASHBORD_FINAL_1.xlsx", 
                                  sheet = "CIMEN(MENAGE)")

#gs4_create(
# "ANALYSEMENAGE",
#sheets = ANALYSEMENAGE
# )


ss <- "1860rErDMSp9lmDGyjo2sI8Q0mtuT3J34dSdAcI2ANok"
sheet_write(ANALYSEMENAGE, ss = ss, #"18hlbIZx0tojEKSgGYORFn8Jj3FUCcB6T8_kcZoE2Y4M",
            sheet = "ANALYSEMENAGE")



ANALYSEZD <- read_excel("EPC_DASHBORD_FINAL_1.xlsx", 
                            sheet = "COMPARAISON")

#gs4_create(
# "ANALYSEZD",
#  sheets = ANALYSEZD
#)


ss <- "1Bn4InRQ5wMewcDtnbsfycX5JCk4fOsxBVfexNoJeSN4"
sheet_write(ANALYSEZD, ss = ss, #"18hlbIZx0tojEKSgGYORFn8Jj3FUCcB6T8_kcZoE2Y4M",
            sheet = "ANALYSEZD")




ANALYSEINDIVIDU <- read_excel("EPC_DASHBORD_FINAL_1.xlsx", 
                            sheet = "CIMMO")

#gs4_create(
# "ANALYSEINDIVIDU",
# sheets = ANALYSEINDIVIDU

#)


ss <- "1BmMTzmekIpjSaihC_EK7zb18WlTazd0uGadbF6Q167U"
sheet_write(ANALYSEINDIVIDU, ss = ss, sheet = "ANALYSEINDIVIDU")


