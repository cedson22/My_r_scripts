setwd("C:/Users/mmala/Downloads/Docs EPC 14 août 2023/Suivi collecte/DASH BORD EPC RGPH-5")

library(pacman)
p_load(tidyverse,expss,readxl,googlesheets4,googledrive,datefixR,expss, questionr, haven, labelled,googlesheets4)


gs4_auth(email = "malandamway@gmail.com")

EMENAGES_INDIV <- read_excel("VF.xlsx",sheet = "CIMEM (2)")

#gs4_create("EMENAGES_INDIV",sheets = EMENAGES_INDIV)EMEN AGES_INDIV



ss <- "1Y3ylkmOTCWZinMyfs59E8gnCJQy9A0QPHwnxACJFJVg"
sheet_write(EMENAGES_INDIV, ss = ss, sheet = "EMENAGES_INDIV")



COMPA <- read_excel("VF.xlsx",sheet = "COMPARAISON")

#gs4_create( "COMPA",sheets = COMPA)


ss <- "1JWBV4x8ARQUzgEmRZ2Brbd7X30MfU-sPHqpH_Et3XQk"
sheet_write(COMPA, ss = ss,sheet = "COMPA")




ZDMEN <- read_excel("VF.xlsx", sheet = "CIMEM_ZDMEN")

#gs4_create("ZDMEN",sheets = ZDMEN)


ss <- "1Pp2ibJHErUnXVdCAIKFcE_z1IOYxgRAk6Q2rn0XeIZU"
sheet_write(ZDMEN, ss = ss, sheet = "ZDMEN")


