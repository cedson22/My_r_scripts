# Définition du chemin du repertoire de travail---- 
 rm(list = ls())
  setwd("~/Cédric/EDS III/Dashboard/Cartographie")

#Importation des packages nécessaires ----
library(pacman)
p_load(tidyverse,expss,readxl,googlesheets4,googledrive,datefixR,expss, questionr, haven, labelled,googlesheets4,openxlsx)

# Définition du mail a utiliser pour utiliser le Drive----
gs4_auth(email = "malandamway@gmail.com")

# Importation des données ----

d <- read_excel("0.Base_Carto_EDS.xlsx") #Données de terrain
d0 <- read_excel("1.MenageRGPH5.xlsx") # Information Grappe RGPH-5
d1 <- MenageRGPH5 <- read_excel("1.MenageRGPH5.xlsx",sheet = "EQUIP")# Info équipe RGPH5
#View(d1)

# Sélection des variables d'intérêt ----
d <- d %>%  select(LCLUSTER,LINTNUM,LREGION,LSTATE,LDISTRICT,LSEGNUM,LSEGHH,
                   LTOTHH,LDATE,LDATEFIN,LTRUEHH,LNUMBER,LSTRUCT,LSTYPE,
                   LHOUSEH,LINTGPS,LLATITUDE,LLONGITUDE,LSTRUCTT,LINFRAS
)

# Renommage des variables ----
names(d) <- c("GRAPPE","CODE_AGENT","DEPARTEMENT","DISTRICT_COMMUNE","CU_CR",
              "SEGMENT_SELECTED","MEN_DANS_GRAPPE_SEGMENT",
              "TOTAL_MEN_DANS_ZD_SEGMENTEE","DATE_DEBUT_GRAPPE","DATE_CLOTURE_GRAPPE",
              "MEN_USAGE_HABITATION","NUM_MENAGE","NUM_PARCELLE","TYPE_PARCEL_BAT",
              "MENAGES_DANS_PARCELLE","CAPTURE_COOR_GPS","LATITUDE","LONGITUDE",
              "STATUT_OCCUPATION_PARCELLE" ,"TYPE_INFRASTRUCTURE")

## Création de la variable équipe----

d <- d %>% mutate( Equipe =
                     case_when(GRAPPE %in% 1:16 ~ "EQUIPE 1",
                               GRAPPE %in% c(17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32) ~ "EQUIPE 2",
                               GRAPPE %in% c(33,34,35,36,37,38,39,40,41,42,43,44,45,46,52,53,61)~ "EQUIPE 3",
                               GRAPPE %in% c(47,48,49,50,51,54,55,56,57,58,59,60,62,63,64,65)~"EQUIPE 4 "	,
                               GRAPPE %in% c(66,67,68,69,70,71,72,73,74,86,87,88,89,90,91)~"EQUIPE 5" 	,
                               GRAPPE %in% c(75,76,77,78,79,80,81,82,83,84,85,92,93,94,95)~ "EQUIPE 6"	,
                               GRAPPE %in% c(108,109,110,111,112,113,114,115,116,118,119,120,121,128,129,130,131,127)~ "EQUIPE 7" 	,
                               GRAPPE %in% c(96,97,98,99,100,101,102,103,104,105,106,107,117,122,123,124,125,126)~ "EQUIPE 8" 	,
                               GRAPPE %in% c(132,133,134,135,136,137,138,139,152,153,154,155,156,157,158,159,160,161)~ "EQUIPE 9"	,
                               GRAPPE %in% c(140,141,142,143,144,145,146,147,148,149,150,151,162,163,164,165,166)~ "EQUIPE 10" 	,
                               GRAPPE %in% c(167,168,169,170,171,172,173,174,175,176,177,178,179,193)~ "EQUIPE 11" 	,
                               GRAPPE %in% c(180,181,182,183,184,185,186,187,188,189,190,191,192,194,195,196)~ "EQUIPE 12" 	,
                               GRAPPE %in% c(203,204,205,206,207,208,209,218,219,220,221,225,226,222,202)~ "EQUIPE 13" 	,
                               GRAPPE %in% c(197,198,199,200,201,210,211,212,213,214,215,216,217,223,224)~ "EQUIPE 14" 	,
                               GRAPPE %in% c(227,228,229,230,231,232,233,234,235,236,244,245,246,247,254,255,256)~ "EQUIPE 15" 	,
                               GRAPPE %in% c(237,238,239,240,241,242,243,248,249,250,251,252,253)~ "EQUIPE 16" 	,
                               GRAPPE %in% c(257,258,259,260,261,262,263,264,265,266,267,268,269,270,285,286)~ "EQUIPE 17" 	,
                               GRAPPE %in% c(271,272,273,274,275,276,277,278,279,280,281,282,283,284)~ "EQUIPE 18" 	,
                               GRAPPE %in% c(287,288,289,290,291,292,293,294,295,296,297,298,299,300,301,313,314)~ "EQUIPE 19" 	,
                               GRAPPE %in% c(302,303,304,305,306,307,308,309,310,311,312,315,316,317,318,319,320)~ "EQUIPE 20" 	,
                               GRAPPE %in% c(328,329,330,331,332,356,357,358,359,360,361,364,365)~ "EQUIPE 21" 	,
                               GRAPPE %in% c(321,322,323,324,325,326,327,333,334,335,336,352,353,354,355,362,363)~ "EQUIPE 22" 	,
                               GRAPPE %in% c(337,338,339,340,341,342,343,344,345,346,347,348,349,350,351)~ "EQUIPE 23" 	,
                               GRAPPE %in% c(366,367,368,369,373,374,375,376,377,378,379,380,398,399,400,401,402,403,404,405,406)~ "EQUIPE 24" 	,
                               GRAPPE %in% c(370,371,372,381,382,383,384,385,386,387,388,389,390,391,392,393,394,395,396,397)~ "EQUIPE 25" 	,
                               GRAPPE %in% c(417,418,419,420,421,422,423,424,425,426,427,428,429,430,431,432,433)~ "EQUIPE 26" 	,
                               GRAPPE %in% c(434,407,408,409,410,411,412,413,414,415,416,435,436,437,438)~ "EQUIPE 27" 	,
                               GRAPPE %in% c(456,457,458,459,460,461,462,463,464,465,466,467,468,469)~ "EQUIPE 28" 	,
                               GRAPPE %in% c(439,440,441,442,443,444,445,446,447,448,449,450,451,452,453,454,455)~ "EQUIPE 29" 	,
                               GRAPPE %in% c(470,471,472,473,474,475,476,477,478,479,480,481,482,483,484,485,486,487,488,489,490,491,492,493,494,495,496,497,498,499,500)~ "EQUIPE 30",
                               
                               TRUE~ NA
                     )
)
# Ajout des informations du RGPH5 sur les ménages et grappes

d <- d %>% left_join(d0) %>% left_join(d1,by=c("EQUIPE_" = "EQUIPE"))

# Convertion des dates de début et de clôture de la grappe en format date ----
d <- d %>%
  mutate(DATE_START_GRAPPE = format(ymd(as.character(d$DATE_DEBUT_GRAPPE)), "%Y-%m-%d"))

d$DATE_START_GRAPPE <- ymd(as.character(d$DATE_START_GRAPPE))

# d <- d %>% mutate(DATE_CLOTURE_GRAPPE_T = na_if(DATE_CLOTURE_GRAPPE, "0")) %>%
#   mutate(DATE_END_GRAPPE = format(ymd(DATE_CLOTURE_GRAPPE_T), "%Y/%m/%d"))

d <- d %>% mutate(DATE_CLOTURE_GRAPPE_T = na_if(DATE_CLOTURE_GRAPPE, "0")) %>%
  mutate(DATE_END_GRAPPE = format(ymd(DATE_CLOTURE_GRAPPE_T), "%Y/%m/%d"))


d$DATE_END_GRAPPE <- ymd(as.character(d$DATE_END_GRAPPE))

d <- d %>%  filter(DATE_START_GRAPPE  > "2025-05-31")#Suppresion des grappes antérieures au 1er juin 2025

#d %>% mutate(DATE_END_GRAPPE_ = ifelse(is.na(DATE_END_GRAPPE),DATE_START_GRAPPE, DATE_END_GRAPPE))

# Recodage des modalités des variables ----

d$STATUT_OCCUPATION_PARCELLE <- recode(d$STATUT_OCCUPATION_PARCELLE ,
                                       1~ "Parcelle habitée",
                                       2~ "Parcelle non habité",
                                       3~ "Infrastructure administrative")

d$TYPE_PARCEL_BAT <- recode(d$TYPE_PARCEL_BAT,
                            1 ~	"Maison (logement) isolée",	
                            2 ~	"Maison à plusieurs logements", 	
                            3	~ "Immeuble (étages) appartements",	
                            4	~ "Concession/Saré")

d$MEN_USAGE_HABITATION <- recode(d$MEN_USAGE_HABITATION,
                                 1 ~ "Oui",
                                 2 ~ "Non")

d$TYPE_INFRASTRUCTURE <- recode(d$TYPE_INFRASTRUCTURE,
                                1 ~ "Infrastructure Scolaire",
                                2	~ "Infrastructure Sanitaire",	
                                3	~ "Administration publique",	
                                4 ~	"Banque",	
                                5	~ "Marché",	
                                6	~ "Station de pompage",	
                                7	~ "Eglise",	
                                8	~ "Hôtel/Auberge/Motel")

d$CAPTURE_COOR_GPS <- recode(d$CAPTURE_COOR_GPS,
                             1 ~	"Coord.  GPS men. Précédent",
                             2 ~	"Prise coord. Maintenant",
                             3 ~	"Pas prise coord. GPS maintenant",
                             9 ~	"Remplacer coord. Existant")


#Calcul des indicateurs d'anayse par grappes-et par équipe---

## Calcule durée de bouclage d'une grappe ----

d <- d %>% group_by(GRAPPE) %>%
  mutate(DUREE_BOUCLAGE_GRAPPE = case_when(is.na(DATE_END_GRAPPE) ~ NA,
  TRUE ~ as.integer(max(DATE_END_GRAPPE,na.rm = TRUE) - min(DATE_START_GRAPPE,na.rm = TRUE))+1),
   ) %>% ungroup()
 




#Calcul durée de travail de l'équipe----

ref <- today() # Date de référence pour le calcul de la durée de travail
d <- d %>% group_by(Equipe) %>% mutate(DUREE_TRVAIL_EQUIP = as.numeric(ref-min(DATE_START_GRAPPE,na.rm = TRUE))) %>% ungroup()


# d <- d %>% mutate(duree_travail_grapJours = 
#                       case_when(
#                         is.na(dat_debut_grappe) | is.na(dat_cloture_grappe) ~ NA_real_,
#                         TRUE ~ as.numeric(dat_cloture_grappe - dat_debut_grappe) + 1
#                       ))




## Nombre de grappes dénombrées par équipe----

d <- d %>% group_by(Equipe) %>% 
  mutate(Grappes_denombr_equip = n_distinct(GRAPPE)) %>% ungroup()

## Ménages dénombrés par grappe----
d <- d %>% group_by(GRAPPE) %>% 
  mutate(Menage_denombr_grappe = n()) %>% ungroup()

## Nombre de ménages dénombrées par équipe----
d <- d %>% group_by(Equipe) %>% 
  mutate(TotMenages_denombr_equip = n())  %>% ungroup()

## Nombre de ménages dénombrées par jour et par equipe----
d <- d %>% group_by(Equipe,DATE_START_GRAPPE) %>% 
  mutate(Men_Equip_Jour = n())  %>% ungroup()




d_Eq_jr <- d %>% group_by(Equipe) %>%  distinct(DATE_START_GRAPPE,.keep_all = T) %>%
                  select(Equipe,EQUIPE_,DEPARTEMENT,MILRES,DATE_START_GRAPPE,GRAPPE,Men_Equip_Jour) %>% 
                  arrange(Equipe,DATE_START_GRAPPE) #%>% view()


d_Eq_jr$DATE_COLLECTE <- format(ymd(as.character(d_Eq_jr$DATE_START_GRAPPE)), "%Y-%m-%d")
d_Eq_jr$DATE_COLLECTE <- ymd(d_Eq_jr$DATE_START_GRAPPE)

# g <- which(d_Eq_jr$GRAPPE==1) # Indices des valeurs manquantes dans GRAPPE )
# 
# d_Eq_jr$DATE_COLLECTE[d_Eq_jr$DATE_COLLECTE == "2025-06-07" & d_Eq_jr$EQUIPE_ == 1] <- "2025-06-06"


d_Eq_jr <- d_Eq_jr %>% group_by(EQUIPE_) %>% arrange(DATE_START_GRAPPE) %>% 
  mutate(JOUR_COLLECTE = paste0("J", row_number()),
         DatSuivante = lead(DATE_START_GRAPPE),
         DureClotGrapEq = as.numeric(DatSuivante-DATE_START_GRAPPE),
         MeanDureCollectEquip= round(mean(DureClotGrapEq,na.rm=TRUE),0),
         DureClotGrapEqCorr=ifelse(is.na(DureClotGrapEq),MeanDureCollectEquip,DureClotGrapEq),
         TRQ_percent = round(Men_Equip_Jour*100/45.1,1),
         Men_EstimEquip_Jour = round(Men_Equip_Jour/DureClotGrapEqCorr,0),
         TRQ_percentCorr = round(Men_EstimEquip_Jour*100/45.1,1)) %>% # Taux de Réalisation Quotidien
  ungroup() %>% 
  mutate(Men_EstimEquip_Jour_Mean= round(mean(Men_EstimEquip_Jour,na.rm=TRUE),0))  %>%
  group_by(EQUIPE_) %>% 
  mutate(TRQ_percentCorr_Mean= round(Men_EstimEquip_Jour*100/Men_EstimEquip_Jour_Mean,1)) %>%
  ungroup() %>%
  select(DEPARTEMENT,MILRES,EQUIPE_,JOUR_COLLECTE,DATE_COLLECTE,
                       GRAPPE,DureClotGrapEq,DureClotGrapEqCorr,Men_Equip_Jour,
                       Men_EstimEquip_Jour,TRQ_percent,Men_EstimEquip_Jour_Mean,TRQ_percentCorr,TRQ_percentCorr_Mean) %>%
  arrange(EQUIPE_,DATE_COLLECTE) #%>% view()



#View(d_Eq_jr)

# k <- which(is.na(d$DATE_END_GRAPPE)) # Indices des valeurs manquantes dans DATE_END_GRAPPE
# d <- d %>% mutate(DATE_END_GRAPPE_EQUIP= DATE_END_GRAPPE) # Création d'une variable pour la date de fin de grappe par équipe
# d$DATE_END_GRAPPE_EQUIP[k] <- max(d$DATE_START_GRAPPE, na.rm = TRUE) # Remplacement des valeurs manquantes de la date de fin de grappe par la date de début de grappe

# d <- d %>% group_by(Equipe) %>%
#   mutate(TotMenages_denombr_equip = n()
#         )  %>% ungroup()

# Calcul Total grappes segmentées ----

d <- d  %>% group_by(GRAPPE) %>%  mutate(GrapSegmented = ifelse(SEGMENT_SELECTED != 0,1,NA_real_ )) %>% ungroup() #%>% dim()

## Grappes segmentables----

d <- d %>% mutate(Grap_segmentable = ifelse(MEN_RGPH5_GRAP >= 300,"segmentable","Non segmentable"))

#Suppression des variables intermédiaires ----



#Idéal de ménage par équipe par duree de collecte ----
d <- d %>% group_by(EQUIPE_) %>%  mutate(TotIdealMenEq = 45.1*DUREE_TRVAIL_EQUIP) %>% ungroup() # Conversion de la variable GRAPPE en caractère

d <- d %>% select(-c(DATE_DEBUT_GRAPPE,
                  DATE_CLOTURE_GRAPPE_T)) # Suppression des variables intermédiaires

BD_grappe <- d %>% select(GRAPPE,Equipe,EQUIPE_,DEPARTEMENT,MILRES,MEN_DANS_GRAPPE_SEGMENT,
                          TOTAL_MEN_DANS_ZD_SEGMENTEE,MEN_RGPH5_GRAP,Menage_denombr_grappe,Grap_segmentable,GrapSegmented,
                          DATE_START_GRAPPE,DUREE_BOUCLAGE_GRAPPE) %>% 
              distinct(GRAPPE,.keep_all = TRUE)#Analyse selon les grappes
                          
                          
                          


BD_Equipe <- d %>% select(-c(MEN_USAGE_HABITATION,NUM_MENAGE,
                             NUM_PARCELLE,TYPE_PARCEL_BAT,MENAGES_DANS_PARCELLE,
                             CAPTURE_COOR_GPS,LATITUDE,	LONGITUDE,
                             STATUT_OCCUPATION_PARCELLE,	TYPE_INFRASTRUCTURE,
                             Equipe,IDZD,
                             MILRES,MEN_DANS_GRAPPE_SEGMENT,
                             TOTAL_MEN_DANS_ZD_SEGMENTEE,MEN_RGPH5_GRAP,
                             Menage_denombr_grappe,Grap_segmentable,
                            ,DUREE_BOUCLAGE_GRAPPE
                             
                             )) %>%  group_by(GRAPPE) %>% 
  mutate(GrapSegmented = ifelse(SEGMENT_SELECTED != 0,1,NA_real_ )) %>%
  distinct(GRAPPE,.keep_all = TRUE)

BD_Equipe <- BD_Equipe %>% group_by(EQUIPE_) %>% 
                            mutate(
                              GrapDenombreEquipe = n_distinct(GRAPPE),
                              TotGrapSegmentedEq= sum(GrapSegmented,na.rm = TRUE)
                            ) %>% 
  distinct(EQUIPE_,.keep_all = TRUE) %>% select(EQUIPE_,GRAPPE,DEPARTEMENT,DISTRICT_COMMUNE,
                                                CU_CR,TOT_GRAP_EQUIP,GrapDenombreEquipe,
                                                Grappes_denombr_equip,TOT_MEN_EQUIP,
                                                TotMenages_denombr_equip,Men_Equip_Jour,
                                                TotIdealMenEq,TotGrapSegmentedEq,GrapDenombreEquipe,DUREE_TRVAIL_EQUIP)#Analyse selon les équipes                               





# Création de la feuille de calcul dans Google Sheets ----
#gs4_deauth() #Révoquer les identifiants existants (si vous les avez déjà mis en cache)
#gs4_auth(scope = "https://www.googleapis.com/auth/spreadsheets")#Révoquer les identifiants existants (si vous les avez déjà mis en cache)

# Ecriture des données dans Google Sheets ----
# gs4_create(
# "EDSCIII_CARTO",sheets = "data"
# )

# gs4_create(
# "EDSCIII_analyse_grappes",sheets = "grappes"
# )

# gs4_create(
#   "EDSCIII_analyse_equipe",sheets = "Equipes"
# )

# gs4_create(
#   "EDSCIII_aPerformances_equipe",sheets = "Performances"
# )


# Lister les fichiers contenus dans le Google Drive-----
#drive_find(type = "spreadsheet") #Utilisation du package gs4
#drive_ls() # Utilisation du package googledrive

#Récupération de l'id drive du fichier google sheet EDSCIII_CARTO créé---
# ID <- drive_find(type = "spreadsheet")
# ss <- ID$id[2] # Récupération de l'id du fichier créé

# Écriture des données dans la feuille de calcul Google Sheets ----

#sheet_write(d, ss = "1z3B8Vhe-N_RtXD5KIQ-DuCF2H88QFI9YMnxZx9xb35A", #"1z3B8Vhe-N_RtXD5KIQ-DuCF2H88QFI9YMnxZx9xb35A"
#sheet = "data")

sheet_write(BD_Equipe, ss = "1A-14zKgT8AeXEOZyLzDdYCICzvD-zOJzmKdFSt6jR-M",
            sheet = "Analyse_equipes")
sheet_write(BD_grappe, ss = "1G9Y0mCrb1PQWg7btNU4VG6rkk-FAzB46h_Mc70y2uEQ",
            sheet = "Analyse_grappes")
sheet_write(d_Eq_jr, ss = "144G-A4ke2wwUgiYZOSPX2h9A5O0aC04wOjUIX-uzYCY",
            sheet = "Performances")

# Exporter la base en excel----
##dir.create("Sous_Bases_Traitées") # Création du dossier pour stocker les fichiers Excel

write.xlsx(d, file = "Sous_Bases_Traitées/Carto_dashboard_ménages.xlsx", sheetName = "Analyse_Ménages", rowNames = FALSE)
write.xlsx(BD_grappe, file = "Sous_Bases_Traitées/Carto_dashboard_grappes.xlsx", sheetName = "Analyse_grappes", rowNames = FALSE)
write.xlsx(BD_Equipe, file = "Sous_Bases_Traitées/Carto_dashboard_equipe.xlsx", sheetName = "Analyse_Equipes", rowNames = FALSE)
write.xlsx(d_Eq_jr, file = "Sous_Bases_Traitées/Carto_dashboard_performances.xlsx", sheetName = "Performances_Equipes", rowNames = FALSE)

sheet_write(d, ss = "1z3B8Vhe-N_RtXD5KIQ-DuCF2H88QFI9YMnxZx9xb35A", #"1z3B8Vhe-N_RtXD5KIQ-DuCF2H88QFI9YMnxZx9xb35A"
            sheet = "data")








# ANALYSEZD <- read_excel("EPC_DASHBORD_FINAL_1.xlsx", 
#                             sheet = "COMPARAISON")

#gs4_create(
# "ANALYSEZD",
#  sheets = ANALYSEZD
#)


# ss <- "1Bn4InRQ5wMewcDtnbsfycX5JCk4fOsxBVfexNoJeSN4"
# sheet_write(ANALYSEZD, ss = ss, #"18hlbIZx0tojEKSgGYORFn8Jj3FUCcB6T8_kcZoE2Y4M",
#             sheet = "ANALYSEZD")
# 
# 
# 
# 
# ANALYSEINDIVIDU <- read_excel("EPC_DASHBORD_FINAL_1.xlsx", 
#                             sheet = "CIMMO")

#gs4_create(
# "ANALYSEINDIVIDU",
# sheets = ANALYSEINDIVIDU

#)


# ss <- "1BmMTzmekIpjSaihC_EK7zb18WlTazd0uGadbF6Q167U"
# sheet_write(ANALYSEINDIVIDU, ss = ss, sheet = "ANALYSEINDIVIDU")





