# modify_caption("**Tableau 1: age moyen des patients selon le sexe**")
rm(list = ls())
options(encoding ="latin1")
setwd("C:/Users/mmala/Documents/SOUTENANCE AHOUYA/Traitement Ahouya")
library(here)
library(pacman)

p_load(rio,readxl,openxlsx,tidyverse, expss,here,haven,labelled,tables,gtsummary,skimr)

#Importation de la base des données
d1 <- read_csv("Exported.csv")
d1 <- d1 %>% mutate(Q2B = as.character(Q2B),Q2C = as.character(Q2C),
             Q2E = as.character(Q2E),Q2F = as.character(Q2F),
             Q2G = as.character(Q2G),Q2GG = as.character(Q2GG),
             Q6_2 = as.character(Q6_2),Q6_3 = as.character(Q6_3),
             Q6_4 =as.character(Q6_4)
             )
#d <- read_sav("Cancer.sav")
#d <- readRDS("C:/Users/mmala/Documents/SOUTENANCE AHOUYA/Traitement Ahouya/cance.rds")


# Importation de la nouvelle base et Fusion des bases
d0 <- read_csv(here("C:/Users/mmala/OneDrive/Bureau/Saisie AHOUYA/Exported/Exported.csv"))
d0$STADE_CLASSIFICATION <- NULL
dim(d1)
dim(d0)

d <- bind_rows(d1,d0)
dim(d)

d <- d %>% mutate(Q3_2 = as.double(Q3_2), Q3_3 =  as.double(Q3_3))

#Traitement des données

## Labellisation des variables et des modalités
var_lab(d$NUM_FICHE)= "Numero fiche"
var_lab(d$JOUR_CONSUL)= "Jour consultation"
var_lab(d$MOIS_CONSUL)= "Mois consultation"
var_lab(d$ANNEE_CONSUL)= "Annee consultation"
var_lab(d$NOM   ) =  "Nom"
var_lab(d$PRENOM) =  "Prenom"
var_lab(d$AGE   ) =  "Age"
var_lab(d$JOUR  ) =  "Jour de naissance"
var_lab(d$MOIS  ) =  "Mois de naissance"
var_lab(d$ANNEE ) =  "Annee de naissance"
var_lab(d$SEXE  ) =  "Sexe"
var_lab(d$Q1A   ) =  "Situation matrimoniale"
var_lab(d$Q1B   ) =  "Nombre d?enfants"
var_lab(d$Q1C   ) =  "Profession"
var_lab(d$Q1D   ) =  "Ville"
var_lab(d$Q2A   ) =  "Nature1 ou nom de la maladie cancéreuse"
var_lab(d$Q2B   ) =  "Nature2 ou nom de la maladie cancéreuse"
var_lab(d$Q2C   ) =  "Cancer solide"
var_lab(d$Q2D   ) =  "Cancer hématologique"
var_lab(d$Q2E   ) =  "Classification_T"
var_lab(d$Q2F   ) =  "Classification_N"
var_lab(d$Q2G   ) =  "Classification_M"
var_lab(d$Q2GG  ) =  "Stade"
var_lab(d$Q3_1  ) =  "ASAT demané"
var_lab(d$Q3_2  ) =  "Valeur ASAT"
var_lab(d$Q3_3  ) =  "Valeur ALAT"
var_lab(d$Q4_1A ) =  "HBS démandé"
var_lab(d$Q4_1B ) =  "Résultat HBS"
var_lab(d$Q4_2A ) =  "AC antiHBC IgG demandé"
var_lab(d$Q4_2B ) =  "Résultat AC antiHBC IgG"
var_lab(d$Q4_2C ) =  "Charge virale du VHB demandée"
var_lab(d$Q4_3A ) =  "Ac antiHBC IgM demandé"
var_lab(d$Q4_3B ) =  "Résultat Ac antiHBC IgM"
var_lab(d$Q4_3C ) =  "charge virale du VHB demandée"
var_lab(d$Q4_3D ) =  "Résultat charge virale VHB"
var_lab(d$Q4_4A ) =  "Ac antiHBS demandée"
var_lab(d$Q4_4B ) =  "Résultat Ac antiHBS"
var_lab(d$Q5_1A ) =  "ADN du VHB demandé après début de la chimio"
var_lab(d$Q5_1B ) =  "Résultat du ADN du VHB"
var_lab(d$Q5_1C ) =  "Traitement antiVHB  débuté"
var_lab(d$Q5_1D ) =  "Fréquence de la réalisation de l'ADN du VHB"
var_lab(d$Q5_1DD) =  "Autre fréquence de la réalisation de l'ADN du VHB"
var_lab(d$Q5_2  ) =  "AgHBS demandé après debut du traitement"
var_lab(d$Q5_3A ) =  "ADN du VHB demandé après initiation de la chimiothérapie"
var_lab(d$Q5_3B ) =  "Résultat ADN du VHB"
var_lab(d$Q5_3C ) =  "Traitement antiVHB débuté"
var_lab(d$Q5_3D ) =  "Valeur de la charge virale du VHB"
var_lab(d$Q5_3E ) =  "Valeur ALAT"
var_lab(d$Q5_3F ) =  "Valeur Bilirubine totale"
var_lab(d$Q5_3G ) =  "Fréquence de la réalisation de l?ADN du VHB"
var_lab(d$Q5_3GG) =  "Autre fréquence de la réalisation de l'ADN du VHB"
var_lab(d$Q6_1  ) =  "Protocol de chimiothérapie reçue1"
var_lab(d$Q6_1A ) =  "Protocol de chimiothérapie reçue2"
var_lab(d$Q6_2  ) =  "Corticothérapie"
var_lab(d$Q6_2A ) =  "Nom de la molécule"
var_lab(d$Q6_2B ) =  "Posologie de la molécule"
var_lab(d$Q6_3  ) =  "Présence d'antracycline dans le protocole"
var_lab(d$Q6_4  ) =  "Agent de déplétion des lymphocytes B (immunothérapie)"
var_lab(d$Q7    ) =  "Dévenir du malade"
var_lab(d$NOM_AGENT)= "Nom"
var_lab(d$PRENOM_AGENT)= "Prenom"

# Pour avoir code et nom de la variable : d %>% map(var_lab)

val_lab(d$SEXE) = num_lab("    
 1 MASCULIN
 2 FEMININ")

val_lab(d$Q1A) = num_lab("     
 1 Célibataire
 2 Marié(e)
 3 Veuf(ve)
 4 Union libre
 5 Séparé(e)/divorcé(e)
 9 ND")

val_lab(d$Q2C) = num_lab("     
    1 Oui
    2 Non")

val_lab(d$Q2D) = num_lab("     
    1 Oui
    2 Non")

val_lab(d$Q3_1) = num_lab("   
    1 Oui
    2 Non")

val_lab(d$Q4_1A) = num_lab("   
    1 Oui
    2 Non")

val_lab(d$Q4_1B) = num_lab("   
    1 Négatif
    2 Positif")

val_lab(d$Q4_2A) = num_lab("   
    1 Oui
    2 Non")

val_lab(d$Q4_2B) = num_lab("  
    1 Positif
    2 Négatif")

val_lab(d$Q4_2C) = num_lab("   
    1 Oui
    2 Non")

val_lab(d$Q4_3A) = num_lab("   
    1 Oui
    2 Non")

val_lab(d$Q4_3B) = num_lab("   
    1 Positif
    2 Négatif")

val_lab(d$Q4_3C) = num_lab("   
    1 Oui
    2 Non")

val_lab(d$Q4_3D) = num_lab("   
 1 Positif
 2 Négatif")

val_lab(d$Q4_4A) = num_lab("   
    1 Oui
    2 Non")

val_lab(d$Q4_4B) = num_lab("   
    1 Positif
    2 Négatif")

val_lab(d$Q5_1A) = num_lab("   
    1 Oui
    2 Non")

val_lab(d$Q5_1B) = num_lab("   
    1 Positif
    2 Négatif")

val_lab(d$Q5_1C) = num_lab("   
    1 Oui
    2 Non")

val_lab(d$Q5_1D) = num_lab("   
    1 Tous les mois
    2 A la fin de chaque cure de chimiothérapie
    3 A la fin de toutes les cures de chimiothérapie
    9 Autre à préciser
    4 Jamais")

val_lab(d$Q5_2) = num_lab("   
 1 Oui
 2 Non")

val_lab(d$Q5_3A) = num_lab("   
    1 Oui
    2 Non")

val_lab(d$Q5_3B) = num_lab("   
    1 Positif
    2 Négatif")

val_lab(d$Q5_3C) = num_lab("   
    1 Oui
    2 Non")

val_lab(d$Q5_3G) = num_lab("   
    1 Tous les mois
    2 A la fin de chaque cure de chimiothérapie
    3 A la fin de toutes les cures de chimiothérapie
    9 Autre à préciser
    4 Jamais")

val_lab(d$Q6_2) = num_lab("   
    1 Oui
    2 Non")
                        
val_lab(d$Q6_3) = num_lab("    
    1 Oui
    2 Non")
                      
val_lab(d$Q6_4) = num_lab("    
    1 Oui
    2 Non")
                      
val_lab(d$Q7) = num_lab("      
     1 Vivant
     2 Décédé
     3 Perdu de vue")
           
d$Q1A_rec <- d$Q1A 
d$Q1A_rec[d$Q1A_rec == 9] <- 1
var_lab(d$Q1A_rec)<- "Situation matrimoniale"

d <- d %>% mutate(Q4_1A_rec = case_when(is.na(Q4_1A)& (Q4_1B == 1)~ 1,
                                        .default = unclass(Q4_1A)
))

var_lab(d$Q4_1A_rec)<- "HBS démandé"

val_lab(d$Q4_1A_rec)= num_lab("
                              1 Oui
                              2 Non
                              ")
d$Q5_1A[is.na(d$Q5_1A)] <- 1
d$Q5_3B[d$Q5_3B == 2]<- NA
## Tri à pla des variables
#triapla <- summarytools::dfSummary(d)

## Age
#Calcul de l'âge à partir de la date d'entrée
d$date_consultation <- make_date(d$ANNEE_CONSUL,d$MOIS_CONSUL,d$JOUR_CONSUL)
d$date_naissance <- make_date(d$ANNEE,d$MOIS,d$JOUR)


d$AGE <- unclass(d$AGE)
d$AGE_REVOLU <- as.period(interval(d$date_naissance, d$date_consultation))@year

d <- d %>% mutate(AGE_rec = case_when(
  AGE == 999 ~ AGE_REVOLU,
  .default =  AGE
))

d %>% select (AGE,JOUR_CONSUL,MOIS_CONSUL,ANNEE_CONSUL,date_consultation,JOUR,MOIS,ANNEE,date_naissance,AGE_rec)%>% filter(d$AGE == 999) %>% view()


d$gpage <- recode(d$AGE_rec,18:29~1,30:39~2,40:49~3,50:59~4,60:69~5,70:110~6,TRUE ~ copy)
var_lab(d$gpage) <- "Groupe d'âge"
val_lab(d$gpage)<- num_lab("
                              1 18-29 ans
                              2 30-39 ans
                              3 40-49 ans
                              4 50-59 ans
                              5 60-69 ans
                              6 70 ans+
                              ")

fre(d$gpage)
##Sexe
d$genre <- d$SEXE
var_lab(d$genre)<- "Sexe"

##Situation matrimonial
fre(d$Q1A)
d$stmat <- to_factor(d$Q1A,"l")

##Nombre d'enfants
d$nbre_enfant <- recode(d$Q1B,7:20~95,TRUE~copy)
var_lab(d$nbre_enfant)= "Nombre d'enfants"
val_lab(d$nbre_enfant) = num_lab("
                                 95 7+
                                 ")

## Profession
#�
d$Q1C <- str_replace(d$Q1C,"�","é") 
d$Q1C <- toupper(d$Q1C)
fre(d$Q1C)

#Tprotocole chimiothérapique
# indice des caractères spécifiques

k <- which(str_detect(d$Q6_1 ,"�"))
for (i in k){
  d$Q6_1[i] <- str_replace(d$Q6_1[i],"�","2")
}

# d$Q6_1_rec <- str_replace(d$Q6_1 ,"M�","M2") 
# d$Q6_1_rec <- str_replace(d$Q6_1 ,"m�","M�") 
# d$Q6_1_rec <- str_replace(d$Q6_1 ,"�","à") 
# which(str_detect(d$Q6_2A,"�"))#64
d$Q6_2A[64] <- str_replace(d$Q6_2A[64],"�","2")

d$Q6_2A_rec <- as.character(d$Q6_2A)
str_replace(d$Q6_2A_rec[64],"�","M2")



##Ville
d <- d %>% mutate(ville_rec = case_when(
  is.na(Q1D)~ "Brazzaville",
  str_detect(Q1D,"Bra|BRA|bra")~ "Brazzaville",
  str_detect(Q1D,"Owan|OWAN|owan")~ "Owando",
  str_detect(Q1D,"POINTE|Pointe|pointe")~ "Pointe-Noire",
  str_detect(Q1D,"KINKA|Kinka|kina")~ "Kinkala",
  .default = "Autre ville"
))
var_lab(d$ville_rec)<- "Ville de provenance"
## Maladie cancéreuse
# fre(d$Q2A)
d <- d %>% mutate(Nature_cancer = case_when(
                          is.na(Q2A)~ "Cancer du sein",# NA,
                          str_detect(Q2A,"REIN")~ "Cancer du rein",
                          str_detect(Q2A,"Colon|colon|COLON|DU COLON")~ "Cancer du Colon",
                          str_detect(Q2A,"UTERIN|COL")~ "Cancer du col de l'utérus",
                          str_detect(Q2A,"PROSTATE")~ "Cancer de la prostate",
                          str_detect(Q2A,"CUTANE|PIED|PEAU")~ "Cancer de la peau",
                          str_detect(Q2A,"POUMON")~ "Cancer du poumon",
                          str_detect(Q2A,"l'OS|l'os|FEMUR")~ "Cancer de l'os",
                          str_detect(Q2A,"ESTOMAC")~ "Cancer d'estomac",
                          str_detect(Q2A,"RECTUM")~ "Cancer du rectum",
                          str_detect(Q2A,"SEIN")~ "Cancer du sein",
                          str_detect(Q2A,"OVAIRE")~ "Cancer de l'ovaire",
                          str_detect(Q2A,"VESSIE")~ "Cancer de la vessie",
                          str_detect(Q2A,"VESSIE")~ "Cancer de la vessie",
                          .default = as.character(Q2A)
)) 
var_lab(d$Nature_cancer)<- "Nature du cancer"
fre(d$Nature_cancer)

##Type de tumeur
d<- d %>% mutate(Type_tumeur = case_when(
  #is.na(d$Q2C)& is.na(d$Q2D)~ NA,
  (d$Q2C == 1)& (d$Q2D == 1)~ "Tumeurs mixtes",
  (d$Q2C == 1)& (d$Q2D != 1)~ "Tumeurs solides",
  (d$Q2C != 1)& (d$Q2D == 1)~ "Tumeurs hématologiques",
    .default = "Autre"
))

d$Type_tumeur[d$Type_tumeur == "Autre"] <- "Tumeurs solides"

var_lab(d$Type_tumeur)<- "Type de la tumeur"
fre(d$Type_tumeur)


##Stade
d$stade <- to_factor(d$Q2GG,explicit_tagged_na = TRUE,sort_levels ="l")
  d <- d %>% mutate(stade = replace_na(stade, "4"))
  fre(d$stade)

  d$stade_rec <- d$stade %>%
    fct_recode(
      "I" = "1",
      "II" = "2",
      "II" = "2A",
      "II" = "IIA",
      "II" = "IIB",
      "III" = "3",
      "III" = "IIIA",
      "III" = "IIIB",
      "III" = "3B",
      "IV" = "4",
      "IV" = "IVA",
      "IV" = "0",
      "IV" = "B",
      "IV" = "LOCALEMENT AVAN",
      "IV" = "4B"
    )
  var_lab(d$stade_rec)="Stade du cancer"
  fre(d$stade_rec)
  d$stade_rec <- fct_relevel(d$stade_rec, sort)#ordonner les niveaux de facteurs
    


# Valeurs des transaminases (ALAT) avant début de la chimiothérapie

  d$SEXE <- to_factor(d$SEXE,"l")
   Moyen_asat_alat <- d %>% select(c(Q3_2,Q3_3)) %>% summarise(Moyen_ASAT= mean(Q3_2,na.rm = TRUE),Moyen_ALAT = mean(Q3_3,na.rm = TRUE))
   Moyen_asat_alat_genre <- d %>% select(c(Q3_2,Q3_3)) %>% group_by(d$SEXE) %>% summarise(Moyen_ASAT= mean(Q3_2,na.rm = TRUE),Moyen_ALAT = mean(Q3_3,na.rm = TRUE))
   d %>% select(Q3_2,Q3_3) %>% summarytools::dfSummary()
# Valeur ALAT
d$valeur_ALAT <- recode(d$Q3_3,lo %thru% 4.99 ~ 1, 5 %thru% 14.99 ~ 2,15 %thru% hi ~ 3,TRUE ~ 0 )
var_lab(d$valeur_ALAT) <- "Valeur ALAT avant début de la chimiothérapie"
val_lab(d$valeur_ALAT) = num_lab(
  "
  1 <5 ULN
  2 5 ULN-15 ULN
  3 15 ULN ou plus
  0 ND
  "
)

fre(d$valeur_ALAT)

d$Q4_1A[is.na(d$Q4_1A)]<- 1
d$Q4_3A[is.na(d$Q4_3A)]<- 1
d$Q4_4A[is.na(d$Q4_4A)]<- 1

#Enregistrement de la base
export(d, "Cancer.rds", format = "rds")

# Statistiques

d <-readRDS("Cancer.rds")

theme_gtsummary_language("fr", decimal.mark = ",", big.mark = " ")

m <- d %>% select(-c(NUM_FICHE,NUM_FICHE,JOUR_CONSUL,MOIS_CONSUL,NOM,PRENOM,JOUR,MOIS,PRENOM_AGENT,
                date_consultation,Q1A,Q4_1A,ANNEE,Q2B,Q2E,Q2F,Q2G,Q2GG,AGE_REVOLU,date_naissance,
                NOM_AGENT,Q4_3D,Q5_1C,Q5_1DD,Q5_3B,Q5_3C,Q5_3D,Q5_3E,Q5_3F,Q5_3G,
                Q5_3GG,Q4_3C,
                Q5_1D,
                Q5_1B,
                Q4_2C,
                Q4_3B,
                Q4_4B,
                Q4_2B,
                Q6_1A,
                Q6_2B))

T1 <-tbl_summary(d,include = c(ANNEE_CONSUL,gpage,genre,Q1A_rec,nbre_enfant,ville_rec,
                          Nature_cancer,Type_tumeur,stade_rec,valeur_ALAT,Q1A_rec,Q4_1A_rec,Q4_1B,Q4_2A,Q4_2B,Q5_1A,Q5_1B,Q5_1C,Q5_1D,Q5_2,
                          Q5_3A,Q5_3B,Q5_3C,Q5_3D,Q5_3E,Q5_3F,Q5_3G
),
            missing = "ifany",
            missing_text = "ND"
)%>% bold_labels()

as_hux_xlsx(T1, "Vue des données.xlsx", include = everything(), bold_header_rows = TRUE)

#Caractéristiques démographiques
## age moyen

Age_moyen_sex <- d %>% tbl_summary(
  include = c("AGE_rec"),
  statistic = list(all_continuous() ~ "{mean} ({sd})"),
  digits = list(AGE_rec ~ 1),
  label = list(AGE_rec ~ "Age moyen"),
  by = "SEXE"
)

Moyenne_age <- d %>% tbl_summary(
  include = c("AGE_rec"),
  statistic = list(all_continuous() ~ "{mean} ({sd})"),
  digits = list(AGE_rec ~ 1),
  label = list(AGE_rec ~ "Age moyen"))

Age_moyen <- tbl_merge(tbls=list(Age_moyen_sex,Moyenne_age), tab_spanner = c("**Sexe**", "**Ensemble**"))
show_header_names(Age_moyen)

Age_moyen <-  Age_moyen %>%
  modify_header(
    list(
      label ~ "**Variable**"
    ))%>%
  modify_footnote(everything() ~ "Moyenne (Ecart-type)") %>% modify_caption("**Tableau 1: age moyen des patients selon le sexe**")

AgMoyen_Nature <- d %>% select(c(AGE_rec,Nature_cancer)) %>% group_by(d$Nature_cancer) %>% summarise(Moyen_Age= mean(AGE_rec,na.rm = TRUE))
AgMoyen_Nature_sex <- d%>% select(c(AGE_rec,Nature_cancer)) %>% group_by(d$Nature_cancer,d$SEXE) %>% summarise(Moyen_Age= mean(AGE_rec,na.rm = TRUE))
Moyen_Age= mean(d$AGE_rec,na.rm = TRUE)
export(list(TOT_Ag_moy = AgMoyen_Nature, AgMoyen_Nat_sex = AgMoyen_Nature_sex,Ens_ag_moy=Moyen_Age),
       here("Résultats","Age_moyen_tumeur.xlsx"))

#Répartition par groupe d'âges et sexe
groupage_sexe <- tbl_cross(d,
                        row = gpage,
                        col = SEXE,
                        percent = c( "column"),
                        margin = c("column", "row"),
                        missing = c( "ifany"),
                        #missing_text = "Manquant",
                        margin_text = "Total")%>%  bold_labels()

# Création de la variable sexe ration par groupe d'âge
# d$Male <- count_if("MASCULIN", d$SEXE)
# unique(d$Male)
# d$Female <- count_if("FEMININ", d$SEXE)
# unique(d$Female)

Male <-  d %>% filter(SEXE == "MASCULIN") %>% group_by(gpage) %>% count()
Female <- d %>% filter(SEXE == "FEMININ") %>% group_by(gpage) %>% count()
SEXE_ratio <- round(Male*100/Female,1)


SEXE_ratio$gpage[1] <- "18-29 ans"
SEXE_ratio$gpage[2] <- "30-39 ans"
SEXE_ratio$gpage[3] <- "40-49 ans"
SEXE_ratio$gpage[4] <- "50-59 ans"
SEXE_ratio$gpage[5] <- "60-69 ans"
SEXE_ratio$gpage[6] <- "70 ans+"
names(SEXE_ratio)[2] <- "Sexe ratio"

#Type de tumeur
d$Type_tumeur <- fct_relevel(d$Type_tumeur, "Tumeurs solides", "Tumeurs hématologiques")
tumeur <- d %>% tbl_summary(
  include = c(Type_tumeur), by = SEXE,
  digits = list(Type_tumeur ~ 1)) %>% add_overall(last=TRUE)
  #label = list(AGE_rec ~ "Age moyen"))

t1 <- d %>% tbl_summary(
  include = c(Type_tumeur),statistic = list(all_categorical() ~ "{n}"
  ))%>%  bold_labels()

t2 <- d %>% tbl_summary(
  include = c(Type_tumeur),statistic = list(all_categorical() ~ "{p}"),
  digits = list(Type_tumeur ~ 1)
  )%>%  bold_labels()

tumeur <- tbl_merge(tbls=list(t1, t2), tab_spanner = c("**Effectif**", "**%**"))

d %>% tbl_cross(row = Type_tumeur, col = SEXE, percent = "cell", 
                margin_text = "total",missing_text = "Unknown",digits = 1) %>%
  add_p() %>%
  bold_labels()

#Nature histolique

n1 <- d %>% tbl_summary(
  include = c(Nature_cancer),statistic = list(all_categorical() ~ "{n}"
  ),
  sort = list(everything() ~ "frequency")
  )%>%  bold_labels()

n2 <- d %>% tbl_summary(
  include = c(Nature_cancer),statistic = list(all_categorical() ~ "{p}"),
  digits = list(Nature_cancer ~ 3),
  sort = list(everything() ~ "frequency")
)%>%  bold_labels()

nature_tumeur <- tbl_merge(tbls=list(n1, n2), tab_spanner = c("**Effectif**", "**%**"))

Nature_histologique <- tbl_merge(tbls=list(n1, n2), tab_spanner = c("**Effectif**", "**%**"))
as_hux_xlsx(tumeur,here("Résultats","Nature histologique.xlsx"), include = everything(), bold_header_rows = TRUE)
as_hux_xlsx(Nature_histologique,here("Résultats","Nature histologique.xlsx"), include = everything(), bold_header_rows = TRUE)

export(SEXE_ratio,format = "xlsx")
# Répartition selon le type de chimiothérapie
type_chimiothérapie <- fre(d$Q6_1_rec)
export(type_chimiothérapie,here("Résultats","type chimiothérapie.xlsx"))

# export(list(a = SEXE_ratio, b = iris), "SEXE_ratio.xlsx") #exportation en plusieurs feuilles Excel
# dir.exists("Résultats")
# dir.create("Résultats")
# list.files()
#list.files(pattern = ".xlsx")#Pour avoir la liste de tous les fichiers Excel


#Exportation des résultats
#as_flex_table(Age_moyen)
export(list(Age_moyen = as_tibble(Age_moyen),RM = SEXE_ratio,Iris = iris,
            groupage_sexe = as.data.frame(groupage_sexe)),
       here("Résultats","Caractéristiques_patients.xlsx"))

as_hux_xlsx(groupage_sexe, "groupage vs sexe.xlsx", include = everything(), bold_header_rows = TRUE)
as_hux_xlsx(Age_moyen,  here("Résultats","Age_moyen.xlsx"), include = everything(), bold_header_rows = TRUE)



table_janitor <- d %>% tabyl(gpage,SEXE) %>% #Cf. chap17 epirhandbook
  adorn_totals(where = c("row", "col")) %>% 
  adorn_percentages(denominator = "col") %>% 
  adorn_pct_formatting(digits = 2) %>% 
  adorn_ns(position = "front") %>% 
  adorn_title(
    row_name = "Groupe d'âges",
    col_name = "Genre"
  )
export(table_janitor,"janitor_table.xlsx")

#Données du dépistage de l’hépatite B avant le début de la chimiothérapie 

test_AntigèneHBS <- fre(d$Q4_1A)

result_AntigèneHBS <- d %>% filter(d$Q4_1A == 1) %>% select(Q4_1B) %>% fre()
export(list(test = test_AntigèneHBS,Résultat = result_AntigèneHBS),
       here("Résultats","Dépistage hépatite B.xlsx"))


marqueurs_hépatite_B <- d %>% tbl_summary(include = c(Q4_1A,Q4_2A,Q4_3A,Q4_4A,Q5_1A),
                                          digits = all_categorical()~ c(0,1)) %>% bold_labels()
Result_marqueurs_hépatite_B <- d %>% tbl_summary(include = c(Q4_1B,Q4_2B,Q4_3B,Q4_4B,Q5_1A),
                                          digits = all_categorical()~ c(0,1),
                                          missing = "no")%>% bold_labels()
as_hux_xlsx(marqueurs_hépatite_B,here("Résultats","Marqueurs Test hépatite B.xlsx"),include = everything(), bold_header_rows = TRUE)
as_hux_xlsx(Result_marqueurs_hépatite_B,here("Résultats","Marqueurs Résults hépatite B.xlsx"),include = everything(), bold_header_rows = TRUE)

wb = createWorkbook()
addWorksheet(wb,"Marqueurs")
addWorksheet(wb,"Résultas Marqueurs")
m1 <- as.etable(marqueurs_hépatite_B) 
m2 <- as.etable(Result_marqueurs_hépatite_B)
xl_write(m1,wb, "Résultas Marqueurs") 
xl_write(m2,wb, "Résultas Marqueurs") 
#activeSheet(wb) <- "tab_6" # La feuille tab_6 est la feuille active
saveWorkbook(wb, here("Résultats","Marqueurs hepatite B.xlsx"), overwrite = TRUE)


#Analyse des facteurs pouvant déterminer le dépistage du VHB
groupage_sexe2 <- d %>% filter(Q4_1A == 1) %>% tbl_cross(row = gpage,
                           col = SEXE,
                           percent = c( "column"),
                           margin = c("column", "row"),
                           digits = c(0,3),
                           missing = c( "ifany"),
                           #missing_text = "Manquant",
                           margin_text = "Total")%>%  bold_labels()


l1 <- d %>% filter(Q4_1A == 1) %>% tbl_summary(
  include = c(Type_tumeur),statistic = list(all_categorical() ~ "{n}"
  ))%>%  bold_labels()

l2 <- d %>% filter(Q4_1A == 1) %>%  tbl_summary(
  include = c(Type_tumeur),statistic = list(all_categorical() ~ "{p}"),
  digits = list(Type_tumeur ~ 3)
)%>%  bold_labels()

tumeur2 <- tbl_merge(tbls=list(l1, l2), tab_spanner = c("**Effectif**", "**%**"))

#Rapport de masculinité
Male2 <-  d %>% filter(Q4_1A == 1) %>%  filter(SEXE == "MASCULIN") %>% group_by(gpage) %>% count()
Female2 <- d %>% filter(Q4_1A == 1) %>%  filter(SEXE == "FEMININ") %>% group_by(gpage) %>% count()
SEXE_ratio2 <- round(Male2*100/Female2,1)


SEXE_ratio2$gpage[1] <- "18-29 ans"
SEXE_ratio2$gpage[2] <- "30-39 ans"
SEXE_ratio2$gpage[2] <- "30-39 ans"
SEXE_ratio2$gpage[2] <- "30-39 ans"
SEXE_ratio2$gpage[3] <- "40-49 ans"
SEXE_ratio2$gpage[3] <- "40-49 ans"
SEXE_ratio2$gpage[3] <- "40-49 ans"
SEXE_ratio2$gpage[4] <- "50-59 ans"
SEXE_ratio2$gpage[5] <- "60-69 ans"
SEXE_ratio2$gpage[6] <- "70 ans+"
names(SEXE_ratio2)[2] <- "Sexe ratio"

i1 <- d %>% filter(Q4_1A == 1) %>% tbl_summary(
  include = c(Nature_cancer),statistic = list(all_categorical() ~ "{n}"
  ),
  sort = list(everything() ~ "frequency")
)%>%  bold_labels()

i2 <- d %>% filter(Q4_1A == 1) %>% tbl_summary(
  include = c(Nature_cancer),statistic = list(all_categorical() ~ "{p}"),
  digits = list(Nature_cancer ~ 3),
  sort = list(everything() ~ "frequency")
)%>%  bold_labels()

nature_tumeur2 <- tbl_merge(tbls=list(i1, i2), tab_spanner = c("**Effectif**", "**%**"))

#Valeur ALAT
d$valeur_ALAT[d$valeur_ALAT==0]<- 3
alat2 <- d %>% filter(Q4_1A == 1) %>% select(valeur_ALAT) %>%  fre()
d %>% filter(Q4_1A == 1) %>% select(valeur_ALAT) %>%  tbl_summary(digits =list(valeur_ALAT~ c(0,3)))


#Caractéristiques des patients n’ayant pas fait de dépistage du VHB
groupage_sexe3 <- d %>% filter(Q4_1A == 2) %>% tbl_cross(row = gpage,
                                                         col = SEXE,
                                                         percent = c( "column"),
                                                         margin = c("column", "row"),
                                                         digits = c(0,3),
                                                         missing = c( "ifany"),
                                                         #missing_text = "Manquant",
                                                         margin_text = "Total")%>%  bold_labels()

L1 <- d %>% filter(Q4_1A == 2) %>% tbl_summary(
  include = c(Type_tumeur),statistic = list(all_categorical() ~ "{n}"
  ))%>%  bold_labels()

L2 <- d %>% filter(Q4_1A == 2) %>%  tbl_summary(
  include = c(Type_tumeur),statistic = list(all_categorical() ~ "{p}"),
  digits = list(Type_tumeur ~ 3)
)%>%  bold_labels()

tumeur3 <- tbl_merge(tbls=list(L1, L2), tab_spanner = c("**Effectif**", "**%**"))


N1 <- d %>% filter(Q4_1A == 2) %>% tbl_summary(
  include = c(Nature_cancer),statistic = list(all_categorical() ~ "{n}"
  ),
  sort = list(everything() ~ "frequency")
)%>%  bold_labels()

N2 <- d %>% filter(Q4_1A == 2) %>% tbl_summary(
  include = c(Nature_cancer),statistic = list(all_categorical() ~ "{p}"),
  digits = list(Nature_cancer ~ 3),
  sort = list(everything() ~ "frequency")
)%>%  bold_labels()

nature_tumeur3 <- tbl_merge(tbls=list(N1, N2), tab_spanner = c("**Effectif**", "**%**"))

d %>% tbl_summary(
  include = c(Q3_2,Q3_3),
  statistic = list(all_continuous() ~ "{mean} ({sd})"),
  digits = list(all_continuous() ~ 1),
  label = list(Q3_2 ~ "Valeur moyenne ASAT",Q3_3~ "Valeur moyenne ALAT"))


Moyen_asat_alat3 <- d %>% filter(Q4_1A == 2) %>%  select(c(Q3_2,Q3_3)) %>% summarise(Moyen_ASAT= mean(Q3_2,na.rm = TRUE),
                                                          SD_ASAT = sd(Q3_2,na.rm = TRUE),
                                                            Moyen_ALAT = mean(Q3_3,na.rm = TRUE),
                                                            SD_ALAT = sd(Q3_3,na.rm = TRUE),
                                                          n = n())
export(Moyen_asat_alat3,here("Résultats","Moyen_asat_alat3.xlsx"))


#Valeur ALAT
#alat3 <- d %>% filter(Q4_1A == 2) %>% select(valeur_ALAT) %>%  fre()
alat3 <- d %>% filter(Q4_1A == 2) %>% select(valeur_ALAT) %>%  tbl_summary(digits =list(valeur_ALAT~ c(0,3)))




#Comparaison entre les patients ayant fait le dépistage et ceux ne l’ayant pas fait (HBS ,HBC , HBS et HBC )
look_for(d,"Q4_")
test_VHB <- d %>% tbl_summary(include = c(Q4_1A,Q4_2A, Q4_3A,Q4_4A),
                  digits = all_categorical()~ c(0,3))

comp_hbs_age <- tbl_cross(d,
          row = gpage,
          col = Q4_1A,
          percent = c( "column"),
          margin = c("column", "row"),
          missing = c("ifany"),
          digits = c(0,3),
          #missing_text = "Manquant",
          margin_text = "Total")%>% add_p() %>%   bold_labels()

comp_hbs_sex <- tbl_cross(d,
                      row = SEXE,
                      col = Q4_1A,
                      percent = c( "column"),
                      margin = c("column", "row"),
                      missing = c("ifany"),
                      digits = c(0,3),
                      #missing_text = "Manquant",
                      margin_text = "Total")%>% add_p() %>%   bold_labels()

d %>% tbl_summary(
  include = c(AGE_rec),
  statistic = list(all_continuous() ~ "{mean} ({sd})"),
  digits = list(AGE_rec ~ 1),
  label = list(AGE_rec ~ "Age moyen"),
  by = "Q4_1A"
) %>% add_p()


comp_AC_antiHBC_IgG <- tbl_cross(d,
                      row = gpage,
                      col = Q4_2A,
                      percent = c( "column"),
                      margin = c("column", "row"),
                      missing = c( "ifany"),
                      digits = c(0,3),
                      #missing_text = "Manquant",
                      margin_text = "Total")%>% add_p() %>%   bold_labels()

tbl_cross(d,
          row = Nature_cancer,
          col = Q4_1A,
          percent = c( "column"),
          margin = c("column", "row"),
          missing = c("ifany"),
          digits = c(0,1),
          #missing_text = "Manquant",
          margin_text = "Total")%>% add_p() %>%   bold_labels()

tbl_cross(d,
          row = valeur_ALAT,
          col = Q4_1A,
          percent = c( "column"),
          margin = c("column", "row"),
          missing = c("ifany"),
          digits = c(0,3),
          #missing_text = "Manquant",
          margin_text = "Total")%>% add_p() %>%   bold_labels()

d %>% tbl_summary(
  include = c(Q3_3),
  statistic = list(all_continuous() ~ "{mean} ({sd})"),
  digits = list(Q3_3 ~ 1),
  label = list(Q3_3 ~ "ALAT"),
  by = "Q4_1A"
) %>% add_p(Q3_3 ~ "t.test")#spécifié le test


d %>% tbl_summary(
  include = c(Q3_3,valeur_ALAT),
  statistic = list(all_continuous() ~ "{mean} ({sd})"),
  digits = list(all_continuous()~ 1, all_categorical()~c(0,3)),
  missing = "no",
  label = list(Q3_3 ~ "Valeur moyenne ALAT"),
  by = "Q4_1A"
) %>% add_overall(last = TRUE) %>% add_p(Q3_3 ~ "t.test")%>%   bold_labels()

d <- d %>% mutate(hbs_hbc = case_when(
  is.na(Q4_1A)&is.na(Q4_2A)~ NA,
  (Q4_1A==1)&(Q4_2A==2)~ 1,#AgHBS seul
  (Q4_1A==1)&(Q4_2A==1)~ 2,#AgHBs seulAgHBS + Ac antiHBc
  (Q4_1A==2)&(Q4_2A==1)~ 3,#Ac antiHBc
  .default = NA
))

d %>% select(Q4_1A,Q4_2A,hbs_hbc) %>% view()

var_lab(d$hbs_hbc)<- "Antigène HBS et HBC"
val_lab(d$hbs_hbc) = num_lab(
  "
  1 AgHBs seul 
  2 AgHBs + Ac antiHBc
  3 Ac antiHBc
  "
  )

stat_HBS_HBC <- d %>% tbl_summary(
  include = c(gpage,SEXE,Type_tumeur,Nature_cancer,valeur_ALAT,hbs_hbc),
  statistic = list(all_continuous() ~ "{mean} ({sd})"),
  digits = list(all_continuous()~ 1, all_categorical()~c(0,3)),
  missing = "no",
  by = "hbs_hbc"
) %>% add_overall(last = TRUE) %>% add_p() %>%  bold_labels()


d <- d %>% mutate(Q6_2A_rec = case_when(
  is.na(Q6_2A)~ NA,
  str_detect(Q6_2A,"SOLUMEDROL 40|SOLUMODROL 40 |SOLUMEDROL40|SOLUMEDRA 40|SOLUMEDUL|SOLUMEDROL 40|solumedrol 40mg")~ "SOLUMEDROL 40mg",
  str_detect(Q6_2A,"SOLUMEDROL 80|SOLUMEDROL 80|SOLUMEDRA 80")~ "SOLUMEDROL 80 mg",
  str_detect(Q6_2A,"SOLUMEDROLE|SOLUMEDROL|solumedrol 40")~ "SOLUMEDROL 40mg",
  # str_detect(Q6_2A,"POINTE|Pointe|pointe")~ "Pointe-Noire",
  # str_detect(Q6_2A,"KINKA|Kinka|kina")~ "Kinkala",
  .default = as.character(Q6_2A)
))

d %>% tbl_summary(include = Q6_2A_rec,digits = list(all_continuous()~ 1, all_categorical()~c(0,3)))

export(d, "Cancer.rds", format = "rds")
## Remplacement de la valeur vide (" ") pas NA
#d$stade2 <- gsub("^$|^ $", NA, d$stade)
## Remplacement de la valeur vide (" ") pas 4
#d$Q2GG <- gsub("^$|^ $", 4, d$stade)
#d$stade <- gsub("^$|^ $", 4, d$stade)

#d %>% filter(Q2E == 4,Q2F == 2,Q2G == 1) %>% select(NUM_FICHE,Q2E,Q2F,Q2G,Q2GG,stade) %>% view()

#Création des variables labelisées
# df <- data_frame(s1 = c("M", "M", "F"), s2 = c(1, 1, 2)) %>%
#   set_variable_labels(s1 = "Sex", s2 = "Question") %>%
#   set_value_labels(s1 = c(Male = "M", Female = "F"), s2 = c(Yes = 1, No = 2))


# df <- foreign:: to_labelled(read.spss(
#   "Cancer.sav",
#     to.data.frame = FALSE,
#     use.value.labels = FALSE,
#     use.missings = FALSE))


#d %>% select(Q2A,Q2B) %>%  view()
