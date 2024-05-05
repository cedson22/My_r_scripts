
rm(list = ls())

setwd("C:/Users/mmala/Documents/REIPER 2023")

library(here)
library(pacman)

p_load(rio,readxl,openxlsx,tidyverse, expss,here,haven,labelled,tables,gtsummary)



d <- read_excel("TOBATELA_BANA Kobo_traitement.xlsx")

names(d)

d <- janitor::clean_names(d)  

glimpse(d)

d <- as_factor(d)
d$situation_enfant <- as.factor(d$quelle_est_la_situation_de_lenfant_3)
d$qualite_du_climat_familial <- d$qualite_du_climat_familial_de_la_relation_avec_les_parents_les_beaux_parents_etc_selon_lenfant_existe_t_il_des_problemes 
d$qualite_du_climat_familial_de_la_relation_avec_les_parents_les_beaux_parents_etc_selon_lenfant_existe_t_il_des_problemes <- NULL
d <- mutate(d,situation_enfant = case_when(situation_enfant == "Enfant en situation de rue" ~ "Situation de rue",
                                      .default = "Milieu carcéral"),
            nom_prenom_bp = nom_prenom_pere_47,
            nationalite_bp = nationalite_pere_48,
            nom_prenom_pere = nom_prenom_pere_51,
            nationalite_pere_ = nationalite_pere_52,
            occupation_pere = occupation_pere_54,
                        )

d <- d %>% select(-c(quelle_est_la_situation_de_lenfant_3 ,                 
                  quelle_est_la_situation_de_lenfant_fiche_enfant_en_situation_de_rue ,                                                    
                  quelle_est_la_situation_de_lenfant_fiche_enfant_en_milieu_carceral,                                              
                  quelle_est_la_situation_de_lenfant_fiche_decoute,photo,                                                                                                                   
                  photo_url,surnom,lieu_de_naissance_15,date_de_naissance,
                  situation_actuelle,                                                                                                      
                  situation_actuelle_a_la_rue,                                                                                             
                  situation_actuelle_en_danger,                                                                                            
                  situation_actuelle_victime_de_traite,                                                                                    
                  situation_actuelle_vagabondage,                                                                                          
                  situation_actuelle_en_famille,                                                                                           
                  situation_actuelle_structure_reiper,situation_matrimoniale_des_parents,
                  situation_matrimoniale_des_parents_ensemble,                                                                             
                  situation_matrimoniale_des_parents_divorces,                                                                             
                  situation_matrimoniale_des_parents_separes,                                                                              
                  situation_matrimoniale_des_parents_pere_decede,                                                                          
                  situation_matrimoniale_des_parents_mere_decedee,                                                                         
                  situation_matrimoniale_des_parents_ne_sait_pas,
                  nationalite_40,occupation,numero_de_la_mere,adresse_mere,                                                                                                            
                  adresse,adresse_pere,numero_du_pere,adresse_tuteur,
                  numero_du_tuteur,lenfant_souhait_il_un_retour_en_famille_oui,
                  lenfant_souhait_il_un_retour_en_famille_non,
                  raisons_et_circonstances_de_la_rupture_familiale_separation_des_parents,                                                 
                  raisons_et_circonstances_de_la_rupture_familiale_desir_dindependance_daventure,                                          
                  raisons_et_circonstances_de_la_rupture_familiale_accusation_de_sorcellerie,                                              
                  raisons_et_circonstances_de_la_rupture_familiale_influence_enthousiasme_des_amis,                                        
                  raisons_et_circonstances_de_la_rupture_familiale_maltraitance_physique_ou_psychologique,                                 
                  raisons_et_circonstances_de_la_rupture_familiale_probleme_cause_par_lenfant_alcool_violence_drogue,                      
                  raisons_et_circonstances_de_la_rupture_familiale_manque_daffection_et_dattention,                                        
                  raisons_et_circonstances_de_la_rupture_familiale_deces_dun_membre_du_foyer,                                              
                  raisons_et_circonstances_de_la_rupture_familiale_par_accident_perdu_egare,                                               
                  raisons_et_circonstances_de_la_rupture_familiale_abandon_des_parents,                                                    
                  raisons_et_circonstances_de_la_rupture_familiale_conditions_materielles_de_vie_difficiles,                               
                  raisons_et_circonstances_de_la_rupture_familiale_mauvaise_relation_avec_un_membre_du_foyer,                              
                  raisons_et_circonstances_de_la_rupture_familiale_autres,
                  par_qui_structure_du_reiper,par_qui_antenne_mobile,par_qui_enfant_lui_meme,                                                                                                 
                  par_qui_cas,par_qui_police,par_qui_autre,niveau_detude_cp_1,                                                                                                   
                  niveau_detude_cp_2,niveau_detude_ce_1,niveau_detude_ce_2,                                                                                                      
                  niveau_detude_cm_1,niveau_detude_cm_2,niveau_detude_6e_106,                                                                                                    
                  niveau_detude_5e_107, niveau_detude_4e_108,niveau_detude_3e_109,                                                                                                    
                  niveau_detude_2nd_110,niveau_detude_1ere,niveau_detude_ter_112,                                                                                                   
                  niveau_detude_cp1,niveau_detude_cp2,niveau_detude_ce1,                                                                                                       
                  niveau_detude_ce2,niveau_detude_cm1,niveau_detude_cm2,                                                                                                       
                  niveau_detude_6e_119,niveau_detude_5e_120,niveau_detude_4e_121,                                                                                                    
                  niveau_detude_3e_122,niveau_detude_2nd_123,niveau_detude_1er,                                                                                                       
                  niveau_detude_ter_125,nom_de_lecole,adresse_de_lecole,
                  nom_du_leader,orientation_sante_infirmiere,orientation_sante_major,
                  orientation_sante_docteur,orientation_sante_autres_142
                  
                  ))
d <- d %>% select(-qualite_du_climat_familial_de_la_relation_avec_les_parents_les_beaux_parents_etc_selon_lenfant_existe_il_des_problemes)


#Traitement du nom et prénom de l'enfant----

d <- d %>% unite("Nom_prenom2", nom,prenom,remove = FALSE, sep = " ")

d <- d %>% mutate (Nom_prenom = case_when(
  
  !is.na(nom) & !is.na(prenom) ~ Nom_prenom2,
  !is.na(nom) & is.na(prenom)  ~ nom,
  is.na(nom) & !is.na(prenom)  ~ prenom,
  is.na(nom) & is.na(prenom) ~ NA
  
))

d %>% select(names(d)[c(3:5,350)]) %>% view()

d <- d %>% select(-c(Nom_prenom2,nom,prenom))

#Récupération des indices des enfants sans nom ou prénom
k <- which(is.na(d$Nom_prenom))

# Suppression de la base des enfants sans nom ou prénom
dim(d)
d <- d[complete.cases(d$Nom_prenom),]
dim(d)

# Sélection des enfants en situation de rue
 var_lab(d$quelle_est_la_situation_de_lenfant_4)= "Situation de l'enfant"
 d$quelle_est_la_situation_de_lenfant_4 <- factor(d$quelle_est_la_situation_de_lenfant_4)
 

 d <- d %>% mutate(
   quelle_est_la_situation_de_lenfant_4 = case_when(
     quelle_est_la_situation_de_lenfant_4 == "ESR" ~ quelle_est_la_situation_de_lenfant_4,
 .default = "EMC"
    
    ) 
)
 
d1 <- d %>% filter(quelle_est_la_situation_de_lenfant_4 == "ESR")# Base enfant en situation de rue
d2 <- d %>% filter(quelle_est_la_situation_de_lenfant_4 == "EMC")# Base enfant en milieu carcéral

# NR_ESR <- questionr::freq.na(d1)
# NR_EMC <- questionr::freq.na(d2)
# export(NR_ESR,format = "xlsx")
# export(NR_EMC,format = "xlsx")


d1$agpe <- recode(d1$age,1 %thru% 9 ~ 1, 10 %thru% 14~2, 15%thru% hi ~ 3)
val_lab(d1$agpe) = make_labels("
                              1 Moins de 10 ans
                              2 10-14 ans
                              3 15-18 ans
                              "
                             )


var_lab(d1$agpe)= "Age de l'enfant"

d1$agpe <- to_factor(d1$agpe,levels = "l")
d1$agpe <- fct_explicit_na(factor(d1$agpe), na_level = "ND")






# Tabulation de la base enfant en situation du rue

l <- d1 %>% select (c(Nom_prenom,genre	,
                      agpe	,
                   rang_dans_la_fratrie	,
                   nombre_de_frere	,
                   nationalite_18,
                   nombre_de_soeur	,
                   situation_matrimoniale_des_parents1	,
                   nationalite_mere	,
                   est_elle_remariee_46	,
                   nationalite_pere_,#missing	
                    nationalite_pere_52	,
                    le_pere_est_il_remarie	,
                    sous_drogue,
                    #circonstance_de_la_rencontre,
                    attitude_de_lenfant,
                    prise_de_contact,
                    nature_du_site,
                    lenfant_est_reference_dans_une_structure,
                    lenfant_a_un_projet,
                    site,
                    lenfant_habite_aupres_dune_tiers_personne	,
                    lenfant_souhaite_t_il_un_retour_en_famille	,
                    est_ce_que_lenfant_est_en_rupture_familiale_73	,
                    lenfant_avait_t_il_deja_quitte_le_foyer_auparavant_88	,
                    y_a_t_il_eu_des_tentatives_de_reunification_familiale_90,
                    niveau_detude,avec_qui_lenfant_vivait_il
))


# fct_explicit_na(f1, na_level = "Manquant")
# summary(d$age)

tab1 <- tabular((factor(d$quelle_est_la_situation_de_lenfant_4)+(Total = 1))~ ((Total = 1)+(pourcentage =Percent("col"))))
#write.table.tabular(tab1,"clipboard", sep="\t")
write.csv.tabular(tab1, here("Résultats","T1.csv"))



tab2 <- tabular((factor(l$agpe)+(Total = 1))~ ((Effectif = 1)+Format(digits=4)*(pourcentage =Percent("col"))))
# write.table.tabular(tab2,"clipboard", sep="\t")
write.csv.tabular(tab2, here("Résultats","T2_Groupe_age.csv"))


tab3 <- tabular((factor(l$genre)+(Total = 1))~ ((Effectif = 1)+(pourcentage =Percent("col"))))
# write.table.tabular(tab2,"clipboard", sep="\t")
write.csv.tabular(tab3, here("Résultats","T3_Sexe.csv"))

l$nationalite_18 <- fct_explicit_na(factor(l$nationalite_18), na_level = "Manquant")
l <- l %>%  
  mutate( Nationalite_enfant = case_when(
  grepl("Centr",nationalite_18) ~ "RCA",
  grepl("cain",nationalite_18) ~ "RCA",
  grepl("Démo",nationalite_18) ~ "RDC",
  grepl("Manquant",nationalite_18) ~ "ND",
  .default = "Congolaise"
  ))

l$Nationalite_enfant <- fct_relevel(l$Nationalite_enfant, "Congolaise", "RDC","RCA")


tab4 <- tabular((factor(l$Nationalite_enfant)+(Total = 1))~ ((Effectif = 1)+Format(digits =2)*(pourcentage =Percent("col"))))
write.csv.tabular(tab4, here("Résultats","T4_Nationalité_enfant.csv")) 






l$rang_fratrie <- recode(l$rang_dans_la_fratrie,4:20 ~ 4, TRUE ~ copy)
var_lab(l$rang_fratrie)= "Rang dans la fatrie"
val_lab(l$rang_fratrie)= num_lab("
                                1 Rang 1
                                2 Rang 2
                                3 Rang 3
                                4 Rang 4+
                               ")
l$rang_fratrie <- to_factor(l$rang_fratrie, levels = "l")

l$rang_fratrie <- fct_explicit_na(l$rang_fratrie, na_level = "ND")




tab5 <- tabular((factor(l$rang_fratrie)+(Ensemble = 1)) ~ (Factor(genre, "Sexe"))*((n=1) + Percent("row"))*Format(digits=2) +(Total = 1) + Percent("row"), data = l)
write.csv.tabular(tab5, here("Résultats","T5_rang_fratrie.csv")) 

tab6 <- l %>% tab_cells(rang_fratrie) %>% 
  tab_cols(genre,total()) %>%
  tab_total_row_position("below") %>% 
  tab_stat_cases(label = "n",total_statistic = "u_cases") %>% # permet de mettre les modalités en colonnes
  tab_stat_rpct(label = " %",total_statistic = "w_rpct") %>%
  tab_pivot(stat_position = "inside_columns") %>% 
  tab_caption("Tableau : Répartition (%) des enfants par rang selon le sexe")


# mtcars %>%
#   tab_cells(cyl) %>%
#   tab_cols(am,vs,total()) %>% tab_stat_cases(label = "n") %>%
#   tab_stat_cpct(total_row_position = "below", label = "%",
#                 total_label =  "Ensemble",
#                 total_statistic = c("u_cpct")) %>%
#   tab_pivot("inside_columns")



wb = createWorkbook()
for (i in 1:6) {addWorksheet(wb,paste("tab",i,sep = "_"))}


xl_write(tab6,wb, "tab_6", 
         # remove '#' sign from totals 
         col_symbols_to_remove = "#",
         row_symbols_to_remove = "#",
         # format total column as bold
         other_col_labels_formats = list("#" = createStyle(textDecoration = "bold")),
         other_cols_formats = list("#" = createStyle(textDecoration = "bold")))

activeSheet(wb) <- "tab_6" # La feuille tab_6 est la feuille active
saveWorkbook(wb, here("Résultats","Mes tableaux.xlsx"), overwrite = TRUE)





gtsummary::tbl_cross(l,
          row = genre,
          col = rang_fratrie,
          percent = c( "column"),
          margin = c("column", "row"),
          missing = c( "ifany"),
          #missing_text = "Manquant",
          margin_text = "Total")



s <- l %>% select_if(is.numeric) 
names(s)

w <- summarytools::descr(s, stats = c("min","mean","sd","max"), transpose = TRUE,
                    headings = FALSE,display.labels = TRUE)
export(w,format = "xlsx")

summarytools::descr(s, stats = c("min","iqr","mean","med", "sd","max"), transpose = TRUE,headings = FALSE,display.labels = TRUE)
s %>% map(summary)

fre(d$nationalite_18)


fre(d$genre)
fre(d$age)
fre(d$rang_dans_la_fratrie)
fre(d$nombre_de_frere)
fre(d$nombre_de_soeur)
fre(d$situation_matrimoniale_des_parents1)
fre(d$nationalite_mere)
fre(d$est_elle_remariee_46)
fre(d$nationalite_pere_48)#missing
fre(d$nationalite_pere_52)
fre(d$le_pere_est_il_remarie)
fre(d$lenfant_habite_aupres_dune_tiers_personne)
fre(d$lenfant_souhaite_t_il_un_retour_en_famille)
fre(d$est_ce_que_lenfant_est_en_rupture_familiale_73)
fre(d$lenfant_avait_t_il_deja_quitte_le_foyer_auparavant_88)
fre(d$y_a_t_il_eu_des_tentatives_de_reunification_familiale_90)
 


j <- fre(d$niveau_detude)

j2 <- questionr ::freq(d$niveau_detude, cum = TRUE, total = TRUE)
export(j,format = "xlsx")

j2 <- questionr ::freq(d$niveau_detude, cum = TRUE, total = TRUE)

export(j2,format = "xlsx")


# library(tables)
v <- tabular( (Species + 1) ~ (n=1) + Sepal.Length + Sepal.Width, data=iris )
write.table.tabular(v,"clipboard", sep="\t") # Ensuite faire un coller sur excel

tab1 <- tabular((factor(d$niveau_detude)+(Total = 1))~ ((Total = 1)+(pourcentage =Percent("col"))))
tab1 <- tabular((factor(niveau_detude) + (Total = 1))~ ((Total = 1)+Format(digits=2)*(pourcentage =Percent("col"))), data=d )
write.table.tabular(tab1,"clipboard", sep="\t")
write.csv.tabular(tab1, "Education.csv")


export(d1, "ESR.rds", format = "rds")# Base ESR
#export(d1, format = "rds")#Commande plus courte
export(d2, "EMC.rds", format = "rds")# Base ESR


# Nombre de frères

l <- mutate(l,nbre_freres = nombre_de_frere)

l$nbre_freres <- recode(l$nombre_de_frere,4 %thru% hi ~ 4, TRUE ~ copy)
var_lab(l$nbre_freres)= "Nombre de frères"
val_lab(l$nbre_freres)= num_lab("
                                1 1
                                2 2
                                3 3
                                4 4+
                               ")

l$nbre_freres <- to_factor(l$nbre_freres,"l")
l$nbre_freres <- fct_explicit_na(l$nbre_freres, na_level = "ND")



tbl_cross(l,  row = nbre_freres ,
          col = genre,
          percent = c( "column"),
          margin = c("column", "row"),
          missing = c( "ifany"),label = list(nbre_freres ~ "Nombre de frères"),
          missing_text = "ND",
          margin_text = "Total") %>% bold_labels()
  



#Nombre de soeurs
l$nbre_soeurs <- recode(l$nombre_de_soeur,4 %thru% hi ~ 4, TRUE ~ copy)
var_lab(l$nbre_soeurs) = "Nombre de soeurs"
val_lab(l$nbre_soeurs)= num_lab("
                                1 1
                                2 2
                                3 3
                                4 4+
                               ")

l$nbre_soeurs <- to_factor(l$nbre_soeurs,"l")
l$nbre_soeurs <- fct_explicit_na(l$nbre_soeurs, na_level = "ND")

#Situation matrimonial des parents

l$Sit_mat <- l$situation_matrimoniale_des_parents1 %>%
  fct_recode(
    "Séparé(e)/Divorcé(e)" = "Divorcés",
    "Veuf/veuve" = "Divorcés Veuve",
    "Séparé(e)/Divorcé(e)" = "Ensemble Séparé",
    "Veuf/veuve" = "Mère décédée",
    "Séparé(e)/Divorcé(e)" = "Séparé",
    "Ne sait pas" = "Séparé Ne sait pas",
    "Veuf/veuve" = "Séparé Veuve",
    "Veuf/veuve" = "Veuf",
    "Veuf/veuve" = "Veufe",
    "Veuf/veuve" = "Veuve",
    "Veuf/veuve" = "Veuve Mère décédée"
  )
var_lab(l$Sit_mat) = "Situation matrimoniale des parents"
l$Sit_mat <- to_factor(l$Sit_mat,"l")
 
# Remariage de la mère et du père
l$rmariage_mere <-to_factor(l$est_elle_remariee_46)
var_lab(l$rmariage_mere) = "Mère remariée"
l$rmariage_mere <- fct_explicit_na(l$rmariage_mere, na_level = "ND")
l$rmariage_mere <- fct_relevel(l$rmariage_mere,"OUI")

l$rmariage_pere <- to_factor(l$le_pere_est_il_remarie)
var_lab(l$rmariage_pere) = "Père remarié"
l$rmariage_pere <- fct_explicit_na(l$rmariage_pere, na_level = "ND")
l$rmariage_pere <- fct_relevel(l$rmariage_pere,"OUI")


# Habitation de l'enfant avec une tiers personne 

l$tiers_personne <- l$lenfant_habite_aupres_dune_tiers_personne
l$tiers_personne <- to_factor(l$tiers_personne)
var_lab(l$tiers_personne) = "Enfant vivant avec une tiers personne"
l$tiers_personne <- fct_relevel(l$tiers_personne,"OUI")
l$tiers_personne <- fct_explicit_na(l$tiers_personne, na_level = "ND")


#Retour en famille de l'enfant

l$retour_en_famille <- l$lenfant_souhaite_t_il_un_retour_en_famille
l$retour_en_famille <- to_factor(l$retour_en_famille)
l$retour_en_famille <- fct_relevel(l$retour_en_famille,"OUI")


#Rupture familiale
l$rupture_familiale <- l$est_ce_que_lenfant_est_en_rupture_familiale_73
l$rupture_familiale <- to_factor(l$rupture_familiale)
l$rupture_familiale <- fct_relevel(l$rupture_familiale,"OUI")


#Quitter le foyer auparavant
l$quitter_foyer_auparavant <- l$lenfant_avait_t_il_deja_quitte_le_foyer_auparavant_88
l$quitter_foyer_auparavant <- to_factor(l$quitter_foyer_auparavant)
l$quitter_foyer_auparavant <- fct_relevel(l$quitter_foyer_auparavant,"OUI")


#Tentative de réunification
l$Tentative_réunification  <- l$y_a_t_il_eu_des_tentatives_de_reunification_familiale_90
l$Tentative_réunification <- to_factor(l$Tentative_réunification)
l$Tentative_réunification <- fct_relevel(l$Tentative_réunification,"OUI")
l$Tentative_réunification <- fct_explicit_na(l$Tentative_réunification,na_level = "ND")

#Niveau d'étude
l$niveau_etude <- l$niveau_detude %>%
  fct_recode(
    "Lycée" = "1èr",
    "Lycée" = "2nd",
    "Collège" = "3è",
    "Collège" = "3è 1èr",
    "Collège" = "4è",
    "Collège" = "5è",
    "Collège" = "6è",
    "Primaire" = "CE1",
    "Primaire" = "CE2",
    "Primaire" = "CM1",
    "Primaire" = "CM2",
    "Primaire" = "CP1",
    "Primaire" = "CP2",
    "Lycée" = "Ter"
  )

l$niveau_etude <- fct_explicit_na(l$niveau_etude,na_level = "ND")
l$niveau_etude <- fct_relevel(l$niveau_etude,c("Primaire","Collège"))




l <- l%>%
  mutate(Enfent_vivait_avec = case_when(
    is.na(avec_qui_lenfant_vivait_il) ~ NA,
    str_detect(avec_qui_lenfant_vivait_il,"grande mère|Grande mère|tante|tante|oncle|Oncle|oncles ") ~ "En famille",
    str_detect(avec_qui_lenfant_vivait_il,"Grandes mère|grand père|Grand père|GRANDS PÈRE|Grand-père") ~ "En famille",
    str_detect(avec_qui_lenfant_vivait_il,"famille|FAMILLE") ~ "En famille",
    str_detect(avec_qui_lenfant_vivait_il,"couple|parent|parents|Papa et maman|couplé|COUPLE|PAPA et maman|père puis la mère|papa et maman ")~ "Les deux parents",
    str_detect(avec_qui_lenfant_vivait_il,"père|pere|Papa|papa|PAPA|Avec papa|son papa|chez papa") ~ "Son père",
    str_detect(avec_qui_lenfant_vivait_il,"Maman|maman|Mère|mère|mere") ~ "Sa mère",
    .default = "Rue|centre|Orphelinat"
  ))
#l %>% select(avec_qui_lenfant_vivait_il,Enfent_vivait_avec) %>% view()
l$Enfent_vivait_avec <- to_factor(l$Enfent_vivait_avec,"Personne avec qui l'enfant vivait")
l$Enfent_vivait_avec <- fct_relevel(l$Enfent_vivait_avec,c("Les deux parents","Son père","Sa mère","En famille"))
l$Enfent_vivait_avec <- fct_explicit_na(l$Enfent_vivait_avec,na_level = "ND")
var_lab(l$Enfent_vivait_avec) <- "Personne avec laquelle l'enfant vivait"





d1 %>%  tbl_summary(include = c(sous_drogue,
#circonstance_de_la_rencontre,
attitude_de_lenfant,
prise_de_contact,
nature_du_site,
lenfant_est_reference_dans_une_structure,
lenfant_a_un_projet,
site))


# Consommation drogue
l$conso_drogue  <- l$sous_drogue
var_lab(l$conso_drogue) = "Consommation drogue"
l$conso_drogue <- to_factor(l$conso_drogue)
l$conso_drogue <- fct_relevel(l$conso_drogue,"OUI")
l$conso_drogue <- fct_explicit_na(l$conso_drogue,na_level = "ND")


#Attitude de l'enfant
l <- l %>% mutate(attitude_enfant = case_when(
                  grepl("Estres|Stressé|Strés|Téméraire",attitude_de_lenfant) ~ "Agité/Stressé",
                  grepl("stres|Tendu|Stress|Stresse",attitude_de_lenfant) ~ "Agité/Stressé",
                  grepl("Agit|Agete|Agi|agité",attitude_de_lenfant) ~ "Agité/Stressé",
                  grepl("Agres",attitude_de_lenfant) ~ "Agressif",
                  grepl("Brigan|Turbu|Violen",attitude_de_lenfant) ~ "Agressif",
                  grepl("Méchant|Nerveux",attitude_de_lenfant) ~ "Agressif",
                  grepl("Drogu",attitude_de_lenfant) ~ "Agressif",
                  grepl("Calm|Enthou|CALME|CALMÉ",attitude_de_lenfant) ~ "Calme/Normal",
                  grepl("Triste|Treste|Fristre|Quémander",attitude_de_lenfant) ~ "Timide/Triste/Fatigué",
                  grepl("Epuisé|Épuisé",attitude_de_lenfant) ~ "Timide/Triste/Fatigué",
                  grepl("Fatigu|fatigué",attitude_de_lenfant) ~ "Timide/Triste/Fatigué",
                  grepl("Normal|NORMAL|Tranquille|Ras|heureux|Heureuse",attitude_de_lenfant) ~ "Calme/Normal",
                  grepl("Timide|timide|peur|TIMIDE",attitude_de_lenfant) ~ "Timide/Triste/Fatigué",
                  
                  is.na(attitude_de_lenfant) ~ "ND"))
                  #.default = "Autre"))
#l %>% select(attitude_de_lenfant,attitude_enfant) %>% filter(is.na(attitude_enfant)) %>%  view()
var_lab(l$attitude_enfant) = "Attitude de l'enfant"
l$attitude_enfant <- to_factor(l$attitude_enfant)
l$attitude_enfant <- fct_relevel(l$attitude_enfant,c("Calme/Normal","Agité/Stressé","Timide/Triste/Fatigué"))


#Prise de contact
l <- l %>% mutate(prise_contact = case_when(
                 str_detect(prise_de_contact,"^([a,A]mi)|^App|^Par") ~ "Ami(e)s",
                 str_detect(prise_de_contact,"^Enfant")~ "Enfant lui-même",
                 #str_detect(prise_de_contact,"^App")~ "Ami(e)s",
                 #str_detect(prise_de_contact,"^Par")~ "Ami(e)s",
                 is.na(prise_de_contact)~ "ND"))

var_lab(l$prise_contact) = "Prise de contact"
l$prise_contact <- to_factor(l$prise_contact)
l$prise_contact <- fct_relevel(l$prise_contact,c("Enfant lui-même","Ami(e)s"))


# Nature_du_site
l$Nature_site <- l$nature_du_site %>%
  fct_recode(
    "Lieu activité survie" = "Autre",
    "Les deux" = "Dortoir Les deux",
    "Lieu activité survie" = "Lieu de l'activité de survie",
    "Lieu activité survie" = "Lieu de l'activité de survie Autre",
    "Les deux" = "Lieu de l'activité de survie Dortoir",
    "Les deux" = "Lieu de l'activité de survie Les deux"
  ) %>% fct_explicit_na(na_level = "ND")

var_lab(l$Nature_site) = "Nature du site de rencontre"
l$Nature_site <- to_factor(l$Nature_site)
l$Nature_site <- fct_relevel(l$Nature_site,c("Dortoir","Lieu activité survie","Les deux"))


#Référencement dans une structure
l$reference_structure <- l$lenfant_est_reference_dans_une_structure %>% fct_explicit_na(na_level = "ND")
var_lab(l$reference_structure) = "Enfant reférencé dans une structure"
l$reference_structure <- to_factor(l$reference_structure)
l$reference_structure <- fct_relevel(l$reference_structure,c("OUI"))


#Projet de l'enfant
l$Enfant_projet <- l$lenfant_a_un_projet %>% fct_explicit_na(na_level = "ND")
var_lab(l$Enfant_projet) = "L'enfant a un projet"
l$Enfant_projet <- to_factor(l$Enfant_projet)
l$Enfant_projet <- fct_relevel(l$Enfant_projet,c("OUI"))
  

#Site
l <- l %>% mutate(site_rec = case_when(
  is.na(site)~ "ND",
  str_detect(site,"^(A[e,é,É])") ~ "Aéroport",
  str_detect(site,"^(B[i,I][f,F])") ~ "Bifouiti",
  str_detect(site,"^(Cata)|(Kataracte)|(Le rapide)|(Les rapides)|(Rapide)|(RApide)|(RAPIDE)|(RAPIDES)") ~ "Cataractes/Rapides",
  #str_detect(site,"^(Centre [EX,EXERIE,experit,EXu perie,Exupéry])") ~ "Saint Exupéry",
  str_detect(site,"^(Centre sporti.)") ~ "Centre sportif",
  str_detect(site,"^([C,c][e,E][n,N][t,T][r,R][e,E,é] sporti[fs,f,ve])|(Mandarine)") ~ "Centre sportif",
  str_detect(site,"^([C,c][e,E][n,N][t,T][r,R][e,E,é] [V,v][i,I][lle,LLE])") ~ "Centre ville",
  str_detect(site,"^([C,c][h,H][u,U])") ~ "CHU",
  str_detect(site,"^(F[r,R][a,A,o,O,e][n,N][T,t][i,I].)") ~ "Frontière",
  str_detect(site,"^La ([f,F][r,R][a,A,o,O,e][n,N][T,t][i,I].)") ~ "Frontière",
  str_detect(site,"^(Gar[é,e])|(Care centrale)|(La gare)") ~ "Gare centrale",
  str_detect(site,"^([K,k][o,O][M,m].)") ~ "Kombé",
  str_detect(site,"^(Lsb)|(LSB)") ~ "LSB",
  str_detect(site,"^(Mazala)|(Mazaka)|(MAZALA)") ~ "Mazala",
  str_detect(site,"^(P K)|(PK)") ~ "PK",
  str_detect(site,"^(POINT)|(PONT)|(Pont)") ~ "Pont 15 août",
  str_detect(site,"(Sstade)|(Stade)|(STADE)") ~ "STADE MASSAMBA Débat",
  str_detect(site,"(Zanga)|(ZANGA )") ~ "Zanga dia ba ngombe",
  str_detect(site,"(Préfecture)|(La préfecture)") ~ "Préfecture",
  str_detect(site,"^Solo") ~ "Solo Béton",
  str_detect(site,"^(Poto poto)|(POTO poto)|(POTO POTO)|(Potopoto)|(Marche poto poto)") ~ "Poto poto",
  str_detect(site,"^(Moukondo)|(MAZALA)|(Mazala)|(Mazaka)") ~ "Mazala/Moukondo",
  str_detect(site,"^(Mayanga)|(MAYANGA)|(Manganga)") ~ "Mayanga",
  str_detect(site,"^(Hugos)|(HUGOS)|(HUGOR)") ~ "Hugos",
  str_detect(site,"^(Hôtel.)") ~ "Hôtel Mikhaels",
  .default = "Autre site"
  ))
#fre(l$site_rec)
var_lab(l$site_rec) = "Site de rencontre"
l$site_rec <- to_factor(l$site_rec)
l$site_rec <- fct_relevel(l$site_rec,"Autre site",after =23)#Autre site devient l'avant dernière modalité
l$site_rec <- fct_relevel(l$site_rec,"ND",after =Inf)#ND devient la dernière modalité


export(l, "l.rds", format = "rds")# Base ESR

####################Caractéristiques socio démographiques de l'enfant###########################

l <- l %>% mutate(genre = to_factor(genre),
                  niveau_etude = to_factor(niveau_etude)
                
                  )

theme_gtsummary_language("fr", decimal.mark = ",", big.mark = " ")

tbl_summary(l,include = c(genre,agpe,niveau_etude,
                        Nationalite_enfant),
            missing = "ifany",
            missing_text = "ND"
            )


t1 <-  l %>%  tbl_summary(include = c(genre,agpe,niveau_etude,
                               Nationalite_enfant),
                   statistic = list(all_categorical() ~ "{n}"
                   ),label = list(genre ~ "Genre",agpe~ "Age" ,niveau_etude ~ "Niveau d'étude",
                                  Nationalite_enfant ~ "Nationalité")
                        )%>%  bold_labels()

t2 <-  l %>%  tbl_summary(include = c(genre,agpe,niveau_etude,
                                      Nationalite_enfant),
                          digits = list(all_categorical() ~ c(2)),
                          statistic = list(all_categorical() ~ "{p}"
                          ),label = list(genre ~ "Genre",agpe~ "Age" ,niveau_etude ~ "Niveau d'étude",
                                         Nationalite_enfant ~ "Nationalité")
                        
                        ) %>%  bold_labels()


T1 <- tbl_merge(tbls=list(t1, t2), tab_spanner = c("**Effectif**", "**%**"))



T1 <- T1 %>% modify_header(list(label ~ "**Variable**",
                   #all_stat_cols(stat_0 = FALSE) ~ "_{level}_ n={n}({style_percent(p)}%)",
                   stat_0_1 ~ "**Effectif Total**  {N}",
                   stat_0_2 ~ "**%** {style_percent(p)}")) %>%
                modify_footnote(everything() ~ NA) %>%
  modify_spanning_header(all_stat_cols() ~ "**Caractéristiques docio démographiques de l'enfant**")
T1
#Exportation en Excel
as_hux_xlsx(T1, "Caractéristiques_enfants.xlsx", include = everything(), bold_header_rows = TRUE)


####################Caractéristiques familiales de l'enfant#####################

l <- readRDS("C:/Users/mmala/Documents/REIPER 2023/l.rds")

t1 <-  l %>%  tbl_summary(include = c(rang_fratrie,nbre_freres,nbre_soeurs,
                                      Sit_mat,rmariage_mere,rmariage_pere,
                                      tiers_personne,Enfent_vivait_avec),
                          statistic = list(all_categorical() ~ "{n}"
                          )#,label = list(rang_fratrie ~ "Genre",agpe~ "Age" ,niveau_etude ~ "Niveau d'étude",
                                         #Nationalite_enfant ~ "Nationalité")
)%>%  bold_labels()

t2 <-  l %>%  tbl_summary(include = c(rang_fratrie,nbre_freres,nbre_soeurs,
                                      Sit_mat,rmariage_mere,rmariage_pere,
                                      tiers_personne,Enfent_vivait_avec),
                          digits = list(all_categorical() ~ c(2)),
                          statistic = list(all_categorical() ~ "{p}"
                          # ),label = list(genre ~ "Genre",agpe~ "Age" ,niveau_etude ~ "Niveau d'étude",
                          #                Nationalite_enfant ~ "Nationalité")
                          
)) %>%  bold_labels()

T2 <- tbl_merge(tbls=list(t1, t2), tab_spanner = c("**Effectif**", "**%**"))

T2 <- T2%>% modify_header(list(label ~ "**Variable**",
                                #all_stat_cols(stat_0 = FALSE) ~ "_{level}_ n={n}({style_percent(p)}%)",
                                stat_0_1 ~ "**Effectif Total**  {N}",
                                stat_0_2 ~ "**%** {style_percent(p)}")) %>%
  modify_footnote(everything() ~ NA) %>%
  modify_spanning_header(all_stat_cols() ~ "**Caractéristiques familiales**")
T2

####################Regroupement familiale/Histoire familiale #######################


t1 <-  l %>%  tbl_summary(include = c(rupture_familiale,retour_en_famille,
                                      quitter_foyer_auparavant,
                                      Tentative_réunification),
                          label = list(rupture_familiale~"Rupture familiale",
                                       retour_en_famille ~ "Retour en famille",
                                       quitter_foyer_auparavant~"Avoir quitté le foyer auparavant",
                                       Tentative_réunification~"Tentative de réunification"),
                          statistic = list(all_categorical() ~ "{n}"
                          )#,label = list(rang_fratrie ~ "Genre",agpe~ "Age" ,niveau_etude ~ "Niveau d'étude",
                          #Nationalite_enfant ~ "Nationalité")
)%>%  bold_labels()

t2 <-  l %>%  tbl_summary(include = c(rupture_familiale,retour_en_famille,
                                      quitter_foyer_auparavant,
                                      Tentative_réunification),
                          label = list(rupture_familiale~"Rupture familiale",
                                       retour_en_famille ~ "Retour en famille",
                                       quitter_foyer_auparavant~"Avoir quitté le foyer auparavant",
                                       Tentative_réunification~"Tentative de réunification"),
                          digits = list(all_categorical() ~ c(2)),
                          statistic = list(all_categorical() ~ "{p}"
                                           # ),label = list(genre ~ "Genre",agpe~ "Age" ,niveau_etude ~ "Niveau d'étude",
                                           #                Nationalite_enfant ~ "Nationalité")
                                           
                          )) %>%  bold_labels()

T3 <- tbl_merge(tbls=list(t1, t2), tab_spanner = c("**Effectif**", "**%**"))

T3 <- T3 %>% modify_header(list(label ~ "**Variable**",
                               #all_stat_cols(stat_0 = FALSE) ~ "_{level}_ n={n}({style_percent(p)}%)",
                               stat_0_1 ~ "**Effectif Total**  {N}",
                               stat_0_2 ~ "**%** {style_percent(p)}")) %>%
  modify_footnote(everything() ~ NA) %>%
  modify_spanning_header(all_stat_cols() ~ "**Regroupement familiale/Histoire familiale**")

T3


####################Situation de rue #######################


t1 <-  l %>%  tbl_summary(include = c(site_rec,attitude_enfant,Nature_site,
                                      conso_drogue,prise_contact,reference_structure,
                          Enfant_projet),
                          statistic = list(all_categorical() ~ "{n}"
                          ))%>%  bold_labels()

t2 <-  l %>%  tbl_summary(include = c(site_rec,attitude_enfant,Nature_site,
                                      conso_drogue,prise_contact,reference_structure,
                                      Enfant_projet),
                          digits = list(all_categorical() ~ c(2)),
                          statistic = list(all_categorical() ~ "{p}")) %>%  bold_labels()

T4 <- tbl_merge(tbls=list(t1, t2), tab_spanner = c("**Effectif**", "**%**"))

T4 <- T4 %>% modify_header(list(label ~ "**Variable**",
                                #all_stat_cols(stat_0 = FALSE) ~ "_{level}_ n={n}({style_percent(p)}%)",
                                stat_0_1 ~ "**Effectif Total**  {N}",
                                stat_0_2 ~ "**%** {style_percent(p)}")) %>%
  modify_footnote(everything() ~ NA) %>%
  modify_spanning_header(all_stat_cols() ~ "**Situation de rue**")

T4


##########################  MINEURS INCARCERES #################################
m <- read_excel("2.Mineurs_incarcérés.xlsx", 
                sheet = "Tableau1")

view(m)
questionr::freq.na(m)

m <- m %>% mutate(Date_ecrou = as.Date(`Date d’écrou`),
                  Age = as.numeric(Age_))

m$Age <- recode(m$Age,12:14~1,15:17~2, 18:28~3)
val_lab(m$Age) <- num_lab("
                          1 12-14 ans
                          2 15-17 ans
                          3 18 ans+
                          ")

# Lieu de naissance
m <- m %>% mutate(Lieu_naissance = case_when(
  str_detect(`Date & lieu de naissance`,"en RDC")~ "RDC",
  is.na(Lieu_naissance)~ NA,
  .default = Lieu_naissance
  ))
m$Lieu_naissance <- to_factor(m$Lieu_naissance)
m$Lieu_naissance <- fct_explicit_na(m$Lieu_naissance, na_level = "ND")


#Nationalité
m <- m %>% mutate(Nationalité = case_when(
  str_detect(Nationalité,"congolaise|Congolaise")~ "Congolaise",
  str_detect(Nationalité,"Française")~ "Française",
  str_detect(Nationalité,"Angolaise")~ "Angolaise",
  str_detect(Nationalité,"RDC")~ "RDC",
  is.na(Nationalité)~ "ND",
  .default = "ND"
))

m$Nationalité <- to_factor(m$Nationalité)
m$Nationalité <- fct_relevel(m$Nationalité,"ND",after =Inf)


#Sexe
m$Sexe[m$Sexe== "F"] <- "Féminin"
m$Sexe[m$Sexe== "M"] <-"Masculin"
m$Sexe <- to_factor(m$Sexe)
m$Sexe <- fct_relevel(m$Sexe,"Masculin")


#Motif
## Recodage de m$Motif en m$Motif_
m$Motif_ <- m$Motif %>% to_factor() %>% 
  fct_recode(
    "ADM /CBV" = "ADM / CBV",
    "Meurtre/Assassinat" = "Assassinat de sa femme",
    "ADM /CBV" = "Association des Malfaiteurs (ADM)",
    "ADM /CBV" = "CBV",
    "ADM /CBV" = "CBV (Perte d’Œil)",
    "ADM /CBV" = "CBV / ADM",
    "ADM /CBV" = "CBV et Vol",
    "ADM /CBV" = "CBV sur sa sœur",
    "ADM /CBV" = "CBVAEMSID",
    "ADM /CBV" = "CBVAEMSID / ADM",
    "ADM /CBV" = "CBVAMSID",
    "Extorsion" = "Extorsion (2 PC)",
    "Meurtre/Assassinat" = "Meurtre",
    "Meurtre/Assassinat" = "Meurtre sur sa Soeur",
    "Viol" = "Viol sur Mineur",
    "Viol" = "Viol sur Mineure",
    "Vol" = "Vol d’argent",
    "Vol" = "Vol d’argent (50000)",
    "Vol" = "Vol d’argent d’1 Ouest africain",
    "Viol" = "Vol et Viol",
    "Vol" = "Vol Qualifié"
  ) %>%
  fct_explicit_na("ND")

m <- m %>% mutate(Motif_ = case_when(
                  str_detect(Motif_,"Sexuelle|Pédophilie")~"Viol",
                  str_detect(Motif_,"Extorsion|verbales|Receleur|Vagabondage|Vagabondage|Tentative")~"Autre",
                  .default = Motif_
                   ))

m$Motif_ <- fct_relevel(m$Motif_,c("Viol","ADM /CBV","Vol","Meurtre/Assassinat","Autre"))

var_lab(m$Motif_) = "Motif d'incarcération"


################# Caractéristiques des mineurs incarcérés ######################
theme_gtsummary_language("fr", decimal.mark = ",", big.mark = " ")
t1 <-  m %>%  tbl_summary(include = c(Sexe,Age,Nationalité,Motif_),
                          statistic = list(all_categorical() ~ "{n}"
                          ))%>%  bold_labels()

t2 <-  m %>%  tbl_summary(include = c(Sexe,Age,Nationalité,Motif_),
                          digits = list(all_categorical() ~ c(2)),
                          statistic = list(all_categorical() ~ "{p}")) %>%  bold_labels()

T5 <- tbl_merge(tbls=list(t1, t2), tab_spanner = c("**Effectif**", "**%**"))

T5 <- T5 %>% modify_header(list(label ~ "**Variable**",
                                #all_stat_cols(stat_0 = FALSE) ~ "_{level}_ n={n}({style_percent(p)}%)",
                                stat_0_1 ~ "**Effectif Total**  {N}",
                                stat_0_2 ~ "**%** {style_percent(p)}")) %>%
  modify_footnote(everything() ~ NA) %>%
  modify_spanning_header(all_stat_cols() ~ "**Caractéristiques des mineurs incarcérés**")

#tbl_summary(m, include = c(Sexe,Age,Nationalité,Motif_))

T5 


d1 <- d %>% filter(quelle_est_la_situation_de_lenfant_4 == "ESR")# Base enfant en situation de rue
d2 <- d %>% filter(quelle_est_la_situation_de_lenfant_4 == "EMC")# Base enfant en milieu carcéral


# Traitement de la base Mineurs incarsérés
l1 <- d2 %>% select(situation_carcerale,motif_dincarceration,date_decrou,)


#Tabulation avec Expss

#l <-  d1 %>% select(genre,agpe,niveau_detude,rang_fratrie,Nationalite_enfant)











# j3 <- summarytools::freq(d$niveau_detude, cumul = FALSE,report.nas = FALSE,style="grid",
                   # Variable.label = "Niveau d'étude",display.type = FALSE)
# export(j3,format = "xlsx")

# trial %>%
#   tbl_cross(row = stage, col = trt, percent = "cell", margin_text = "total",missing_text = "Unknown") %>%
#   add_p() %>%
#   bold_labels()



# t1 <- tbl_cross(trial,
#           row = stage,
#           col = trt,
#           percent = c( "column"),
#           margin = c("column", "row"),
#           missing = c( "ifany"),
#           missing_text = "Unknown",
#           margin_text = "Total")



# t2 <- tbl_cross(trial,
#                 row = stage,
#                 col = grade,
#                 percent = c( "column"),
#                 margin = c("column", "row"),
#                 missing = c( "ifany"),
#                 missing_text = "Unknown",
#                 margin_text = "Total") %>% bold_labels()


# tbl_merge(tbls=list(t1, t2), tab_spanner = c("**Tumor Response**", "**Time to Death**"))
# tbl_stack(list(t1, t2), group_header = c("Modèle bivarié", "Modèle multivarié"))


#Tableaux complexe ou imbiqués


# trial %>%
#   select(age, grade, stage, trt) %>%
#   mutate(grade = paste("Grade", grade)) %>%
#   tbl_strata(
#     strata = grade,
#     .tbl_fun =
#       ~ .x %>%
#       tbl_summary(by = trt, missing = "no") %>%
#       add_n()
#   )



# trial %>%
#   tbl_summary(
#     include = grade,
#     by = trt,
#     percent = "row",
#     statistic = ~"{p}%",
#     digits = ~1
#   ) %>%
#   add_overall(
#     last = TRUE,
#     statistic = ~"{p}% (n={n})",
#     digits = ~ c(1, 0)
#   )




# tbl <- questionr::freq(d$genre,total = TRUE, valid = FALSE)
# tbl1 <- fre(d$genre)
# export(tbl1,"my_tbl.xlsx",format = ".xlsx")

#questionr::freq.na(d)

#w <- trial %>%  dplyr::select(trt, age, grade) %>%  tbl_summary(by = trt) %>%
 # add_p() %>% add_overall(last  = TRUE) 
#as_hux_xlsx(w, "mo.xlsx", include = everything(), bold_header_rows = TRUE)


# 
# tbl1 <- trial %>%
#   tbl_summary(
#     include = c(age, grade),
#     by = trt
#   ) %>%
#   add_overall() %>%
#   add_p()%>% add_significance_stars()



# tbl2 <- tbl1 %>%
  # modify_header(list(label ~ "**Variable**",
  #    all_stat_cols(stat_0 = FALSE) ~ "_{level}_ n={n}({style_percent(p)}%)",
  #    stat_0 ~ "**TOTAL**  N={N}({style_percent(p)}%)",
  #    p.value ~ "**Test de comparaison**\n(p-valeur)"))%>%
  # modify_footnote(everything() ~ NA) %>%
  # modify_spanning_header(all_stat_cols() ~ "**Traitement**")




# Tableaux avec EXPSS

# mtcars %>%
#   tab_cells(cyl) %>%
#   tab_cols(am,vs,total()) %>% tab_stat_cases(label = "n") %>% 
#   tab_stat_cpct(total_row_position = "below", label = "%",
#                 total_label =  "Ensemble",
#                 total_statistic = c("u_cpct")) %>%
#   tab_pivot("inside_columns")


# wb = createWorkbook()
# for (i in 3:8) {addWorksheet(wb,paste("tab",i,sep = "_"))}
# saveWorkbook(wb, "Mes tableaux.xlsx", overwrite = TRUE)


# Remplacement multiple
v <- c(NA,"F",NA,"F","M","F",NA,"M","F","M","M","F","M","M","F",NA,"F","M")
v
v2 <- str_replace_all(v,c("F"="Féminin","M"="Masculin"))
v2
v2 <- str_replace_na(v,"99")
v2
