#EMIM Variables a nivel entidad federativa y total de manufacturas a precios constantes
#Cifras originales de produccion en miles de pesos, horas en miles de horas y trabajadores en promedio mensual de personas
#Rankings personal ocupado y horas trabajadas
#Nota para la Macro: Herramientas -> Referencias/Microsoft Word 16.0 Object Library

#### INICIO ####
#Limpiar variables
rm(list=ls())

#FECHA
#MODIFICACIÓN DE DATOS AUTOMÁTICA (la que se usa siempre, a menos que el mes en que se corra no coincida con el mes de publicación de INEGI)
#Determinar el mes y el año de la encuesta, bajo el supuesto que se publica dos meses después
#Hacer adecuaciones en caso de que se ejecute el código en una fecha distinta a la publicación de INEGI (ver "MODIFICACIÓN DE DATOS MANUAL" 4 líneas abajo)
mes = as.numeric(substr(seq(Sys.Date(), length = 2, by = "-2 months")[2],6,7)) #mes de la base
año = as.numeric(substr(seq(Sys.Date(), length = 2, by = "-2 months")[2],1,4)) #año de la base

#### CAMBIOS TEMPORALES ####
#A mano
#mes=mes-1

#Fecha para bases de csv
fcsv=año

#Fecha para Excel
fxls=paste(año,sprintf("%02d", mes),sep="_")

#Periodos de carpeta
pericar=paste(año,sprintf("%02d", mes),sep=" ")

#Directorio
#"C:/Users/arturo.carrillo/Documents/EMIM/"
setwd(paste0("C:/Users/arturo.carrillo/Documents/EMIM/",pericar))

#Meses para descripción de textos
#Leer CSV de meses
mespal = c("enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre")

#Números cardinales DEL 1 AL 20 en masculino
num_cardinales = c("primer","segundo","tercer","cuarto","quinto","sexto","séptimo","octavo","noveno","décimo","decimoprimer","decimosegundo","decimotercer","decimocuarto","decimoquinto","decimosexto","decimoséptimo","decimoctavo","decimonoveno","vigésimo")


#PAQUETES
library(dplyr)
library(openxlsx)
library(rjson)
library(inegiR)
library(lubridate)
library(tidyverse)




##### DESCARGAR BASE DE INEGI ####
temp = tempfile() #crear archivo temporal
#Descargar archivo del INEGI:
#download.file("https://www.inegi.org.mx/contenidos/programas/emim/2013/datosabiertos/emim_variables_entidad_csv.zip",temp)
download.file("https://www.inegi.org.mx/contenidos/programas/emim/2018/datosabiertos/emim_variables_entidad_csv.zip",temp)

#Nombre de las bases a buscar en el archivo de INEGI
nombre_base_csv=paste0("tr_variable_total_entidad_mensual_2018_",fcsv,".csv")
#nombre_base_csv2=paste0("tr_variable_subsector_entidad_mensual_2013_",fcsv,".csv")

#Descomprimir archivo y leer base
base_csv = read.csv(unz(temp,paste0("conjunto_de_datos/",nombre_base_csv)),encoding = "UTF-8")
colnames(base_csv)[1]="CODIGO_ACTIVIDAD"
#base_csv2 = read.csv(unz(temp,paste0("conjunto_de_datos/",nombre_base_csv2)),encoding = "UTF-8")
unlink(temp) #borrar archivo  temporal (el descargado de INEGI)
#colnames(base_csv2)[1]="CODIGO_ACTIVIDAD"



#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/
#### GRAF HIS GENERAL ####
#Elementos generales para Gráficas Históricas de Jalisco
#   Personal ocupado total (PO_T5)
#   Total de horas trabajadas (HH_T5)
#   Valor de producción (VP)


#Datos mensuales de Jalisco
dat_men_jal=filter(base_csv,CODIGO_ENTIDAD==14) %>%
  select(ANIO,MES,POT=PO_T5,HOR=HH_T5, VP)

#Cambiar meses de número a letras
dat_men_jal$MES=gsub(10,"OCT",dat_men_jal$MES)
dat_men_jal$MES=gsub(11,"NOV",dat_men_jal$MES)
dat_men_jal$MES=gsub(12,"DIC",dat_men_jal$MES)
dat_men_jal$MES=gsub(1,"ENE",dat_men_jal$MES)
dat_men_jal$MES=gsub(2,"FEB",dat_men_jal$MES)
dat_men_jal$MES=gsub(3,"MAR",dat_men_jal$MES)
dat_men_jal$MES=gsub(4,"ABR",dat_men_jal$MES)
dat_men_jal$MES=gsub(5,"MAY",dat_men_jal$MES)
dat_men_jal$MES=gsub(6,"JUN",dat_men_jal$MES)
dat_men_jal$MES=gsub(7,"JUL",dat_men_jal$MES)
dat_men_jal$MES=gsub(8,"AGO",dat_men_jal$MES)
dat_men_jal$MES=gsub(9,"SEP",dat_men_jal$MES)

#Fechas para graficar
dhm=nrow(dat_men_jal)-12 #datos históricos mensuales a graficar con promedio
#Crear base vacía
fecha_gra=matrix(data="",dhm,2)
fecha_gra=as.data.frame(fecha_gra)
fecha_gra=data.frame(lapply(fecha_gra, as.character), stringsAsFactors=FALSE)
names(fecha_gra)=c("ANIO","MES")

for (i in 1:dhm){
  if(dat_men_jal[i,2]=="ENE"){
    fecha_gra[i,1]=dat_men_jal[i+12,1]
  }
  fecha_gra[i,2]=dat_men_jal[i+12,2]
}


# Funcion para Gráfica Histórica
fgrafhis <- function(var){
  #Datos historicos mensuales de variable
  dat_his_men=var
  dhm=length(dat_his_men)-12
  
  #Promedio de historicos mensuales
  pro_his=matrix(data=0,dhm,1) #Crear base vacía
  pro_his=as.data.frame(pro_his)
  names(pro_his)="PROM"
  
  for (i in 1:(dhm)){
    pro_his[i,1]=mean(dat_his_men[(i+1):(i+12)])
  }
  
  dat_his_graf=dat_his_men[13:length(dat_his_men)] #datos históricos de variable a graficar
  graf_his=cbind.data.frame(fecha_gra, VAL=dat_his_graf,PRO=pro_his)
  return(graf_his)
}




#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/
#### GRAF HIS JAL POT HOR ####
#Gráficas Históricas de Jalisco de Personal ocupado total (PO_T5) y Total de horas trabajadas (HH_T5)

#Gráfica Histórica de Personal Ocupado Total
graf_his_pot = fgrafhis(dat_men_jal$POT)

#Gráfica Histórica de Total de Horas Trabajadas
graf_his_hor = fgrafhis(dat_men_jal$HOR)




#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/
#### VP DEF ####
#Deflactar valor de la producción en precios corrientes para dejar constantes
#Pasos de preparación previa para poder travajar cifras de valor de la producción

#TOTAL ENTIDADES VALOR DE PRODUCCION
#Crear variable de mes con ceros
nrb=nrow(base_csv)
m0=rep(NA,nrb)
for (i in 1:nrb){
  if (base_csv$MES[i] < 10){
    m0[i]=paste0("0",base_csv$MES[i])
  }
  else{
    m0[i]=base_csv$MES[i]
  }
}

#Crear variable de fecha con anio y mes
FECHA=paste(base_csv$ANIO,m0,sep="/")

#Base de EMIM de total de entidades con fecha y variables seleccionadas
emim_tot_ent=cbind.data.frame(FECHA,base_csv[,2:3],base_csv[,5],base_csv[,10])
colnames(emim_tot_ent)[4:5]=c("ID_ENT","VP")

#Nombres de Entidades Federativas
# nom_ent=read.csv("ID_ENT_NOM_ENT.csv")
nom_ent=data.frame(matrix(c(1:32, "Aguascalientes", "Baja California", "Baja California Sur", "Campeche", "Coahuila", "Colima", "Chiapas", "Chihuahua", "Ciudad de México", "Durango", "Guanajuato", "Guerrero", "Hidalgo", "Jalisco", "Estado de México", "Michoacán", "Morelos", "Nayarit", "Nuevo León", "Oaxaca", "Puebla", "Querétaro", "Quintana Roo", "San Luis Potosí", "Sinaloa", "Sonora", "Tabasco", "Tamaulipas", "Tlaxcala", "Veracruz", "Yucatán", "Zacatecas"),32,2))
colnames(nom_ent)=c("ID_ENT",	"NOM_ENT")

#Nombres de Entidades Federativas
emim_tot_ent=merge(emim_tot_ent,nom_ent, by="ID_ENT",all=TRUE)


#Deflactor del productor con descarga de INEGI
#fim=73+(año-2019)*12+mes #fila del índice del mes
fim=2+(año-2018)*12+mes #fila del índice del mes
token="539a68cd-649d-087f-c7cf-0ca99f81093b"
inpps=inegi_series("673099",token) #Actividades secundarias con petróleo (Índice base julio 2019 = 100)
def_prod=inpps[2:fim,c(1,3)] #Recorta al mes anterior para que el INPP coincida con la EMIM
def_prod=def_prod[order(def_prod$date),]
colnames(def_prod)=c("FECHA","DEF_PROD")
#Cambiar formato de fecha para poder hacer el merge
def_prod$FECHA=format(as.Date(def_prod$FECHA),'%Y/%m')
def_prod$DEF_PROD=as.numeric(as.character(def_prod$DEF_PROD))
#def_prod$DEF_PROD=def_prod$DEF_PROD/100
def_prod$DEF_PROD=def_prod$DEF_PROD/def_prod$DEF_PROD[length(def_prod$DEF_PROD)]

#Unir deflactor a valor de producción
emim_tot_ent=merge(emim_tot_ent,def_prod, by="FECHA",all=TRUE)
emim_tot_ent=emim_tot_ent[,c(1,3,4,2,6,5,7)]
emim_tot_ent=emim_tot_ent[order(emim_tot_ent$FECHA,emim_tot_ent$ID_ENT),]
rownames(emim_tot_ent)=c()
emim_tot_ent <- na.omit(emim_tot_ent) #Quitar las filas que no coinciden


#Valor de la produccion deflactada en millones de pesos (constantes a precios del último periodo)
emim_tot_ent=mutate(emim_tot_ent,VPD=(emim_tot_ent$VP/1000)/emim_tot_ent$DEF_PROD) #Valor de la produccion deflactada
emim_tot_ent=select(emim_tot_ent,-VP,-DEF_PROD)




#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/
#### VP HIS JAL ####
#Datos mensuales de Jalisco de valor de la produccion en la entidad en millones de pesos a precios constantes
vp_men_jal=filter(emim_tot_ent,ID_ENT==14) %>%
  select(ANIO,MES,VP=VPD) 

#Cambiar meses de número a letras
vp_men_jal$MES=gsub(10,"OCT",vp_men_jal$MES)
vp_men_jal$MES=gsub(11,"NOV",vp_men_jal$MES)
vp_men_jal$MES=gsub(12,"DIC",vp_men_jal$MES)
vp_men_jal$MES=gsub(1,"ENE",vp_men_jal$MES)
vp_men_jal$MES=gsub(2,"FEB",vp_men_jal$MES)
vp_men_jal$MES=gsub(3,"MAR",vp_men_jal$MES)
vp_men_jal$MES=gsub(4,"ABR",vp_men_jal$MES)
vp_men_jal$MES=gsub(5,"MAY",vp_men_jal$MES)
vp_men_jal$MES=gsub(6,"JUN",vp_men_jal$MES)
vp_men_jal$MES=gsub(7,"JUL",vp_men_jal$MES)
vp_men_jal$MES=gsub(8,"AGO",vp_men_jal$MES)
vp_men_jal$MES=gsub(9,"SEP",vp_men_jal$MES)

#Gráfica Histórica de Personal Ocupado Total
graf_his_vp = fgrafhis(vp_men_jal$VP)




#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/
#### RANK GENERAL ####

# Datos para rankings
dat_pr = base_csv %>% filter(ANIO == año | ANIO == año-1) %>%
  select(ANIO,MES,ID_ENT=CODIGO_ENTIDAD,POT=PO_T5,HOR=HH_T5) %>%
  left_join(select(emim_tot_ent,ANIO,MES,ID_ENT,NOM_ENT,VP=VPD), by=c('ANIO'='ANIO','MES'='MES','ID_ENT'='ID_ENT')) %>%
  relocate(NOM_ENT,.after = ID_ENT)

#Función para Rankings
franking <- function(dat,vari){
  #Datos
  #Ultimo dato
  base_ud=filter(dat,ANIO==año,MES==mes) %>%
    select(ANIO,MES,ID_ENT,NOM_ENT,VAR=`vari`)
  #Nacional
  base_ud_nac=base_ud[1,]
  base_ud_nac$ID_ENT=0
  base_ud_nac$NOM_ENT="Nacional"
  base_ud_nac$VAR=sum(base_ud$VAR)
  base_ud=rbind.data.frame(base_ud,base_ud_nac)
  #Datos de variable de interés
  var_ult_dat=base_ud[,5]
  
  #Año anterior
  base_aa=filter(dat,ANIO==año-1,MES==mes) %>%
    select(ANIO,MES,ID_ENT,NOM_ENT,VAR=`vari`)
  #Nacional
  base_aa_nac=base_aa[1,]
  base_aa_nac$ID_ENT=0
  base_aa_nac$NOM_ENT="Nacional"
  base_aa_nac$VAR=sum(base_aa$VAR)
  base_aa=rbind.data.frame(base_aa,base_aa_nac)
  #Datos de variable de interés
  var_año_ant=base_aa[,5]
  
  #Unión de bases
  var_ult_ant = cbind.data.frame(NOM_ENT=base_ud$NOM_ENT,ULT=var_ult_dat,ANT=var_año_ant,VAR=(var_ult_dat/var_año_ant-1)*100)
  
  #Grafica para ranking de variación porcentual
  graf_rank_var = select(var_ult_ant,NOM_ENT,VAR) %>% arrange(VAR)
  
  return(graf_rank_var)
}




#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/#/
#### RANK POT HOR VP ####
#Rankings de Variación Anual de Personal ocupado total (PO_T5), Total de horas trabajadas (HH_T5) y Valor de la producción deflactada (VP)

#Ranking de Variación Anual de Personal Ocupado Total
graf_rank_pot = franking(dat_pr,'POT')

#Ranking de Variación Anual de Total de Horas Trabajadas
graf_rank_hor = franking(dat_pr,'HOR')

#Ranking de Variación Anual de Valor de la producción
graf_rank_vp = franking(dat_pr,'VP')



#### TEXTO DESCRIPCION ####

##### TEXTO DESCRIPCIÓN 1: graf_his_pot #####

#variables necesarias para la descripción
mesp = mespal[mes]
ud = graf_his_pot[nrow(graf_his_pot), 3] #Último dato
md = graf_his_pot[nrow(graf_his_pot)-1,3] #Dato mes anterior
ad = graf_his_pot[nrow(graf_his_pot)-12,3] #Dato año anterior
amd = graf_his_pot[nrow(graf_his_pot)-13,3] #Dato año y un mes anterior
a2d = graf_his_pot[nrow(graf_his_pot)-24,3] #Dato anterior 2 años antes
uv = round((ud/ad-1)*100,1) #Última variación anual
av = round((md/amd-1)*100,1) #Anterior variación anual (mes anterior)
up = round(graf_his_pot[nrow(graf_his_pot),4],0) #Último promedio
ap = round(graf_his_pot[nrow(graf_his_pot)-1,4],0) #Promedio anterior
vaa = round((ad/a2d-1)*100,1) #Variación año anterior


#Descripción
descripcion_1_1 <- "De acuerdo con datos de la Encuesta Mensual de la Industria Manufacturera (EMIM), el personal ocupado en la industria manufacturera de Jalisco "

if(ud > md){
  descripcion_1_1 <- paste0(descripcion_1_1, "aumentó de ", format(md, big.mark = ","), " a ", format(ud, big.mark = ","), " en ", 
                          mesp, " de ", año, " respecto al mes inmediato anterior. ")
}else if(ud < md){
  descripcion_1_1 <- paste0(descripcion_1_1, "disminuyó de ", format(md, big.mark = ","), " a ", format(ud, big.mark = ","), " en ", 
                          mesp, " de ", año, " respecto al mes inmediato anterior. ")
}else{
  descripcion_1_1 <- paste0(descripcion_1_1, " fue de ", format(ud, big.mark = ","), " trabajadores ",
                          "en línea con la cifra del mes inmediato anterior. ")
}


#conector

if(ud > md){
  if(ud > ad){
    conector <- "Además, "
  }else{
    conector <- "Sin embargo, "
  }
}else{ #ud <= md
  if(ud > ad){
    conector <- "Sin embargo, "
  }else{
    conector <- "Además, "
  }
}

if(ud > ad){
  descripcion_1_1 <- paste0(descripcion_1_1, conector, "esta cifra fue superior a la de ",
                          mesp, " de ", año-1, ", la cual se ubicaba en ", format(ad, big.mark = ","),
                          " personas ocupadas. Esto representó un crecimiento anual de ", format(uv,nsmall = 1),
                          "%")
}else if(ud < ad){
  descripcion_1_1 <- paste0(descripcion_1_1, conector, "esta cifra fue inferior a la de ",
                          mesp, " de ", año-1, ", la cual se ubicaba en ", format(ad, big.mark = ","),
                          " personas ocupadas. Esto representó una reducción anual de ", format(abs(uv),nsmall = 1),
                          "%")
}else{
  descripcion_1_1 <- paste0(descripcion_1_1, conector, "esta cifra se mantuvo en línea con la de ", mesp,
                          " de ", año-1)
}


# 3 escenarios uv > av y 3 escenarios uv < av
  if(uv > av){
    if(uv > 0 & av > 0){
      descripcion_1_1 <- paste0(descripcion_1_1, ", crecimiento superior al observado el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
    }else if(uv > 0 & av < 0){
      descripcion_1_1 <- paste0(descripcion_1_1, ", cifra superior a la observada el mes anterior, ",
                              "cuando se presentó una disminución de ", format(abs(av),nsmall = 1), "% anual.")
    }else if(uv < 0 & av < 0){
      descripcion_1_1 <- paste0(descripcion_1_1, ", caída menor a la observada el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
    }
  }else if(uv < av){
    if(uv > 0 & av > 0){
      descripcion_1_1 <- paste0(descripcion_1_1, ", crecimiento inferior al observado el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
    }else if(uv < 0 & av > 0){
      descripcion_1_1 <- paste0(descripcion_1_1, ", cifra inferior a la observada el mes anterior, ",
                              "cuando se presentó un crecimiento de ", format(abs(av),nsmall = 1), "% anual.")
    }else if(uv < 0 & av < 0){
      descripcion_1_1 <- paste0(descripcion_1_1, ", caída mayor a la observada el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
    }
  }else{
    descripcion_1_1 <- paste0(descripcion_1_1, ", en línea con la cifra observada el mes anterior.")
  }


# Por otra parte,
descripcion_1_1 <- paste0(descripcion_1_1, " Por otra parte, ")
if(uv >= 0){
  if(uv > vaa & vaa > 0){
    descripcion_1_1 <- paste0(descripcion_1_1,"el crecimiento anual del personal ocupado de ")
    descripcion_1_1 <- paste0(descripcion_1_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_1_1 <- paste0(descripcion_1_1, "fue superior al de ", mesp, " de ", año-1, ", ")
    descripcion_1_1 <- paste0(descripcion_1_1, "cuando se presentó un incremento de ", format(abs(vaa),nsmall = 1), "%. ")
  }else if(uv > vaa & vaa < 0){
    descripcion_1_1 <- paste0(descripcion_1_1,"la variación anual del personal ocupado de ")
    descripcion_1_1 <- paste0(descripcion_1_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_1_1 <- paste0(descripcion_1_1, "fue superior a la de ", mesp, " de ", año-1, ", ")
    descripcion_1_1 <- paste0(descripcion_1_1, "cuando se presentó una disminución de ", format(abs(vaa),nsmall = 1), "%. ")
  }else if(uv > vaa & vaa == 0){
    descripcion_1_1 <- paste0(descripcion_1_1,"el crecimiento del personal ocupado de ")
    descripcion_1_1 <- paste0(descripcion_1_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_1_1 <- paste0(descripcion_1_1, "fue superior al de ", mesp, " de ", año-1, ", ")
    descripcion_1_1 <- paste0(descripcion_1_1, "cuando no se presentó variación respecto a ", mesp, " de ", año-2, ". ")
  }else if(uv < vaa & vaa > 0){
    descripcion_1_1 <- paste0(descripcion_1_1,"el crecimiento anual del personal ocupado de ")
    descripcion_1_1 <- paste0(descripcion_1_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_1_1 <- paste0(descripcion_1_1, "fue inferior al de ", mesp, " de ", año-1, ", ")
    descripcion_1_1 <- paste0(descripcion_1_1, "cuando se presentó un incremento de ", format(abs(vaa),nsmall = 1), "%. ")
  }
}else if(uv < 0){
  if(uv > vaa){
    descripcion_1_1 <- paste0(descripcion_1_1,"la disminución anual del personal ocupado de ")
    descripcion_1_1 <- paste0(descripcion_1_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_1_1 <- paste0(descripcion_1_1, "fue menor a la caída de ", mesp, " de ", año-1, ", ")
    descripcion_1_1 <- paste0(descripcion_1_1, "cuando se presentó una reducción de ", format(abs(vaa),nsmall = 1), "%. ")
  }else if(uv < vaa & vaa < 0){
    descripcion_1_1 <- paste0(descripcion_1_1,"la disminución  anual del personal ocupado de ")
    descripcion_1_1 <- paste0(descripcion_1_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_1_1 <- paste0(descripcion_1_1, "fue mayor a la caída de ", mesp, " de ", año-1, ", ")
    descripcion_1_1 <- paste0(descripcion_1_1, "cuando se presentó una reducción de ", format(abs(vaa),nsmall = 1), "%. ")
  }else if(uv < vaa & vaa >= 0){
    descripcion_1_1 <- paste0(descripcion_1_1,"la variación del personal ocupado de ")
    descripcion_1_1 <- paste0(descripcion_1_1, mesp, " de ", año, ", de ", format(uv,nsmall = 1), "%, ")
    descripcion_1_1 <- paste0(descripcion_1_1, "fue inferior a la de ", mesp, " de ", año-1, ", ")
    descripcion_1_1 <- paste0(descripcion_1_1, "cuando se presentó un crecimiento de ", format(vaa,nsmall = 1), "%. ")
  }
}


#El promedio de los últimos 12 meses
descripcion_1_2 <- "El promedio de los últimos 12 meses "

#Promedio
if(up > ap){#Aumenta promedio
  ac = 0 #aumento consecutivo
  while(graf_his_pot[nrow(graf_his_pot)-ac,4] > (graf_his_pot[nrow(graf_his_pot)-1-ac,4])){
    ac = ac + 1
  }
  dcp = 0 #descenso consecutivo previo
  while(graf_his_pot[nrow(graf_his_pot)-ac-1-dcp,4] < graf_his_pot[nrow(graf_his_pot)-ac-2-dcp,4]){
    dcp = dcp + 1
  }
  dcp = dcp + 1
  
  if(ac == 1 & dcp <3){
    descripcion_1_2 = paste0(descripcion_1_2,"registró un aumento")
  }else if(ac == 1 & dcp >= 3){
    descripcion_1_2 = paste0(descripcion_1_2,"registró su primer aumento después de ",dcp,
                             " meses consecutivos de descensos")
  }else if(ac > 1 & ac <= 20){
    descripcion_1_2 = paste0(descripcion_1_2,"aumentó por ",num_cardinales[ac]," mes consecutivo")
  }else if(ac > 1 & ac > 20){
    descripcion_1_2 = paste0(descripcion_1_2,"mantiene su tendencia creciente")
  }
  descripcion_1_2 = paste0(descripcion_1_2,", al pasar de ",format(ap,big.mark=","), " a ",format(up,big.mark=","), " ocupados en ",mesp, " de ",
                           año, " respecto al mes inmediato anterior.")
}else if( up < ap){#Disminuye promedio
  dc = 0 #descenso consecutivo
  while(graf_his_pot[nrow(graf_his_pot)-dc,4] < (graf_his_pot[nrow(graf_his_pot)-1-dc,4])){
    dc = dc + 1
  }
  acp = 0 #ascenso consecutivo previo
  while(graf_his_pot[nrow(graf_his_pot)-dc-1-acp,4] > graf_his_pot[nrow(graf_his_pot)-dc-2-acp,4]){
    acp = acp + 1
  }
  acp = acp + 1
  
  if(dc == 1 & acp <3){
    descripcion_1_2 = paste0(descripcion_1_2,"registró un descenso")
  }else if(dc == 1 & acp >= 3){
    descripcion_1_2 = paste0(descripcion_1_2,"registró su primer descenso después de ",acp,
                             " meses consecutivos de incrementos")
  }else if(dc > 1 & dc <= 20){
    descripcion_1_2 = paste0(descripcion_1_2,"disminuyó por ",num_cardinales[dc]," mes consecutivo")
  }else if(dc > 1 & dc > 20){
    descripcion_1_2 = paste0(descripcion_1_2,"mantiene su tendencia decreciente")
  }
  descripcion_1_2 = paste0(descripcion_1_2,", al pasar de ",format(ap,big.mark=","), " a ",format(up,big.mark=","), " ocupados en ",mesp, " de ",
                           año, " respecto al mes inmediato anterior.")
}else{
  descripcion_1_2 = paste0(descripcion_1_2,conector1,"se mantuvo sin cambios.")
}

descripcion_1 = paste0(descripcion_1_1,descripcion_1_2)




##### TEXTO DESCRIPCIÓN 2: graf_rank_pot #####
#Variables
mesactual= mespal[as.numeric(mes)]
udatjal <- round(graf_rank_pot[graf_rank_pot$NOM_ENT=="Jalisco",2],1) #último dato de variación anual de Jalisco
datprom <- round(graf_rank_pot[graf_rank_pot$NOM_ENT=="Nacional",2],1) #dato del promedio nacional
rank_pot = graf_rank_pot[graf_rank_pot$NOM_ENT!="Nacional",] %>% arrange(desc(VAR)) #Ranking sin dato nacional
lug_pot = which(rank_pot$NOM_ENT=="Jalisco") #Lugar del ranking


#Descripción
descripcion_2 <- paste0("Por otra parte, el personal ocupado en la industria manufacturera de Jalisco que")

if (udatjal > 0){
  descripcion_2 = paste0(descripcion_2," aumentó ", format(abs(udatjal),nsmall = 1), "% a tasa anual en ", mesactual, " de ", año, ", ")
} else if (udatjal < 0){
  descripcion_2 = paste0(descripcion_2," disminuyó ", format(abs(udatjal),nsmall = 1), "% a tasa anual en ", mesactual, " de ", año, ", ")
} else {
  descripcion_2 = paste0(descripcion_2," se mantuvo sin cambios en ", mesactual, " de ", año, ", ")
}


if (udatjal > datprom & udatjal > 0 & datprom > 0){
  descripcion_2 = paste0(descripcion_2,"presentó un crecimiento superior al promedio nacional de ",format(datprom,nsmall = 1),"%")
} else if (udatjal < datprom & udatjal >= 0 & datprom > 0){
  descripcion_2 = paste0(descripcion_2,"presentó un crecimiento inferior al promedio nacional de ",format(datprom,nsmall = 1),"%")
} else if (udatjal > datprom & udatjal > 0 & datprom < 0){
  descripcion_2 = paste0(descripcion_2,"presentó un crecimiento superior a la variación nacional de ",format(datprom,nsmall = 1),"%")
} else if (udatjal > datprom & udatjal > 0 & datprom == 0){
  descripcion_2 = paste0(descripcion_2,"presentó un crecimiento superior a la cifra nacional de ", format(abs(datprom),nsmall = 1),"%")
} else if (udatjal < datprom & udatjal < 0 & datprom < 0){
  descripcion_2 = paste0(descripcion_2,"presentó una caída mayor a la variación nacional de ", format(datprom,nsmall = 1),"%")
} else if (udatjal > datprom & udatjal < 0 & datprom < 0){
  descripcion_2 = paste0(descripcion_2,"presentó una caída menor al descenso nacional de ", format(abs(datprom),nsmall = 1),"%")
} else if (udatjal < datprom & udatjal < 0 & datprom > 0){
  descripcion_2 = paste0(descripcion_2,"presentó una cifra inferior a la variación nacional de ", format(datprom,nsmall = 1),"%")
} else if (udatjal < datprom & udatjal < 0 & datprom == 0){
  descripcion_2 = paste0(descripcion_2,"presentó una cifra inferior a la variación nacional de ", format(abs(datprom),nsmall = 1),"%")
} else if (udatjal == datprom & udatjal == 0 & datprom == 0){
  descripcion_2 = paste0(descripcion_2,"en línea con el comportamiento nacional que tampoco presentó variación")
}

#Posición en el ranking
descripcion_2 <- paste0(descripcion_2," y ubicó a Jalisco en el ")
if (lug_pot <= 20) {
  descripcion_2 = paste0(descripcion_2,num_cardinales[lug_pot]," lugar a nivel nacional en cuanto a crecimiento del empleo en esta industria.")
} else {
  descripcion_2 = paste0(descripcion_2,"lugar ",lug_pot," a nivel nacional en cuanto a crecimiento del empleo en esta industria.")
}



##### TEXTO DESCRIPCIÓN 3: graf_his_hor #####

#variables necesarias para la descripción
mesp = mespal[mes]
ud = graf_his_hor[nrow(graf_his_hor), 3] #Último dato
md = graf_his_hor[nrow(graf_his_hor)-1,3] #Dato mes anterior
ad = graf_his_hor[nrow(graf_his_hor)-12,3] #Dato año anterior
amd = graf_his_hor[nrow(graf_his_hor)-13,3] #Dato año y un mes anterior
a2d = graf_his_hor[nrow(graf_his_hor)-24,3] #Dato anterior 2 años antes
uv = round((ud/ad-1)*100,1) #Última variación anual
av = round((md/amd-1)*100,1) #Anterior variación anual (mes anterior)
up = round(graf_his_hor[nrow(graf_his_hor),4],0) #Último promedio
ap = round(graf_his_hor[nrow(graf_his_hor)-1,4],0) #Promedio anterior
vaa = round((ad/a2d-1)*100,1) #Variación año anterior

#Nota: como las horas están en miles de horas, se divide entre 1,000 para que sean millones de horas, pero
#esto se hace después de sacar la variación para poder hacer el cálculo de variación con todos los decimales
ud = ud/1000
md = md/1000
ad = ad/1000
up = up/1000
ap = ap/1000

#Descripción
descripcion_3_1=paste0("En cuanto a las horas trabajadas por el personal ocupado total en la industria manufacturera de Jalisco, el número de horas ")


if(ud > md){
  descripcion_3_1 <- paste0(descripcion_3_1, "aumentó de ", format(round(md,2), nsmall = 2), " a ", format(round(ud,2), nsmall = 2), " millones de horas trabajadas en ", 
                            mesp, " de ", año, " respecto al mes inmediato anterior. ")
}else if(ud < md){
  descripcion_3_1 <- paste0(descripcion_3_1, "disminuyó de ", format(round(md,2), nsmall = 2), " a ", format(round(ud,2), nsmall = 2), " millones de horas trabajadas en ", 
                            mesp, " de ", año, " respecto al mes inmediato anterior. ")
}else{
  descripcion_3_1 <- paste0(descripcion_3_1, " fue de ", format(round(ud,2), nsmall = 2), " millones de horas trabajadas ",
                            "en línea con la cifra del mes inmediato anterior. ")
}


#conector

if(ud > md){
  if(ud > ad){
    conector <- "Además, "
  }else{
    conector <- "Sin embargo, "
  }
}else{ #ud <= md
  if(ud > ad){
    conector <- "Sin embargo, "
  }else{
    conector <- "Además, "
  }
}

if(ud > ad){
  descripcion_3_1 <- paste0(descripcion_3_1, conector, "esta cifra fue superior a la de ",
                            mesp, " de ", año-1, ", la cual se ubicaba en ", format(round(ad,2), nsmall = 2),
                            " millones de horas trabajadas. Esto representó un crecimiento anual de ", format(uv,nsmall = 1),
                            "%")
}else if(ud < ad){
  descripcion_3_1 <- paste0(descripcion_3_1, conector, "esta cifra fue inferior a la de ",
                            mesp, " de ", año-1, ", la cual se ubicaba en ", format(round(ad,2), nsmall = 2),
                            " millones de horas trabajadas. Esto representó una reducción anual de ", format(abs(uv),nsmall = 1),
                            "%")
}else{
  descripcion_3_1 <- paste0(descripcion_3_1, conector, "esta cifra se mantuvo en línea con la de ", mesp,
                            " de ", año-1)
}


# 3 escenarios uv > av y 3 escenarios uv < av
if(uv > av){
  if(uv > 0 & av > 0){
    descripcion_3_1 <- paste0(descripcion_3_1, ", crecimiento superior al observado el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
  }else if(uv > 0 & av < 0){
    descripcion_3_1 <- paste0(descripcion_3_1, ", cifra superior a la observada el mes anterior, ",
                              "cuando se presentó una disminución de ", format(abs(av),nsmall = 1), "% anual.")
  }else if(uv < 0 & av < 0){
    descripcion_3_1 <- paste0(descripcion_3_1, ", caída menor a la observada el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
  }
}else if(uv < av){
  if(uv > 0 & av > 0){
    descripcion_3_1 <- paste0(descripcion_3_1, ", crecimiento inferior al observado el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
  }else if(uv < 0 & av > 0){
    descripcion_3_1 <- paste0(descripcion_3_1, ", cifra inferior a la observada el mes anterior, ",
                              "cuando se presentó un crecimiento de ", format(abs(av),nsmall = 1), "% anual.")
  }else if(uv < 0 & av < 0){
    descripcion_3_1 <- paste0(descripcion_3_1, ", caída mayor a la observada el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
  }
}else{
  descripcion_3_1 <- paste0(descripcion_3_1, ", en línea con la cifra observada el mes anterior.")
}


# Por otra parte,
descripcion_3_1 <- paste0(descripcion_3_1, " Por otra parte, ")
if(uv >= 0){
  if(uv > vaa & vaa > 0){
    descripcion_3_1 <- paste0(descripcion_3_1,"el crecimiento anual de las horas trabajadas de ")
    descripcion_3_1 <- paste0(descripcion_3_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_3_1 <- paste0(descripcion_3_1, "fue superior al de ", mesp, " de ", año-1, ", ")
    descripcion_3_1 <- paste0(descripcion_3_1, "cuando se presentó un incremento de ", format(abs(vaa),nsmall = 1), "%. ")
  }else if(uv > vaa & vaa < 0){
    descripcion_3_1 <- paste0(descripcion_3_1,"la variación anual de las horas trabajadas de ")
    descripcion_3_1 <- paste0(descripcion_3_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_3_1 <- paste0(descripcion_3_1, "fue superior a la de ", mesp, " de ", año-1, ", ")
    descripcion_3_1 <- paste0(descripcion_3_1, "cuando se presentó una disminución de ", format(abs(vaa),nsmall = 1), "%. ")
  }else if(uv > vaa & vaa == 0){
    descripcion_3_1 <- paste0(descripcion_3_1,"el crecimiento de las horas trabajadas de ")
    descripcion_3_1 <- paste0(descripcion_3_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_3_1 <- paste0(descripcion_3_1, "fue superior al de ", mesp, " de ", año-1, ", ")
    descripcion_3_1 <- paste0(descripcion_3_1, "cuando no se presentó variación respecto a ", mesp, " de ", año-2, ". ")
  }else if(uv < vaa & vaa > 0){
    descripcion_3_1 <- paste0(descripcion_3_1,"el crecimiento anual de las horas trabajadas de ")
    descripcion_3_1 <- paste0(descripcion_3_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_3_1 <- paste0(descripcion_3_1, "fue inferior al de ", mesp, " de ", año-1, ", ")
    descripcion_3_1 <- paste0(descripcion_3_1, "cuando se presentó un incremento de ", format(abs(vaa),nsmall = 1), "%. ")
  }
}else if(uv < 0){
  if(uv > vaa){
    descripcion_3_1 <- paste0(descripcion_3_1,"la disminución anual de las horas trabajadas de ")
    descripcion_3_1 <- paste0(descripcion_3_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_3_1 <- paste0(descripcion_3_1, "fue menor a la caída de ", mesp, " de ", año-1, ", ")
    descripcion_3_1 <- paste0(descripcion_3_1, "cuando se presentó una reducción de ", format(abs(vaa),nsmall = 1), "%. ")
  }else if(uv < vaa & vaa < 0){
    descripcion_3_1 <- paste0(descripcion_3_1,"la disminución  anual de las horas trabajadas de ")
    descripcion_3_1 <- paste0(descripcion_3_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_3_1 <- paste0(descripcion_3_1, "fue mayor a la caída de ", mesp, " de ", año-1, ", ")
    descripcion_3_1 <- paste0(descripcion_3_1, "cuando se presentó una reducción de ", format(abs(vaa),nsmall = 1), "%. ")
  }else if(uv < vaa & vaa >= 0){
    descripcion_3_1 <- paste0(descripcion_3_1,"la variación de las horas trabajadas de ")
    descripcion_3_1 <- paste0(descripcion_3_1, mesp, " de ", año, ", de ", format(uv,nsmall = 1), "%, ")
    descripcion_3_1 <- paste0(descripcion_3_1, "fue inferior a la de ", mesp, " de ", año-1, ", ")
    descripcion_3_1 <- paste0(descripcion_3_1, "cuando se presentó un crecimiento de ", format(vaa,nsmall = 1), "%. ")
  }
}


#El promedio de los últimos 12 meses
descripcion_3_2 <- "El promedio de los últimos 12 meses "

#Promedio
if(up > ap){#Aumenta promedio
  ac = 0 #aumento consecutivo
  while(graf_his_hor[nrow(graf_his_hor)-ac,4] > (graf_his_hor[nrow(graf_his_hor)-1-ac,4])){
    ac = ac + 1
  }
  dcp = 0 #descenso consecutivo previo
  while(graf_his_hor[nrow(graf_his_hor)-ac-1-dcp,4] < graf_his_hor[nrow(graf_his_hor)-ac-2-dcp,4]){
    dcp = dcp + 1
  }
  dcp = dcp + 1
  
  if(ac == 1 & dcp <3){
    descripcion_3_2 = paste0(descripcion_3_2,"registró un aumento")
  }else if(ac == 1 & dcp >= 3){
    descripcion_3_2 = paste0(descripcion_3_2,"registró su primer aumento después de ",dcp,
                             " meses consecutivos de descensos")
  }else if(ac > 1 & ac <= 20){
    descripcion_3_2 = paste0(descripcion_3_2,"aumentó por ",num_cardinales[ac]," mes consecutivo")
  }else if(ac > 1 & ac > 20){
    descripcion_3_2 = paste0(descripcion_3_2,"mantiene su tendencia creciente")
  }
  descripcion_3_2 = paste0(descripcion_3_2,", al pasar de ",format(round(ap,2),nsmall = 2), " a ",format(round(up,2),nsmall = 2), " millones de horas trabajadas en ",mesp, " de ",
                           año, " respecto al mes inmediato anterior.")
}else if( up < ap){#Disminuye promedio
  dc = 0 #descenso consecutivo
  while(graf_his_hor[nrow(graf_his_hor)-dc,4] < (graf_his_hor[nrow(graf_his_hor)-1-dc,4])){
    dc = dc + 1
  }
  acp = 0 #ascenso consecutivo previo
  while(graf_his_hor[nrow(graf_his_hor)-dc-1-acp,4] > graf_his_hor[nrow(graf_his_hor)-dc-2-acp,4]){
    acp = acp + 1
  }
  acp = acp + 1
  
  if(dc == 1 & acp <3){
    descripcion_3_2 = paste0(descripcion_3_2,"registró un descenso")
  }else if(dc == 1 & acp >= 3){
    descripcion_3_2 = paste0(descripcion_3_2,"registró su primer descenso después de ",acp,
                             " meses consecutivos de incrementos")
  }else if(dc > 1 & dc <= 20){
    descripcion_3_2 = paste0(descripcion_3_2,"disminuyó por ",num_cardinales[dc]," mes consecutivo")
  }else if(dc > 1 & dc > 20){
    descripcion_3_2 = paste0(descripcion_3_2,"mantiene su tendencia decreciente")
  }
  descripcion_3_2 = paste0(descripcion_3_2,", al pasar de ",format(round(ap,2),nsmall = 2), " a ",format(round(up,2),nsmall = 2), " millones de horas trabajadas en ",mesp, " de ",
                           año, " respecto al mes inmediato anterior.")
}else{
  descripcion_3_2 = paste0(descripcion_3_2,conector1,"se mantuvo sin cambios.")
}

descripcion_3 = paste0(descripcion_3_1,descripcion_3_2)




##### TEXTO DESCRIPCIÓN 4: graf_rank_hor #####
#Variables
mesactual= mespal[as.numeric(mes)]
udatjal <- round(graf_rank_hor[graf_rank_hor$NOM_ENT=="Jalisco",2],1) #último dato de variación anual de Jalisco
datprom <- round(graf_rank_hor[graf_rank_hor$NOM_ENT=="Nacional",2],1) #dato del promedio nacional
rank_pot = graf_rank_hor[graf_rank_hor$NOM_ENT!="Nacional",] %>% arrange(desc(VAR)) #Ranking sin dato nacional
lug_pot = which(rank_pot$NOM_ENT=="Jalisco") #Lugar del ranking


#Descripción
descripcion_4 <- paste0("Por otra parte, el número de horas trabajadas por el personal ocupado en la industria manufacturera de Jalisco que")

if (udatjal > 0){
  descripcion_4 = paste0(descripcion_4," aumentó ", format(abs(udatjal),nsmall = 1), "% a tasa anual en ", mesactual, " de ", año, ", ")
} else if (udatjal < 0){
  descripcion_4 = paste0(descripcion_4," disminuyó ", format(abs(udatjal),nsmall = 1), "% a tasa anual en ", mesactual, " de ", año, ", ")
} else {
  descripcion_4 = paste0(descripcion_4," se mantuvo sin cambios en ", mesactual, " de ", año, ", ")
}


if (udatjal > datprom & udatjal > 0 & datprom > 0){
  descripcion_4 = paste0(descripcion_4,"presentó un crecimiento superior al promedio nacional de ",format(datprom,nsmall = 1),"%")
} else if (udatjal < datprom & udatjal >= 0 & datprom > 0){
  descripcion_4 = paste0(descripcion_4,"presentó un crecimiento inferior al promedio nacional de ",format(datprom,nsmall = 1),"%")
} else if (udatjal > datprom & udatjal > 0 & datprom < 0){
  descripcion_4 = paste0(descripcion_4,"presentó un crecimiento superior a la variación nacional de ",format(datprom,nsmall = 1),"%")
} else if (udatjal > datprom & udatjal > 0 & datprom == 0){
  descripcion_4 = paste0(descripcion_4,"presentó un crecimiento superior a la cifra nacional de ", format(abs(datprom),nsmall = 1),"%")
} else if (udatjal < datprom & udatjal < 0 & datprom < 0){
  descripcion_4 = paste0(descripcion_4,"presentó una caída mayor a la variación nacional de ", format(datprom,nsmall = 1),"%")
} else if (udatjal > datprom & udatjal < 0 & datprom < 0){
  descripcion_4 = paste0(descripcion_4,"presentó una caída menor al descenso nacional de ", format(abs(datprom),nsmall = 1),"%")
} else if (udatjal < datprom & udatjal < 0 & datprom > 0){
  descripcion_4 = paste0(descripcion_4,"presentó una cifra inferior a la variación nacional de ", format(datprom,nsmall = 1),"%")
} else if (udatjal < datprom & udatjal < 0 & datprom == 0){
  descripcion_4 = paste0(descripcion_4,"presentó una cifra inferior a la variación nacional de ", format(abs(datprom),nsmall = 1),"%")
} else if (udatjal == datprom & udatjal == 0 & datprom == 0){
  descripcion_4 = paste0(descripcion_4,"en línea con el comportamiento nacional que tampoco presentó variación")
}

#Posición en el ranking
descripcion_4 <- paste0(descripcion_4," y ubicó a Jalisco en el ")
if (lug_pot <= 20) {
  descripcion_4 = paste0(descripcion_4,num_cardinales[lug_pot]," lugar a nivel nacional en cuanto a crecimiento de las horas trabajadas en esta industria.")
} else {
  descripcion_4 = paste0(descripcion_4,"lugar ",lug_pot," a nivel nacional en cuanto a crecimiento de las horas trabajadas en esta industria.")
}




##### TEXTO DESCRIPCIÓN 5: graf_his_vp #####
#Nota: el valor de la producción ya está en millones de pesos constantes

#variables necesarias para la descripción
mesp = mespal[mes]
ud = graf_his_vp[nrow(graf_his_vp), 3] #Último dato
md = graf_his_vp[nrow(graf_his_vp)-1,3] #Dato mes anterior
ad = graf_his_vp[nrow(graf_his_vp)-12,3] #Dato año anterior
amd = graf_his_vp[nrow(graf_his_vp)-13,3] #Dato año y un mes anterior
a2d = graf_his_vp[nrow(graf_his_vp)-24,3] #Dato anterior 2 años antes
uv = round((ud/ad-1)*100,1) #Última variación anual
av = round((md/amd-1)*100,1) #Anterior variación anual (mes anterior)
up = round(graf_his_vp[nrow(graf_his_vp),4],0) #Último promedio
ap = round(graf_his_vp[nrow(graf_his_vp)-1,4],0) #Promedio anterior
vaa = round((ad/a2d-1)*100,1) #Variación año anterior


#Descripción
descripcion_5_1=paste0("En lo que respecta al valor de la producción de la industria manufacturera de Jalisco, el valor de la producción ")


if(ud > md){
  descripcion_5_1 <- paste0(descripcion_5_1, "aumentó de ", format(round(md), big.mark = ","), " a ", format(round(ud), big.mark = ","), " millones de pesos constantes en ", 
                            mesp, " de ", año, " respecto al mes inmediato anterior. ")
}else if(ud < md){
  descripcion_5_1 <- paste0(descripcion_5_1, "disminuyó de ", format(round(md), big.mark = ","), " a ", format(round(ud), big.mark = ","), " millones de pesos constantes en ", 
                            mesp, " de ", año, " respecto al mes inmediato anterior. ")
}else{
  descripcion_5_1 <- paste0(descripcion_5_1, " fue de ", format(round(ud), big.mark = ","), " millones de pesos constantes ",
                            "en línea con la cifra del mes inmediato anterior. ")
}


#conector

if(ud > md){
  if(ud > ad){
    conector <- "Además, "
  }else{
    conector <- "Sin embargo, "
  }
}else{ #ud <= md
  if(ud > ad){
    conector <- "Sin embargo, "
  }else{
    conector <- "Además, "
  }
}

if(ud > ad){
  descripcion_5_1 <- paste0(descripcion_5_1, conector, "esta cifra fue superior a la de ",
                            mesp, " de ", año-1, ", la cual se ubicaba en ", format(round(ad), big.mark = ","),
                            " millones de pesos constantes. Esto representó un crecimiento anual de ", format(uv,nsmall = 1),
                            "%")
}else if(ud < ad){
  descripcion_5_1 <- paste0(descripcion_5_1, conector, "esta cifra fue inferior a la de ",
                            mesp, " de ", año-1, ", la cual se ubicaba en ", format(round(ad), big.mark = ","),
                            " millones de pesos constantes. Esto representó una reducción anual de ", format(abs(uv),nsmall = 1),
                            "%")
}else{
  descripcion_5_1 <- paste0(descripcion_5_1, conector, "esta cifra se mantuvo en línea con la de ", mesp,
                            " de ", año-1)
}


# 3 escenarios uv > av y 3 escenarios uv < av
if(uv > av){
  if(uv > 0 & av > 0){
    descripcion_5_1 <- paste0(descripcion_5_1, ", crecimiento superior al observado el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
  }else if(uv > 0 & av < 0){
    descripcion_5_1 <- paste0(descripcion_5_1, ", cifra superior a la observada el mes anterior, ",
                              "cuando se presentó una disminución de ", format(abs(av),nsmall = 1), "% anual.")
  }else if(uv < 0 & av < 0){
    descripcion_5_1 <- paste0(descripcion_5_1, ", caída menor a la observada el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
  }
}else if(uv < av){
  if(uv > 0 & av > 0){
    descripcion_5_1 <- paste0(descripcion_5_1, ", crecimiento inferior al observado el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
  }else if(uv < 0 & av > 0){
    descripcion_5_1 <- paste0(descripcion_5_1, ", cifra inferior a la observada el mes anterior, ",
                              "cuando se presentó un crecimiento de ", format(abs(av),nsmall = 1), "% anual.")
  }else if(uv < 0 & av < 0){
    descripcion_5_1 <- paste0(descripcion_5_1, ", caída mayor a la observada el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
  }
}else{
  descripcion_5_1 <- paste0(descripcion_5_1, ", en línea con la cifra observada el mes anterior.")
}


# Por otra parte,
descripcion_5_1 <- paste0(descripcion_5_1, " Por otra parte, ")
if(uv >= 0){
  if(uv > vaa & vaa > 0){
    descripcion_5_1 <- paste0(descripcion_5_1,"el crecimiento anual del valor de la producción de ")
    descripcion_5_1 <- paste0(descripcion_5_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_5_1 <- paste0(descripcion_5_1, "fue superior al de ", mesp, " de ", año-1, ", ")
    descripcion_5_1 <- paste0(descripcion_5_1, "cuando se presentó un incremento de ", format(abs(vaa),nsmall = 1), "%. ")
  }else if(uv > vaa & vaa < 0){
    descripcion_5_1 <- paste0(descripcion_5_1,"la variación anual del valor de la producción de ")
    descripcion_5_1 <- paste0(descripcion_5_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_5_1 <- paste0(descripcion_5_1, "fue superior a la de ", mesp, " de ", año-1, ", ")
    descripcion_5_1 <- paste0(descripcion_5_1, "cuando se presentó una disminución de ", format(abs(vaa),nsmall = 1), "%. ")
  }else if(uv > vaa & vaa == 0){
    descripcion_5_1 <- paste0(descripcion_5_1,"el crecimiento anual del valor de la producción de ")
    descripcion_5_1 <- paste0(descripcion_5_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_5_1 <- paste0(descripcion_5_1, "fue superior al de ", mesp, " de ", año-1, ", ")
    descripcion_5_1 <- paste0(descripcion_5_1, "cuando no se presentó variación respecto a ", mesp, " de ", año-2, ". ")
  }else if(uv < vaa & vaa > 0){
    descripcion_5_1 <- paste0(descripcion_5_1,"el crecimiento anual del valor de la producción de ")
    descripcion_5_1 <- paste0(descripcion_5_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_5_1 <- paste0(descripcion_5_1, "fue inferior al de ", mesp, " de ", año-1, ", ")
    descripcion_5_1 <- paste0(descripcion_5_1, "cuando se presentó un incremento de ", format(abs(vaa),nsmall = 1), "%. ")
  }
}else if(uv < 0){
  if(uv > vaa){
    descripcion_5_1 <- paste0(descripcion_5_1,"la disminución anual del valor de la producción de ")
    descripcion_5_1 <- paste0(descripcion_5_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_5_1 <- paste0(descripcion_5_1, "fue menor a la caída de ", mesp, " de ", año-1, ", ")
    descripcion_5_1 <- paste0(descripcion_5_1, "cuando se presentó una reducción de ", format(abs(vaa),nsmall = 1), "%. ")
  }else if(uv < vaa & vaa < 0){
    descripcion_5_1 <- paste0(descripcion_5_1,"la disminución anual del valor de la producción de ")
    descripcion_5_1 <- paste0(descripcion_5_1, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_5_1 <- paste0(descripcion_5_1, "fue mayor a la caída de ", mesp, " de ", año-1, ", ")
    descripcion_5_1 <- paste0(descripcion_5_1, "cuando se presentó una reducción de ", format(abs(vaa),nsmall = 1), "%. ")
  }else if(uv < vaa & vaa >= 0){
    descripcion_5_1 <- paste0(descripcion_5_1,"la variación anual del valor de la producción de ")
    descripcion_5_1 <- paste0(descripcion_5_1, mesp, " de ", año, ", de ", format(uv,nsmall = 1), "%, ")
    descripcion_5_1 <- paste0(descripcion_5_1, "fue inferior a la de ", mesp, " de ", año-1, ", ")
    descripcion_5_1 <- paste0(descripcion_5_1, "cuando se presentó un crecimiento de ", format(vaa,nsmall = 1), "%. ")
  }
}


#El promedio de los últimos 12 meses
descripcion_5_2 <- "El promedio de los últimos 12 meses "

#Promedio
if(up > ap){#Aumenta promedio
  ac = 0 #aumento consecutivo
  while(graf_his_vp[nrow(graf_his_vp)-ac,4] > (graf_his_vp[nrow(graf_his_vp)-1-ac,4])){
    ac = ac + 1
  }
  dcp = 0 #descenso consecutivo previo
  while(graf_his_vp[nrow(graf_his_vp)-ac-1-dcp,4] < graf_his_vp[nrow(graf_his_vp)-ac-2-dcp,4]){
    dcp = dcp + 1
  }
  dcp = dcp + 1
  
  if(ac == 1 & dcp <3){
    descripcion_5_2 = paste0(descripcion_5_2,"registró un aumento")
  }else if(ac == 1 & dcp >= 3){
    descripcion_5_2 = paste0(descripcion_5_2,"registró su primer aumento después de ",dcp,
                             " meses consecutivos de descensos")
  }else if(ac > 1 & ac <= 20){
    descripcion_5_2 = paste0(descripcion_5_2,"aumentó por ",num_cardinales[ac]," mes consecutivo")
  }else if(ac > 1 & ac > 20){
    descripcion_5_2 = paste0(descripcion_5_2,"mantiene su tendencia creciente")
  }
  descripcion_5_2 = paste0(descripcion_5_2,", al pasar de ",format(round(ap), big.mark = ","), " a ",format(round(up), big.mark = ","), " millones de pesos constantes en ",mesp, " de ",
                           año, " respecto al mes inmediato anterior.")
}else if( up < ap){#Disminuye promedio
  dc = 0 #descenso consecutivo
  while(graf_his_vp[nrow(graf_his_vp)-dc,4] < (graf_his_vp[nrow(graf_his_vp)-1-dc,4])){
    dc = dc + 1
  }
  acp = 0 #ascenso consecutivo previo
  while(graf_his_vp[nrow(graf_his_vp)-dc-1-acp,4] > graf_his_vp[nrow(graf_his_vp)-dc-2-acp,4]){
    acp = acp + 1
  }
  acp = acp + 1
  
  if(dc == 1 & acp <3){
    descripcion_5_2 = paste0(descripcion_5_2,"registró un descenso")
  }else if(dc == 1 & acp >= 3){
    descripcion_5_2 = paste0(descripcion_5_2,"registró su primer descenso después de ",acp,
                             " meses consecutivos de incrementos")
  }else if(dc > 1 & dc <= 20){
    descripcion_5_2 = paste0(descripcion_5_2,"disminuyó por ",num_cardinales[dc]," mes consecutivo")
  }else if(dc > 1 & dc > 20){
    descripcion_5_2 = paste0(descripcion_5_2,"mantiene su tendencia decreciente")
  }
  descripcion_5_2 = paste0(descripcion_5_2,", al pasar de ",format(round(ap), big.mark = ","), " a ",format(round(up), big.mark = ","), " millones de pesos constantes en ",mesp, " de ",
                           año, " respecto al mes inmediato anterior.")
}else{
  descripcion_5_2 = paste0(descripcion_5_2,conector1,"se mantuvo sin cambios.")
}

descripcion_5 = paste0(descripcion_5_1,descripcion_5_2)




##### TEXTO DESCRIPCIÓN 6: graf_rank_vp #####
#Variables
mesactual= mespal[as.numeric(mes)]
udatjal <- round(graf_rank_vp[graf_rank_vp$NOM_ENT=="Jalisco",2],1) #último dato de variación anual de Jalisco
datprom <- round(graf_rank_vp[graf_rank_vp$NOM_ENT=="Nacional",2],1) #dato del promedio nacional
rank_pot = graf_rank_vp[graf_rank_vp$NOM_ENT!="Nacional",] %>% arrange(desc(VAR)) #Ranking sin dato nacional
lug_pot = which(rank_pot$NOM_ENT=="Jalisco") #Lugar del ranking


#Descripción
descripcion_6 <- paste0("Por otra parte, el valor de la producción en la industria manufacturera de Jalisco que")

if (udatjal > 0){
  descripcion_6 = paste0(descripcion_6," aumentó ", format(abs(udatjal),nsmall = 1), "% a tasa anual en ", mesactual, " de ", año, ", ")
} else if (udatjal < 0){
  descripcion_6 = paste0(descripcion_6," disminuyó ", format(abs(udatjal),nsmall = 1), "% a tasa anual en ", mesactual, " de ", año, ", ")
} else {
  descripcion_6 = paste0(descripcion_6," se mantuvo sin cambios en ", mesactual, " de ", año, ", ")
}


if (udatjal > datprom & udatjal > 0 & datprom > 0){
  descripcion_6 = paste0(descripcion_6,"presentó un crecimiento superior al promedio nacional de ",format(datprom,nsmall = 1),"%")
} else if (udatjal < datprom & udatjal >= 0 & datprom > 0){
  descripcion_6 = paste0(descripcion_6,"presentó un crecimiento inferior al promedio nacional de ",format(datprom,nsmall = 1),"%")
} else if (udatjal > datprom & udatjal > 0 & datprom < 0){
  descripcion_6 = paste0(descripcion_6,"presentó un crecimiento superior a la variación nacional de ",format(datprom,nsmall = 1),"%")
} else if (udatjal > datprom & udatjal > 0 & datprom == 0){
  descripcion_6 = paste0(descripcion_6,"presentó un crecimiento superior a la cifra nacional de ", format(abs(datprom),nsmall = 1),"%")
} else if (udatjal < datprom & udatjal < 0 & datprom < 0){
  descripcion_6 = paste0(descripcion_6,"presentó una caída mayor a la variación nacional de ", format(datprom,nsmall = 1),"%")
} else if (udatjal > datprom & udatjal < 0 & datprom < 0){
  descripcion_6 = paste0(descripcion_6,"presentó una caída menor al descenso nacional de ", format(abs(datprom),nsmall = 1),"%")
} else if (udatjal < datprom & udatjal < 0 & datprom > 0){
  descripcion_6 = paste0(descripcion_6,"presentó una cifra inferior a la variación nacional de ", format(datprom,nsmall = 1),"%")
} else if (udatjal < datprom & udatjal < 0 & datprom == 0){
  descripcion_6 = paste0(descripcion_6,"presentó una cifra inferior a la variación nacional de ", format(abs(datprom),nsmall = 1),"%")
} else if (udatjal == datprom & udatjal == 0 & datprom == 0){
  descripcion_6 = paste0(descripcion_6,"en línea con el comportamiento nacional que tampoco presentó variación")
}

#Posición en el ranking
descripcion_6 <- paste0(descripcion_6," y ubicó a Jalisco en el ")
if (lug_pot <= 20) {
  descripcion_6 = paste0(descripcion_6,num_cardinales[lug_pot]," lugar a nivel nacional en cuanto a crecimiento del valor de la producción en esta industria.")
} else {
  descripcion_6 = paste0(descripcion_6,"lugar ",lug_pot," a nivel nacional en cuanto a crecimiento del valor de la producción en esta industria.")
}




#### Exportar a Excel ####
#Periodos de titulos
periodo1=paste0("enero 2019-",mesactual," ",año)
periodo2=paste0(mespal[mes]," ",año)

#Titulos de ficha y gráficos
ti_ficha <- paste0("Indicadores de la industria manufacturera de Jalisco en ",mesactual," de ",año)
ti_his_pot <- paste0("Personal ocupado de la industria manufacturera en Jalisco, cifras mensuales, ",periodo1)
ti_rank_pot <- paste0("Variación porcentual anual del personal ocupado de la industria manufacturera por entidad federativa, ",periodo2)
ti_his_hor <- paste0("Horas trabajadas por el personal ocupado de la industria manufacturera en Jalisco, cifras mensuales en miles de horas, ","enero 2019-",periodo1)
ti_rank_hor <- paste0("Variación porcentual anual de las horas trabajadas por el personal ocupado de la industria manufacturera por entidad federativa, ",periodo2)
ti_his_vp <- paste0("Valor de la producción de la industria manufacturera en Jalisco, cifras mensuales en millones de pesos constantes, ","enero 2019-",periodo1)
ti_rank_vp <- paste0("Variación porcentual anual del valor de la producción de la industria manufacturera en términos reales por entidad federativa, ",periodo2)

#Fuente y titulos para gráficas
fue = "Fuente: IIEG, con información de INEGI. EMIM."

#Notas de gráficas
notas=data.frame(c(
  "Nota: El personal ocupado total corresponde a la suma del personal dependiente de la razón social, del personal suministrado por otra razón social y del personal no remunerado. El promedio se refiere al de los últimos doce meses.",
  "Nota: La variación anual es la variación con respecto al mismo mes del año anterior.",
  "Nota: Las horas trabajadas por el personal ocupado total son la suma de las horas normales y extraordinarias efectivamente trabajadas por el personal dependiente de la razón social, el personal suministrado por otra razón social y el personal no remunerado. Cifras en Miles de horas. El promedio se refiere al de los últimos doce meses.",
  "Nota: La variación anual es la variación con respecto al mismo mes del año anterior.",
  "Nota: El valor de la producción se encuentra en millones de pesos constantes deflactados al último periodo publicado de la EMIM mediante el índice nacional de precios al productor para actividades secundarias con petróleo. El promedio se refiere al de los últimos doce meses.",
  "Nota: La variación anual es la variación con respecto al mismo mes del año anterior. La variación es en términos reales."
))

#Nombres de variables
nombre_var=data.frame(c("Año","Mes","Valor","Promedio","Entidad","Variación"))
names(graf_his_pot)=nombre_var[1:4,1]
names(graf_rank_pot)=nombre_var[5:6,1]
names(graf_his_hor)=nombre_var[1:4,1]
names(graf_rank_hor)=nombre_var[5:6,1]
names(graf_his_vp)=nombre_var[1:4,1]
names(graf_rank_vp)=nombre_var[5:6,1]


#Datos para Gráficas de Excel
wb=createWorkbook("IIEG DIEEF")
addWorksheet(wb, "HIS POT")
titulo=paste0("Figura ##. ",ti_his_pot)
writeData(wb, sheet=1, titulo, startCol=1, startRow=1)
writeData(wb, sheet=1, fue, startCol=1, startRow=2)
writeData(wb, sheet=1, notas[1,1], startCol=1, startRow=3)
writeData(wb, sheet=1, graf_his_pot, startCol=1, startRow=5)

addWorksheet(wb, "RANK POT")
titulo=paste0("Figura ##. ",ti_rank_pot)
writeData(wb, sheet=2, titulo, startCol=1, startRow=1)
writeData(wb, sheet=2, fue, startCol=1, startRow=2)
writeData(wb, sheet=2, notas[2,1], startCol=1, startRow=3)
writeData(wb, sheet=2, graf_rank_pot, startCol=1, startRow=5)

addWorksheet(wb, "HIS HOR")
titulo=paste0("Figura ##. ",ti_his_hor)
writeData(wb, sheet=3, titulo, startCol=1, startRow=1)
writeData(wb, sheet=3, fue, startCol=1, startRow=2)
writeData(wb, sheet=3, notas[3,1], startCol=1, startRow=3)
writeData(wb, sheet=3, graf_his_hor, startCol=1, startRow=5)

addWorksheet(wb, "RANK HOR")
titulo=paste0("Figura ##. ",ti_rank_hor)
writeData(wb, sheet=4, titulo, startCol=1, startRow=1)
writeData(wb, sheet=4, fue, startCol=1, startRow=2)
writeData(wb, sheet=4, notas[4,1], startCol=1, startRow=3)
writeData(wb, sheet=4, graf_rank_hor, startCol=1, startRow=5)

addWorksheet(wb, "HIS VP")
titulo=paste0("Figura ##. ",ti_his_vp)
writeData(wb, sheet=5, titulo, startCol=1, startRow=1)
writeData(wb, sheet=5, fue, startCol=1, startRow=2)
writeData(wb, sheet=5, notas[5,1], startCol=1, startRow=3)
writeData(wb, sheet=5, graf_his_vp, startCol=1, startRow=5)

addWorksheet(wb, "RANK VP")
titulo=paste0("Figura ##. ",ti_rank_vp)
writeData(wb, sheet=6, titulo, startCol=1, startRow=1)
writeData(wb, sheet=6, fue, startCol=1, startRow=2)
writeData(wb, sheet=6, notas[6,1], startCol=1, startRow=3)
writeData(wb, sheet=6, graf_rank_vp, startCol=1, startRow=5)


#Texto
addWorksheet(wb,"TEXTO")
writeData(wb, sheet=7, "Título Ficha:", startCol = 1, startRow = 1)
writeData(wb, sheet=7, ti_ficha, startCol = 1, startRow = 2)
#
writeData(wb, sheet=7, "HIS POT", startCol = 1, startRow = 4)
writeData(wb, sheet=7, "Texto:", startCol = 1, startRow = 5)
writeData(wb, sheet=7, descripcion_1, startCol = 2, startRow = 5)
writeData(wb, sheet=7, "Gráfica:", startCol = 1, startRow = 6)
writeData(wb, sheet=7, ti_his_pot, startCol = 2, startRow = 6)
writeData(wb, sheet=7, "Fuente:", startCol = 1, startRow = 7)
writeData(wb, sheet=7, fue, startCol = 2, startRow = 7)
writeData(wb, sheet=7, "Nota:", startCol = 1, startRow = 8)
writeData(wb, sheet=7, notas[1,1], startCol = 2, startRow = 8)
#
writeData(wb, sheet=7, "RANK POT:", startCol = 1, startRow = 10)
writeData(wb, sheet=7, "Texto:", startCol = 1, startRow = 11)
writeData(wb, sheet=7, descripcion_2, startCol = 2, startRow = 11)
writeData(wb, sheet=7, "Gráfica:", startCol = 1, startRow = 12)
writeData(wb, sheet=7, ti_rank_pot, startCol = 2, startRow = 12)
writeData(wb, sheet=7, "Fuente:", startCol = 1, startRow = 13)
writeData(wb, sheet=7, fue, startCol = 2, startRow = 13)
writeData(wb, sheet=7, "Nota:", startCol = 1, startRow = 14)
writeData(wb, sheet=7, notas[2,1], startCol = 2, startRow = 14)
#
writeData(wb, sheet=7, "HIS HOR", startCol = 1, startRow = 16)
writeData(wb, sheet=7, "Texto:", startCol = 1, startRow = 17)
writeData(wb, sheet=7, descripcion_3, startCol = 2, startRow = 17)
writeData(wb, sheet=7, "Gráfica:", startCol = 1, startRow = 18)
writeData(wb, sheet=7, ti_his_hor, startCol = 2, startRow = 18)
writeData(wb, sheet=7, "Fuente:", startCol = 1, startRow = 19)
writeData(wb, sheet=7, fue, startCol = 2, startRow = 19)
writeData(wb, sheet=7, "Nota:", startCol = 1, startRow = 20)
writeData(wb, sheet=7, notas[3,1], startCol = 2, startRow = 20)
#
writeData(wb, sheet=7, "RANK HOR:", startCol = 1, startRow = 22)
writeData(wb, sheet=7, "Texto:", startCol = 1, startRow = 23)
writeData(wb, sheet=7, descripcion_4, startCol = 2, startRow = 23)
writeData(wb, sheet=7, "Gráfica:", startCol = 1, startRow = 24)
writeData(wb, sheet=7, ti_rank_hor, startCol = 2, startRow = 24)
writeData(wb, sheet=7, "Fuente:", startCol = 1, startRow = 25)
writeData(wb, sheet=7, fue, startCol = 2, startRow = 25)
writeData(wb, sheet=7, "Nota:", startCol = 1, startRow = 26)
writeData(wb, sheet=7, notas[4,1], startCol = 2, startRow = 26)
#
writeData(wb, sheet=7, "HIS VP", startCol = 1, startRow = 28)
writeData(wb, sheet=7, "Texto:", startCol = 1, startRow = 29)
writeData(wb, sheet=7, descripcion_5, startCol = 2, startRow = 29)
writeData(wb, sheet=7, "Gráfica:", startCol = 1, startRow = 30)
writeData(wb, sheet=7, ti_his_vp, startCol = 2, startRow = 30)
writeData(wb, sheet=7, "Fuente:", startCol = 1, startRow = 31)
writeData(wb, sheet=7, fue, startCol = 2, startRow = 31)
writeData(wb, sheet=7, "Nota:", startCol = 1, startRow = 32)
writeData(wb, sheet=7, notas[5,1], startCol = 2, startRow = 32)
#
writeData(wb, sheet=7, "RANK VP:", startCol = 1, startRow = 34)
writeData(wb, sheet=7, "Texto:", startCol = 1, startRow = 35)
writeData(wb, sheet=7, descripcion_6, startCol = 2, startRow = 35)
writeData(wb, sheet=7, "Gráfica:", startCol = 1, startRow = 36)
writeData(wb, sheet=7, ti_rank_vp, startCol = 2, startRow = 36)
writeData(wb, sheet=7, "Fuente:", startCol = 1, startRow = 37)
writeData(wb, sheet=7, fue, startCol = 2, startRow = 37)
writeData(wb, sheet=7, "Nota:", startCol = 1, startRow = 38)
writeData(wb, sheet=7, notas[6,1], startCol = 2, startRow = 38)
#

#Guardar lo hecho en R a un archivo de Excel
nombre_wb=paste0("EMIM_R-Excel_",fxls,".xlsx")
saveWorkbook(wb, nombre_wb, overwrite = TRUE)
