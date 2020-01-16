# librerias ---------------------------------------------------------------
library(dplyr)
library(tidyverse)
library(readxl)
library(stringi)
library(stringr)
library(countrycode)
library(openxlsx)
library(janitor)
# variables ---------------------------------------------------------------
yeari <- "2007"
yearf <- "2019"
month <- "dic"
# base_nombre 
base_nombre <- paste0("input/","base_pais_sector_dic2019_v.xlsx")
base_resultados <- paste0("input/", "GIIED_dic2019_v.xlsx")
# tolerancia
tol <- 0.001 # para los paises
tolb <- 0.1001 # para "otros" paises en la base
# cambiar el numero de decimales
options(digits = 12)
# cargar bases por sector -------------------------------------------------
base_empresarial <- list()
# funcion para cargar base
load_empresarial <- function(base, y) {
  suppressMessages({
    x <- read_excel(paste(base, sep = "/"), # cargar base
      sheet = paste("Empresarial", y) # seleccion de hoja por year
    )
  })
  vars <- x[1, ] %>% # selecionar primera fila
    as.character() %>% # convertir a caracteres
    tolower() %>% # convertir en minisculuas
    stri_trans_general(., "latin-ascii") # quitar acentos
  x <- x %>%
    set_names(vars) %>%
    slice(-1)
}
# loop para cargar las bases
base_empresarial <- lapply(yeari:yearf, function(i)
  load_empresarial(base_nombre, i))
# poner nombres como año
names(base_empresarial) <- yeari:yearf
# comparar nombres respecto a 2007
varsbase <- colnames(base_empresarial[["2007"]]) %>%
  as.character() %>%
  tolower() %>%
  stri_trans_general(., "latin-ascii") %>%
  unlist()
diff_nombres <- function(base, y) {
  vary <- base_empresarial[[paste(y)]] %>%
    colnames() %>%
    as.character() %>% # convertir a caracteres
    tolower() %>%
    stri_trans_general(., "latin-ascii") %>%
    unlist()
  {
    if (length(base) > length(vary)) {
      nvars <- which(!(base %in% vary))
      vary <- base
      vary[nvars] <- rep(" ", length(nvars))
      difnombres <- data.frame(base = base, var = vary)
    }
    else if (length(base) == length(vary)) {
      difnombres <- data.frame(vbase = base, var = vary[match(base, vary)])
    }
    else {
      nvars <- which(!(vary %in% base))
      base[nvars] <- rep(" ", length(nvars))
      difnombres <- data.frame(vbase = base, var = vary[order(match(vary, base))])
    }
  }
  difnombres <- difnombres %>%
    mutate_if(is.factor, as.character) %>%
    mutate(condicion = (base == var)) %>%
    mutate(year = y) %>%
    mutate(condicion = case_when(
      is.na(condicion) ~ FALSE,
      TRUE ~ condicion
    )) %>%
    filter(condicion == FALSE)
  difnombres
}
# lista de nombres malos --------------------------------------------------
badsector <- list()
badsector <- lapply(yeari:yearf, function(i) diff_nombres(varsbase, i)) %>%
  bind_rows()
# este muestra los errores en sectores si hay
if (nrow(badsector) != 0) {
  View(badsector)
}
# !!!!!!!!!!!corregir nombre de sectores ----------------------------------
# Error 1. Error en 2010
colnames(base_empresarial[["2010"]])[1] <- "pais"
# Error 2. Agregar columnas a las bases
# funcion
agg_columns <- function(base, newnames, y) {
  base <- base %>% mutate(year = y)
  x <- base
  xnames <- names(x)
  xnamesn <- c(xnames, newnames)
  xnamesn <- unique(xnamesn)
  dif <- setdiff(xnamesn, xnames)
  if (length(dif) > 0) {
    base <- base %>% mutate(!!dif := "0")
    base
  } else {
    base
  }
}
# datos
base_empresarial <- lapply(yeari:yearf, function(i)
  agg_columns(base_empresarial[[paste0(i)]], "otros", i))
# poner nombres como año
names(base_empresarial) <- yeari:yearf
# modifcacion de variables ------------------------------------------------
# cargar nombres de paises del Banco Mundial
# https://wits.worldbank.org/wits/wits/witshelp-es/Content/Codes/Country_Codes.htm
countries_names <- read_excel("auxi/countries_names.xlsx")
# la base de countries sin repetidos (para pegar por iso)
cnames <- countries_names %>%
  select(-pais_clean) %>%
  distinct(iso, countries, .keep_all = TRUE)
# funcion para modificar valores de las variables
mod_variables <- function(base) {
  x <- base
  x <- x %>% mutate(pais = tolower(pais))
  x <- x %>% mutate(pais = stri_trans_general(pais, "latin-ascii"))
  xp <- which(x$pais %in% c("total", "total general"))
  if (!is_empty(xp)) {
    x <- x %>% slice(xp * (-1))
  } else {
    x <- x
  }
  x <- x %>%
    mutate(pais_clean = pais) %>%
    mutate(pais_clean = gsub("^argentina$", "argentina", pais)) %>%
    mutate(pais_clean = gsub("^corea$", "corea del sur", pais_clean)) %>%
    mutate(pais_clean = gsub("^islas Vigenes$", "islas Virgenes", pais_clean)) %>%
    mutate(pais_clean = gsub("^portuga$", "portugal", pais_clean)) %>%
    mutate(pais_clean = gsub("^tuquia$", "turquia", pais_clean)) %>%
    mutate(pais_clean = case_when(
      pais_clean == "otros paises" ~ "otros",
      TRUE ~ pais_clean
    ))
  suppressMessages({
  x <- x %>%
    left_join(countries_names)
  })
  y <- x %>% # lista de paises sin match
    slice(which(is.na(iso))) %>%
    select(pais_clean, countries, iso, matches("year"))

  list(base = x, nonames = y)
}
badcountries <- lapply(base_empresarial, function(i) mod_variables(i)[["nonames"]]) %>%
  bind_rows() %>%
  distinct(pais_clean, .keep_all = T)
# Mostrar los paises con nombres mal
if (nrow(badcountries) != 0) {
  View(badcountries)
}
base_empresarial <- lapply(base_empresarial, function(i) mod_variables(i)[["base"]])
# !!!!!!!!!!!coregir nombres de paises ------------------------------------
# Para set no hay error en nombres más que los especificados
# anteriormente, es decir hay una categoria otros que hay que mantener para
# los no clasificados, y usarlo en inmobiliario
# unir bases --------------------------------------------------------------
bempresarial <- bind_rows(base_empresarial) %>%
  select(countries, iso, year, varsbase, -pais) %>%
  pivot_longer(
    cols = varsbase[which(!(varsbase %in% "pais"))],
    names_to = "sector", values_to = "valor"
  ) %>%
  mutate(valor = as.numeric(valor)) %>%
  filter(sector != "total general")
# cargar base INMOBILIARIO ------------------------------------------------
suppressMessages({
  inmo <- read_excel(paste(base_nombre, sep = "/"), # cargar base
                     sheet = paste("Inmobiliaria")
  )
})
# transformar base inmo ---------------------------------------------------
inmo <- inmo %>%
  filter(complete.cases(.)) %>%
  set_names("pais", paste(yeari:yearf))
inmo <- mod_variables(inmo)
badcountries2 <- inmo[["nonames"]] %>%
  distinct(pais_clean, .keep_all = T)
inmo <- inmo[["base"]]
# !!!!!!!!!! corregir nombres de paies en inmobiliarios -------------------
if (nrow(badcountries2) != 0) {
  View(badcountries2)
}
# # Error 1. corregir otros paises por otros para igualarlo a otros de empresarial
# inmo <- inmo %>%
#   mutate(pais = case_when(
#     pais_clean == "otros paises" ~ "otros",
#     TRUE ~ pais_clean
#   ))
# unir bases --------------------------------------------------------------
# homogeneizar base inmo
inmo <- inmo %>%
  pivot_longer(
    cols = paste(yeari:yearf),
    names_to = "year",
    values_to = "valor"
  ) %>%
  mutate(sector = "inmobiliaria") %>%
  select(countries, iso, year, sector, valor) %>%
  mutate(year = as.numeric(year)) %>%
  mutate(valor = valor * 1000000)
# unir base
ied_final <- bind_rows(bempresarial, inmo)
# funcion para agregar paises que faltan de asignar valor
verif <- function(base, corregir = 0) {
  cn <- ied_final %>%
    select(countries) %>%
    distinct(., .keep_all = TRUE)
  y <- ied_final %>%
    select(year) %>%
    distinct(., .keep_all = TRUE)
  s <- ied_final %>%
    select(sector) %>%
    distinct(., .keep_all = TRUE)
  vf <- data.frame(conutries <- cn)
  vf[, as.character(unlist(y))] <- 0
  vf <- vf %>%
    pivot_longer(
      cols = -countries,
      names_to = "year",
      values_to = "valor"
    )
  vf[, as.character(unlist(s))] <- 0
  # este evita la duplicacion por pais clean
  suppressMessages({
   vf <- vf %>%
    select(-valor) %>%
    pivot_longer(
    cols = -c(countries, year),
    names_to = "sector",
    values_to = "valor"
    ) %>%
    select(-valor) %>%
    mutate(year = as.numeric(year)) %>%
    left_join(ied_final) %>%
    select(-iso) %>%
    left_join(cnames) %>%
    select(year, iso, countries, sector, valor) %>%
    mutate(valor = if_else(is.na(valor), corregir, valor))
   vf
  })
}
ied_final <- verif(ied_final)
# resultados --------------------------------------------------------------
# leer las hojas del excel
hojas <- excel_sheets(path = paste(base_resultados, sep = "/"))
# limpar datos
hojas_clean <- hojas %>%
  tolower() %>% # convertir en minisculuas
  stri_trans_general(., "latin-ascii") %>% # quitar acentos
  sub(" ", "_", .) %>% # cambiar espacios con guion bajo
  sub("-", "_", .) # cambiar guion por guion bajo
# guardar excel con los nombres correctos
wb <- loadWorkbook(paste(base_resultados, sep = "/"),
  xlsxFile = NULL
)
names(wb) <- hojas_clean
saveWorkbook(wb, paste(path,
  paste0("names", base_resultados),
  sep = "/"
), overwrite = TRUE)
# funcion para cargar base
load_result <- function(sheets) {
  base_resultados_names <- paste(path,
    paste0("names", base_resultados),
    sep = "/"
  )
  t1 <- sheets[1] %in% hojas_clean # verificador hoja 1
  t2 <- sheets[2] %in% hojas_clean # verificador hoja 2
  if (!t1) stop("La hoja uno no se encuentra en el excel")
  if (!t1) stop("La hoja dos no se encuentra en el excel")
  suppressMessages({
    x0 <- read_excel(base_resultados_names,
      sheet = paste(sheets[1])
    )

    x1 <- read_excel(base_resultados_names,
      sheet = paste(sheets[2])
    ) %>%
      filter(complete.cases(.))
  })
  list(pd_ed_pais = x0, pd_ed_sector = x1)
}
verhojas <- c("pd_ed_pais", "pd_ed_sector")
hj <- load_result(verhojas)
# #### resumen por pais
ied_r1 <- ied_final %>%
  select(year, iso, sector, valor) %>%
  group_by(year, iso) %>%
  summarize(valor1 = sum(valor))
# Obtener base homogeneizada del resumen por pais
r_x_p <- function(base, colref = "PAÍS") {
 suppressMessages({
  x0 <- base
  lim0 <- which(x0[, 1] == colref)
  suppressWarnings({
    x.0.0 <- x0 %>%
      slice(-(1:(lim0 - 1))) %>%
      row_to_names(1) %>%
      clean_names()
  })
  x.0.0 <- x.0.0 %>%
    select("pais", paste0("x", yeari:yearf)) %>%
    filter(complete.cases(.)) %>%
    filter(!(pais %in% c(
      "América Central",
      "América del Norte",
      "Otros relevantes de América",
      "Relevantes de Europa",
      "Relevantes de Asia",
      "TOTAL"
    )))
  x.0.1 <- mod_variables(x.0.0)
  nonames <- data.frame(x.0.1["nonames"])
  if (nrow(nonames) != 0) {
    View(nonames)
    stop("Un nombre de pais no tiene contraparte ISO numerico")
  } else {
    x.0.1 <- x.0.1[["base"]]
  }
  x.0.2 <- x.0.1 %>%
    pivot_longer(
      cols = starts_with("x"), names_to = "year",
      values_to = "valor2"
    ) %>%
    mutate(year = as.numeric(sub("x", "", year))) %>%
    mutate(valor2 = valor2 * 1000000) %>%
    select(year, iso, valor2)
  
  x.0.3 <- x.0.2 %>% 
    left_join(ied_r1, by = c("year", "iso"))
  
  if (any(is.na(x.0.3$valor1))) {
    sl <- which(is.na(x.0.3$valor1))
    View(x.0.3[sl, ])
    stop("No existe un valor asignado en la IED final para ese pais (1)")
  }
  if (any(is.na(x.0.3$valor2))) {
    sl <- which(is.na(x.0.3$valor2))
    View(x.0.3[sl, ])
    stop("No existe un valor asignado en la base de resultados para ese pais (1)")
  }
  x.0.4 <- x.0.2 %>% 
    mutate(iso = case_when(
      iso == "OTR" ~ "OTB", # otros paises del cuadro resumen es más agregado 
      TRUE ~ iso,  # que los que se muestran en ied_final
    ))
  x.0.5 <- x.0.4 %>% 
           select(year, iso) %>%
           mutate(ver = 1)
  x.0.6 <- ied_r1 %>% 
           left_join(x.0.5, by = c("year", "iso"))

  x.0.7 <- x.0.6 %>% 
           mutate(iso2 = case_when(
             is.na(ver) ~ "OTB",
             TRUE ~ iso
           )) %>% 
           select(year, iso, iso2)
           #        valor1) %>% 
           # na.omit %>% 
           # select(year, iso, iso2)
  x.0.8 <- x.0.7 %>% 
           left_join(ied_r1) %>% 
           left_join(x.0.2) %>% 
           group_by(year, iso2) %>% 
           summarise(valor1 = sum(valor1, na.rm = TRUE),
                     valor2 = sum(valor2, na.rm = TRUE)) %>% 
           select(year, iso2, valor1, valor2) %>% 
           rename(iso = iso2)
  x.0.8
 })
}
t_x_p <- hj[["pd_ed_pais"]]
t_x_p <- r_x_p(t_x_p)
# si da error hay que agregar algun pais
# funcion de tolerancia
tolerancia <- function(base, type = "iso", x = NULL, y = NULL, tol1 = tol, tol2 = NULL, round_digits = 0) {
  x <- round(base[, x], round_digits)
  y <- round(base[, y], round_digits)
  z <- abs(x - y)
  z0 <- !(z <= tol1)
  if (!(is.null(tol2))) {
  z1 <- !(z <= tol2)
  xy <- base[, "iso"]
  z <- ifelse(xy == "OTB", z1,z0)
  } else {
  z <- z0
  }
  if (type == "iso"){
  suppressMessages({
      error <- slice(base, which(z)) %>%
        left_join(cnames) %>%
        distinct(iso, year, .keep_all = TRUE)  
  })
  } else {
      error <- slice(base, which(z))
  }

  if (nrow(error) != 0) {
    View(error)
    if (type == "iso"){
      stop("Existe un problema en los datos por pais")
    } else {
      stop("Existe un problema en los datos por sector")
    }
  } else {
    if(type == "iso") {
      print("No existe problema con los datos por pais")
    } else {
      print("No existe problema con los datos por sector")
    }
  }
}
# comprobar tolerancia por pais
tolerancia(t_x_p, type = "iso", x = "valor2", y = "valor1", tol2 = tolb)
# ### resumen por sector
ied_r2 <- ied_final %>%
  select(year, iso, sector, valor) %>%
  group_by(year, sector) %>%
  summarize(valor1 = sum(valor))
# Obtener base homogeneizada del resumen por sector
r_x_s <- function(base, y = yearf){
  colyears <- (ncol(base) - 1) 
  yearsnms <- (colyears / 5)
  yearirps <- (as.numeric(yearf) - (yearsnms - 1))
  trimsnms <- lapply(yearirps:yearf, function(i) lapply(1:5, function(j) paste0(i,"_", j))) %>% 
              unlist()
  colnmsps <- c("sector", trimsnms)
  colnames(base) <- colnmsps
  x        <- base %>% 
           mutate_all(as.character) %>% 
           pivot_longer(cols = trimsnms,
                       names_to = "date", values_to = "valor2") %>% 
           mutate(year = substr(date, 1, 4),
                  quat = substr(date, 6, 6),
                  valor2 = (as.numeric(valor2)) * 1000000) %>% 
           mutate(year = as.numeric(year)) %>% 
           filter(quat != 5) %>% 
           group_by(year, sector) %>% 
           summarize(valor2 = sum(valor2)) %>% 
           select(year, sector, valor2) %>% 
           mutate(sector = tolower(sector)) %>% 
           mutate(sector = stri_trans_general(sector, "latin-ascii")) %>% 
           mutate(sector = gsub("^otros$", "otros", sector)) %>% 
           mutate(sector = str_replace(sector, "industria manufacturera", "industria")) %>% 
           mutate(sector = gsub(".*turismo.*", "turistico", sector))
    
    
  xp <- which(x$sector %in% c("total", "total general"))
  if (!is_empty(xp)) {
    x <- x %>% slice(xp * (-1))
  } else {
    x <- x
  }
  suppressMessages({
  x <- x %>% 
       left_join(ied_r2) %>% 
       mutate(valor1 = if_else(is.na(valor1), 0, valor1)) %>% 
       filter(year >= yeari)
  })
}
t_x_s <- hj[["pd_ed_sector"]]
t_x_s <- r_x_s(t_x_s)
# comprobar tolerancia por sector
tolerancia(base = t_x_s, type = "sector", x = "valor2", y = "valor1", tol1 = 0.01)
# exportar excel ----------------------------------------------------------
suppressMessages({
# correcion de tildes y mayusculas por sector
ied_final <- ied_final %>% 
             mutate(sector = str_replace(sector, "turistico", "Turístico")) %>%
             mutate(sector = str_to_title(sector))
# cargar acuerdos
cod_19032019_a <- read_excel("ied/cod_19032019.xlsx", 
                           sheet = "Acuerdo")
names(cod_19032019_a) <- c("number", "pais", "Acuerdo")
cod_19032019_a <- cod_19032019_a %>% 
                  mutate(number = as.numeric(number)) %>% 
                  left_join(cnames) %>% 
                  select(iso, Acuerdo)
# cargar bloque
cod_19032019_b <- read_excel("ied/cod_19032019.xlsx", 
                           sheet = "Bloque")
names(cod_19032019_b) <- c("number", "pais", "Bloque")
cod_19032019_b <- cod_19032019_b %>% 
                  mutate(number = as.numeric(number)) %>% 
                  left_join(cnames) %>% 
                  select(iso, Bloque, number)
# pegar acuerdos y bloques
ied_final2 <- ied_final %>% 
             left_join(cod_19032019_a) %>% 
             left_join(cod_19032019_b) %>% 
             select(year, number, iso, countries, Acuerdo, Bloque, sector, valor) %>% 
             mutate(valormiles = valor / 1000000) %>% 
             rename("Año" = "year", "CódigoPaís" = "number", "País" = "countries",
                    "Sector" = "sector", "Valor US$" = "valor", "Valor US$ mil" = "valormiles")
})
         
write.xlsx(ied_final2, sheetName = "Datos",
           paste(path, paste0(paste("ied", month, yearf, sep = "_"), ".xlsx"), sep = "/"))

wb <- loadWorkbook(paste(path, paste0(paste("ied", month, yearf, sep = "_"), ".xlsx"), sep = "/"))
addWorksheet(wb, "td_acuerdo")
addWorksheet(wb, "td_sector")                   

saveWorkbook(wb, file = paste(path, paste0(paste("ied", month, yearf, sep = "_"), ".xlsx"), sep = "/"), overwrite = TRUE)










