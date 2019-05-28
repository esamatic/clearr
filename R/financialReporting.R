library(readxl)
library(tabulizer)
library(tidyverse)

#' Avaa HelpostiLaskun www-sivu
#'
#' @return
#' @export
#'
#' @examples
openHelpostilasku <- function() {
  browseURL("https://ssl.helpostilasku.mobi/App/Login")
}

#' Parsi laskutetut laskut
#'
#' @param filename pdf-tiedoston nimi
#'
#' @return Laskut tibble-muodossa
#' @export
#'
#' @examples
parseLaskutetutLaskut <- function(filename = file.choose()) {
  out <- 
    extract_tables(file = filename, encoding = "UTF-8") %>%
    do.call(rbind, .) %>%
    as_tibble() %>%
    select(pvm = V1, laskuno = V2, asiakas = V3, eur = V4) %>%
    mutate(pvm = as.Date(pvm, format = "%d.%m.%Y")) %>%
    mutate(eur = as.numeric(eur)) %>%
    mutate(viitenro = NA)
  
  ## laske viitenumero
  mult <- paste0(rep("137", 100), collapse = "")
  for (i in 1:nrow(out)) {
    laskuno <- out[[i, "laskuno"]] 
    kerroin <- substr(mult, nchar(mult)-nchar(laskuno)+1, nchar(mult))
    x       <- strsplit(kerroin, "") %>% unlist() %>% as.integer()
    y       <- strsplit(laskuno, "") %>% unlist() %>% as.integer()
    tarkistussumma <- as.character((10 - (sum(x * y) %% 10)) %% 10)
    out[[i, "viitenro"]] <- paste0(laskuno, tarkistussumma)
  }
  return(out)
}

#' Parsi tuotteet pdf-tiedostosta
#'
#' @param filename pdf-tiedosto
#'
#' @return Tuotteet tibble-muodossa
#' @export
#'
#' @examples
parseTuotteet <- function(filename = file.choose()) {
 extract_tables(file = filename, encoding = "UTF-8")[[1]] %>%
    as_tibble() %>%
    select(tili = V1, nimi = V2) %>%
    mutate(tili = as.integer(tili))
}

#' Parsi myynti tuotteittain pdf-tiedostosta
#'
#' @param filename pdf-tiedoston nimi
#'
#' @return Myynti tuotteittain tibble-muodossa
#' @export
#'
#' @examples
parseMyyntiTuotteittain <- function(filename = file.choose()) {
  extract_areas(file = filename, encoding = "UTF-8")[[1]] %>%
    as_tibble() %>%
    select(nimi = V1, kpl = V2, eur = V3) %>%
    filter(!(nimi %in% c("Yhteensä EUR", "Tuote"))) %>%
    mutate(kpl = gsub("[^0-9]", "", kpl) %>% as.integer()) %>%
    mutate(eur = as.numeric(eur)) %>%
    group_by(nimi) %>%
    summarize(kpl = sum(kpl), eur = sum(eur))
}

#' Parsi budjetti
#'
#' @param filename Excel-tiedosto
#'
#' @return Data frame
#' @export
#'
#' @examples
parseBudjetti <- function(filename = file.choose()) {
  tulot <- 
    readxl::read_excel(path = filename, sheet = "2018-2019", range = "A7:D30", col_names = FALSE) %>%
    mutate(tyyppi = "tulo")
  menot <- 
    readxl::read_excel(path = filename, sheet = "2018-2019", range = "A36:D65", col_names = FALSE) %>%
    mutate(tyyppi = "meno")
  bind_rows(tulot, menot) %>%
    setNames(c("tili", "nimi", "eur_budjetti", "ed_kausi", "tyyppi")) %>%
    select(tili, nimi, eur_budjetti, tyyppi) %>%
    mutate(tili = as.integer(tili)) %>%
    mutate(eur_budjetti = as.numeric(eur_budjetti))
}

#' Tee seuranta
#'
#' @param filename Tiedosto, johon seuranta talletetaan
#'
#' @return Ei mitaan
#' @export
#'
#' @examples
teeSeuranta <- function(filename = "seuranta.xlsx") {
  budjetti <- parseBudjetti("data/Budjetti 2018-2019.xlsx")
  tilit    <- budjetti %>% select(nimi, tili)  
  tuotteet <- parseTuotteet("data/Tuotteet.pdf")
  myynti_tuotteittain <- parseMyyntiTuotteittain("data/Myynti tuotteittain.pdf")
  laskutetut_laskut <- parseLaskutetutLaskut("data/Laskutetut laskut.pdf")
  helpostilasku <- 
    full_join(tuotteet, myynti_tuotteittain, by = "nimi") %>%
    mutate(eur = replace(eur, is.na(eur), 0)) %>%
    mutate(kpl = replace(kpl, is.na(kpl), 0))
  danske <- 
    loadDanske(muoto = "tilit") %>%
    left_join(tilit, by = "tili") %>%
    mutate(kpl = 1)
  kaikki <- 
    bind_rows(helpostilasku, danske) %>%
    group_by(tili, nimi) %>%
    summarize(eur = sum(eur), kpl = sum(kpl))
  seuranta <- 
    kaikki %>%
    full_join(budjetti, by = c("nimi", "tili")) %>%
    mutate(eur = replace(eur, is.na(eur), 0)) %>%
    mutate(kpl = replace(kpl, is.na(kpl), 0)) %>%
    filter(!(eur == 0 & eur_budjetti == 0)) %>%
    mutate(erotus = eur_budjetti-abs(eur)) %>%
    mutate(prosenttia_budjetoidusta = round(100*abs(eur)/eur_budjetti, 0))
  if (file.exists(filename)) file.remove(filename)
  wb <- XLConnect::loadWorkbook(filename = filename, create = T)
  XLConnect::createSheet(wb, name = "tulot")
  XLConnect::createSheet(wb, name = "menot")
  XLConnect::createSheet(wb, name = "yhteenveto")
  filter(seuranta, tyyppi == "tulo") %>%
    select(-tyyppi) %>%
    arrange(tili) %>%
    XLConnect::writeWorksheet(wb, data = ., sheet = "tulot")
  XLConnect::setColumnWidth(wb, sheet = "tulot", column = 1:7, width = -1)
  filter(seuranta, tyyppi == "meno") %>%
    select(-tyyppi) %>%
    arrange(tili) %>%
    XLConnect::writeWorksheet(wb, data = ., sheet = "menot")
  XLConnect::setColumnWidth(wb, sheet = "menot", column = 1:7, width = -1)
  yhteenveto <- matrix(NA, 3, 3)
  yhteenveto[1, 1] <- seuranta %>% filter(tyyppi == "tulo") %>% pull(eur_budjetti) %>% sum
  yhteenveto[2, 1] <- seuranta %>% filter(tyyppi == "meno") %>% pull(eur_budjetti) %>% sum
  
  yhteenveto[1, 2] <- seuranta %>% filter(tyyppi == "tulo") %>% pull(eur) %>% sum
  yhteenveto[2, 2] <- seuranta %>% filter(tyyppi == "meno") %>% pull(eur) %>% abs %>% sum

  yhteenveto[1, 3] <- yhteenveto[1, 1] - yhteenveto[1, 2]
  yhteenveto[2, 3] <- yhteenveto[2, 1] - yhteenveto[2, 2]
  
  yhteenveto[3, 1] <- yhteenveto[1, 1] - yhteenveto[2, 1]
  yhteenveto[3, 2] <- yhteenveto[1, 2] - yhteenveto[2, 2]
  
  
  yhteenveto %>% 
    as.data.frame() %>%
    setNames(c("budjetoitu", "toteutunut", "erotus")) %>%
    bind_cols(data.frame("."=c("tulot", "menot", "erotus")), .) %>%
    XLConnect::writeWorksheet(wb, data = ., sheet = "yhteenveto")
  XLConnect::setColumnWidth(wb, sheet = "yhteenveto", column = 1:4, width = -1)
  
  XLConnect::saveWorkbook(wb)
}


#' Danske Bankin datan tiedostonimi
#'
#' @return
#' @export
#'
#' @examples
danskeFilename <- function() { "danske.xlsx" }

#' Laske erityyppisetn tapahtumien rivimaarat
#'
#' @param x 
#'
#' @return
#' @export
#'
#' @examples
laskeMaarat <- function(x) {
  if (is.data.frame(x)) {
    return(nrow(x))
  } else {
    return(unlist(sapply(x, laskeMaarat)))
  }
}

#' Lisaa Danske Bankin tilitapahtumat kirjanpitoon
#'
#' @param filenames Tiedostojen (csv) nimet
#'
#' @return Ei mitaan
#' @export
#'
#' @examples
lisaaDanskeTilitapahtumat <- function(filenames = choose.files()) {
  danske_tapahtumat <- danskeFilename()
  if (file.exists(danske_tapahtumat)) {
    cat("Luetaan aiemmat tapahtumat tiedostosta", danske_tapahtumat, "\n")
    vanhat_tapahtumat <- loadDanske(danske_tapahtumat, muoto = "raw")
  } else { 
    vanhat_tapahtumat <- NULL
  }
  paivitetyt_tapahtumat <- vanhat_tapahtumat
  if (length(filenames)) {
    uudet_tapahtumat <- parseDanskeRaw(filenames)
    if (is.null(vanhat_tapahtumat)) {
      paivitetyt_tapahtumat <- uudet_tapahtumat
      lisatyt_yhteensa <- sum(laskeMaarat(paivitetyt_tapahtumat))
    } else {
      lisatyt_yhteensa <- 0
      for (tyyppi in names(uudet_tapahtumat)) {
        if (!exists(tyyppi, where = vanhat_tapahtumat)) {
          paivitetyt_tapahtumat[[tyyppi]] <- list() 
        }
        for (alatyyppi in names(uudet_tapahtumat[[tyyppi]])) {
          cat(tyyppi, "/", alatyyppi, "\n")
          uusi_data <- uudet_tapahtumat[[tyyppi]][[alatyyppi]]
          if (nrow(uusi_data)) {
            if (!exists(alatyyppi, where = paivitetyt_tapahtumat[[tyyppi]])) {
              paivitetyt_tapahtumat[[tyyppi]][[alatyyppi]] <- uusi_data 
              cat(paste0("Lisatty ", nrow(uusi_data), " rivia (",  tyyppi, "/", alatyyppi, "\n"))
              lisatyt_yhteensa <- lisatyt_yhteensa + nrow(uusi_data)
            } else {
              comp1 <- uusi_data %>% select(-ktili)
              comp2 <- vanhat_tapahtumat[[tyyppi]][[alatyyppi]] %>% select(-ktili)
              lisaa_rivit <- c()
              for (i in 1:nrow(comp1)) {
                if (nrow(setdiff(comp1[i, ], comp2))) {
                  lisaa_rivit <- c(lisaa_rivit, i)
                }
              }
              if (length(lisaa_rivit)) {
                paivitetyt_tapahtumat[[tyyppi]][[alatyyppi]] <-
                  bind_rows(vanhat_tapahtumat[[tyyppi]][[alatyyppi]],
                            uusi_data[lisaa_rivit, ])
                cat(paste0("Lisatty ", length(lisaa_rivit), " rivia (",  tyyppi, "/", alatyyppi, ")\n"))
                lisatyt_yhteensa <- lisatyt_yhteensa + length(lisaa_rivit)
              }
            }
          }
        }
      }
    }
    cat("Lisatty yhteensa", lisatyt_yhteensa, "rivia\n")
    cat("Tallennus tiedostoon", danske_tapahtumat, "\n")
    saveDanske(paivitetyt_tapahtumat)
  }
  return(invisible(NULL))
}

#' Parsi Dansken tilitiedot (csv)
#'
#' @param filenames Tiedostojen nimet
#'
#' @return Data (listamuodossa)
#' @export
#'
#' @examples
parseDanskeRaw <- function(filenames = choose.files()) {
  all_data <- NULL
  if (length(filenames)) {
    for (filename in filenames) {
      all_data <-
        read.csv2(file = filename, stringsAsFactors = FALSE) %>% 
        setNames(make.names(colnames(.))) %>%
        bind_rows(all_data)
    }
  }
  all_data <- 
    all_data %>%
    mutate(viitenro = stringr::str_extract(Viite.Maksaja.tai.saaja, "^[0-9]+")) %>%
    mutate(viitenro = stringr::str_replace(viitenro, "^[0]*", "")) %>%
    mutate(viitenro = replace(viitenro, is.na(viitenro), "puuttuu")) %>%
    mutate(tyyppi = strsplit(Viite.Maksaja.tai.saaja, " ") %>%
                             sapply(function(x) x[1]) %>% 
                               gsub("^[RF]*[0-9]+", "VIITENRO", .)) %>%
    mutate(arvo = ifelse(sign(Määrä.EUR) == -1, "maksu", "talletus"))
  yhdistelmat <-
    all_data %>%
    distinct(tyyppi, arvo)
  out <- list()
  for (t in unique(yhdistelmat$tyyppi)) {
    out[[t]] <- list()
    for (a in filter(yhdistelmat, tyyppi == t) %>% pull(arvo)) {
      out[[t]][[a]] <- 
        all_data %>%
        filter(tyyppi == t & arvo == a)
    }
  }
  out$VIITENRO$talletus <- 
    out$VIITENRO$talletus$Viite.Maksaja.tai.saaja %>%
    stringr::str_split("[ ]+Maksajan") %>% 
    sapply(function(x) x[1]) %>% 
    stringr::str_replace("[0]*", "") %>%
    str_match("([0-9]+)[ ]+([^ ].+)") %>%
    as.data.frame(stringsAsFactors = FALSE) %>%
    select(nimi = V3, viitenro = V2) %>%
    bind_cols(out$VIITENRO$talletus)
  out$VIITENRO$maksu <- 
    out$VIITENRO$maksu$Viite.Maksaja.tai.saaja %>%
    stringr::str_match("([0-9]+)[ ]+(.+) (FI[0-9]+)") %>%
    as.data.frame(stringsAsFactors = FALSE) %>%
    select(nimi = V3, viitenro = V2, tili = V4) %>%
    bind_cols(out$VIITENRO$maksu)
 
  if (F) { ## ota kayttoon myohemmin, parsii viestin, helpottaa tiliointia
  out$PANKKISIIRTO$talletus$viesti <- 
    sapply(strsplit(out$PANKKISIIRTO$talletus$Viite.Maksaja.tai.saaja, "Maksun.+tiedot:\\s*"), function(x) x[2]) %>%
    gsub("\\s+", " ", .) %>% 
    replace(., is.na(.), "ei viestia")
  
  out$SIIRTO$talletus$viesti <- 
    sapply(strsplit(out$SIIRTO$talletus$Viite.Maksaja.tai.saaja, "Maksun.+tiedot:\\s*"), function(x) x[2]) %>% 
    gsub("\\s+", " ", .) %>% 
    replace(., is.na(.), "ei viestia")
  
  out$SIIRTO$maksu$viesti <- 
    sapply(strsplit(out$SIIRTO$maksu$Viite.Maksaja.tai.saaja, "FI[0-9]+\\s*"), function(x) x[2]) %>% 
    gsub("\\s+", " ", .) %>% 
    replace(., is.na(.), "ei viestia")
  }
  for (t in unique(yhdistelmat$tyyppi)) {
    for (a in filter(yhdistelmat, tyyppi == t) %>% pull(arvo)) {
      out[[t]][[a]] <- 
        out[[t]][[a]] %>%
        mutate(ktili = NA)
    }
  }
  return(out)
}

#' Tallenna dansken tilitiedot
#'
#' @param x Data (listamuodossa)
#' @param filename Tiedosto, johon tallennetaan
#'
#' @return
#' @export
#'
#' @examples
saveDanske <- function(x, filename = danskeFilename()) {
  if (file.exists(filename)) file.remove(filename)
  wb <- XLConnect::loadWorkbook(filename = filename, create = T)
  for (tyyppi in names(x)) {
    for (alatyyppi in names(x[[tyyppi]])) {
      sheet_name <- paste0(tyyppi, "#", alatyyppi)
      XLConnect::createSheet(wb, name = sheet_name) 
      XLConnect::writeWorksheet(wb, x[[tyyppi]][[alatyyppi]], sheet = sheet_name)
    }
  }
  XLConnect::saveWorkbook(wb)
}

#' Lataa dansken aineisto Excelistä
#'
#' @param filename Tiedostonimi
#'
#' @return Data (listamuodossa)
#' @export
#'
#' @examples
loadDanske <- function(filename = danskeFilename(), muoto = c("raw", "tilit")) {
  if (!file.exists(filename)) {
    return(list())
  }
  muoto <- match.arg(muoto)
  wb <- XLConnect::loadWorkbook(filename)
  sheet_names <- XLConnect::getSheets(wb)
  if (muoto == "raw") {
    out <- list()
  } else {
    out <- NULL
  }
  for (sheet_name in sheet_names) {
    d <- XLConnect::readWorksheet(wb, sheet = sheet_name)
    if (muoto == "raw") {
      parsed <- str_split(sheet_name, "[#]", simplify = T) %>% c()
      if (!exists(parsed[1], where = out)) {
        out[[parsed[1]]] <- list()
      }
      out[[parsed[1]]][[parsed[2]]] <- d
    } else {
      out <- 
        d %>%
        select(eur = Määrä.EUR, tili = ktili) %>%
        filter(!is.na(tili)) %>%
        bind_rows(out)
    }
  }
  return(out)
}
