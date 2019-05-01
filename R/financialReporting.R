library(tabulizer)
library(tidyverse)

openHelpostilasku <- function() {
  browseURL("https://ssl.helpostilasku.mobi/App/Login")
}

parseHelpostiLasku <- function(filename = file.choose()) {
  out <- 
    extract_tables(file = filename, encoding = "UTF-8") %>%
    do.call(rbind, .) %>%
    as_tibble() %>%
    select(pvm = V1, laskuno = V2, asiakas = V3, eur = V4) %>%
    mutate(pvm = as.Date(pvm, format = "%d.%m.%Y")) %>%
    mutate(eur = as.numeric(eur)) %>%
    mutate(viitenro = NA)
  
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

parseDanske <- function(filenames = choose.files()) {
  all_data <- NULL
  if (length(filenames)) {
    for (filename in filenames) {
      all_data <- 
        read.csv2(file = filename, stringsAsFactors = FALSE) %>% 
        setNames(make.names(colnames(.))) %>%
        bind_rows(all_data)
    }
  }
  out <- list(siirto = list(),
              pankkisiirto = list(),
              viitesiirto = list())
  siirto <- 
    all_data %>%
    filter(stringr::str_detect(Viite.Maksaja.tai.saaja, "^SIIRTO"))
  out[["siirto"]][["talletus"]] <- filter(siirto, Määrä.EUR > 0)
  out[["siirto"]][["maksu"]] <- filter(siirto, Määrä.EUR < 0)
  pankkisiirto <- 
    all_data %>%
    filter(stringr::str_detect(Viite.Maksaja.tai.saaja, "^PANKKISIIRTO"))
  out[["pankkisiirto"]][["talletus"]] <- filter(pankkisiirto, Määrä.EUR > 0)
  out[["pankkisiirto"]][["maksu"]]    <- filter(pankkisiirto, Määrä.EUR < 0) 
  siirto <- 
    all_data %>%
    filter(stringr::str_detect(Viite.Maksaja.tai.saaja, "^SIIRTO"))
  out[["siirto"]][["talletus"]] <- filter(siirto, Määrä.EUR > 0)
  out[["siirto"]][["maksu"]]    <- filter(siirto, Määrä.EUR < 0)
  viitesiirto <- 
    all_data %>% filter(stringr::str_detect(Viite.Maksaja.tai.saaja, "^[0-9]"))
  out[["viitesiirto"]][["talletus"]] <- filter(viitesiirto, Määrä.EUR > 0)
  out[["viitesiirto"]][["maksu"]]    <- filter(viitesiirto, Määrä.EUR < 0) 
  
  return(out)  
}