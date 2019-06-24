library(tidyverse)

workflow <- function() {
  x <-
    parseKesayoRegistrations() %>%
    fixData
  openxlsx::write.xlsx(x, file = "ilmoittautumiset.xlsx")
  return(invisible(x))
}

parseKesayoRegistrations <- function(input_file = file.choose()) {
  x <-
    openxlsx::read.xlsx(xlsxFile = input_file, sheet = 1) %>%
    setNames(c("Aikaleima", "Etunimi", "Sukunimi", "Lempinimi", "Seura")) %>%
    dplyr::select(c("Etunimi", "Sukunimi", "Lempinimi", "Seura",
                    "1..PELILUOKKA", "1..PELILUOKASSA.nelinpeliparin.nimi",
                    "2..PELILUOKKA", "2..PELILUOKASSA.nelinpeliparin.nimi",
                    "3..PELILUOKKA", "3..PELILUOKASSA.nelinpeliparin.nimi",
                    "Terveiset,.risut.ja.ruusut.kilpailuj?rjest?jille.")) %>%
    dplyr::rename(Terveiset = "Terveiset,.risut.ja.ruusut.kilpailuj?rjest?jille.")

  person_cols <-
    dplyr::select(x, c("Etunimi", "Sukunimi", "Lempinimi", "Seura", "Terveiset"))
  all_data <- NULL
  Peliluokka <- Pari <- c()
  for (i in 1:3) {
    Peliluokka <- x[[paste0(i, "..PELILUOKKA")]]
    Pari <- x[[paste0(i, "..PELILUOKASSA.nelinpeliparin.nimi")]]
    class_cols <- data.frame(Peliluokka, Pari, stringsAsFactors = FALSE)
    all_data <- rbind(all_data, cbind(person_cols, class_cols))
  }
  return(all_data)
}

fixData <- function(x) {
  x <-
    x %>%
    dplyr::mutate(Seura = replace(Seura, is.na(Seura), "")) %>%
    dplyr::mutate(Seura = fixClub(Seura)) %>%
    dplyr::mutate(Seura = replace(Seura, Seura %in% c("ei", "-"), "")) %>%
    dplyr::mutate(Etunimi = ifelse(is.na(Lempinimi), Etunimi, paste0(Etunimi, ' "', Lempinimi, '"'))) %>%
    dplyr::mutate(Terveiset = replace(Terveiset, is.na(Terveiset), "")) %>%
    dplyr::mutate(Peliluokka = replace(Peliluokka, is.na(Peliluokka), "")) %>%
    dplyr::mutate(Peliluokka = fixClass(Peliluokka)) %>%
    dplyr::mutate(Pari = replace(Pari, is.na(Pari), "")) %>%
    dplyr::select(-Lempinimi) %>%
    dplyr::arrange(Peliluokka, Sukunimi, Etunimi) %>%
    dplyr::filter(!(Peliluokka == "" & Pari == ""))
  return(x)
}

fixClub <- function(x) {
  mapping <- list(
                  "Clear" = c("clear", "clera", "clear?"),
                  "ESB" = c("esb", "espoon sulkapallo badminton"),
                  "Drive" = c("drive", "sulkapalloseura drive"),
                  "Euran Veivi" = c("euran veivi"),
                  "Geneve Badminton Club" = c("geneve badminton club"),
                  "HBC"=c("hbc"),
                  "H?mSu" = c("h?msu", "h?meenlinnan sulkapalloilijat"),
                  "HalSu" = c("halsu", "halikon sulkis"),
                  "HanSu" = c("hansu", "bc hanhensulka"),
                  "HYSY" = c("hysy", "helsingin yliopiston sulkapalloyhdistys (hysy)"),
                  "Joen Sulka" = c("joen sulka"),
                  "Juankosken Pyrkiv?"=c("juankosken pyrkiv?"),
                  "KaaSu" = c("kaasu", "kaarinan sulka"),
                  "Karjalohjan Sulka" = c("karjalohjan sulka"),
                  "Klaki" = c("klaki"),
                  "Kosken Liikuntahalli" = c("kosken liikuntahalli"),
                  "Kyry" = c("kyry"),
                  "Lohjan Teho" = c("lohjan teho"),
                  "Loimaan Seudun Sulkis" = c("loimaan seudun sulkis"),
                  "NBC" = c("nbc"),
                  "?IF" = c("?if"),
                  "Onnen Urheilijat" = c("onnen urheiljat"),
                  "OSUKO" = c("osuko"),
                  "ParBa" = c("parba", "paraisten badminton"),
                  "Parks"= c("parks"),
                  "PoPy" = c("porin pyrint?"),
                  "PuiU" = c("puiu", "puistolan urheilijat", "puistolan urheilijat (puiu)", "puistolan sulkapallo"),
                  "RaSu" = c("rasu", "raision sulkapalloilijat", "raision sulkapallolijat", "raisusulka"),
                  "SaSu" = c("savon sulka"),
                  "Sjuba" = c("sjuba, sjunde? badmington boys and girls"),
                  "Smash" = c("smash sulkis"),
                  "Sulkaset" = c("sulkaset"),
                  "TS" = c("tapion sulka", "ts", "tapionsulka"),
                  "ToiVal" = c("toival", "toijalan valpas", "toijalan valpas ry"),
                  "TuPy" = c("pyrkiv?", "tupy", "turun pyrkiv?"),
                  "TuSu" = c("turun sulka"),
                  "TVS" = c("tvs"),
                  "Vaadin" = c("vaadin", "vaadin oy"),
                  "Vaasan Sulkis" = c("vasaan sulkis", "vaasansulkis"))
  return(mapValues(tolower(trimws(x)), mapping))
}

fixClass <- function(x) {
  mapping <- list("SN-H"="SEKANELINPELI Harraste",
                  "VL-H"="AIKUINEN/LAPSI nelinpeli HARRASTELUOKKA",
                  "VL-K"="AIKUINEN/LAPSI nelinpeli KILPALUOKKA",
                  "MK-50H"="Miesten KAKSINPELI +50 Harraste",
                  "MK-H"="Miesten KAKSINPELI Harraste",
                  "MN-H"="Miesten NELINPELI Harraste",
                  "NK-H"="Naisten KAKSINPELI Harraste",
                  "NN-H"="Naisten NELINPELI Harraste")
  return(mapValues(x, mapping))
}

mapValues <- function(x, mapping, factor_conversion = c("character", "integer"),
         na_to_na = TRUE, not_mapped = NULL) {
  ## factors are always converted
  if (is.factor(x)) {
    factor_conversion <- match.arg(factor_conversion)
    x <- as(x, factor_conversion)
    msg("Converted factor data to ", factor_conversion, type = "warning")
  }
  ## sanity check
  if (max(table(unlist(mapping), useNA = "ifany")) > 1) {
    msg("Some value(s) are mapped more than once", type = "warning")
  }
  data_class <- class(x)
  ## map
  mapped_values <- x
  for (new_value in names(mapping)) {
    if (new_value == "NA" && na_to_na) {
      new_value_casted <- NA
    } else {
      new_value_casted <- as(new_value, data_class)
    }
    mapped_values[x %in% mapping[[new_value]]] <- new_value_casted
  }
  if (!is.null(not_mapped)) {
    mapped_values[!(mapped_values %in% names(mapping))] <- as(not_mapped, data_class)
  }
  return(mapped_values)
}
