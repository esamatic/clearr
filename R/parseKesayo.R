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
    dplyr::select(c("Etunimi", "Sukunimi", "Lempinimi", "Seura",
                    "1..PELILUOKKA", "1..PELILUOKASSA.nelinpeliparin.nimi",
                    "2..PELILUOKKA", "2..PELILUOKASSA.nelinpeliparin.nimi",
                    "3..PELILUOKKA", "3..PELILUOKASSA.nelinpeliparin.nimi",
                    Terveiset="Terveiset,.risut.ja.ruusut.kilpailujärjestäjille."))

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
    dplyr::mutate(Seura = fixClub(Seura)) %>%
    dplyr::mutate(Etunimi = ifelse(is.na(Lempinimi), Etunimi, paste0(Etunimi, ' "', Lempinimi, '"'))) %>%
    dplyr::select(-Lempinimi) %>%
    dplyr::mutate(Seura = replace(Seura, is.na(Seura) | Seura == "ei", "")) %>%
    dplyr::arrange(Peliluokka, Sukunimi, Etunimi) %>%
    dplyr::filter(!(is.na(Peliluokka) & is.na(Pari))) %>%
    dplyr::mutate(Terveiset = replace(Terveiset, is.na(Terveiset), ""))
  return(x)
}

fixClub <- function(x) {
  club <- tolower(trimws(x))
  mapping <- list("Clear" = c("clear", "clera", "clear?"),
                  "ESB" = c("esb", "espoon sulkapallo badminton"),
                  "Drive" = c("drive", "sulkapalloseura drive"),
                  "Euran Veivi" = c("euran veivi"),
                  "Geneve Badminton Club" = c("geneve badminton club"),
                  "HämSu" = c("hämsu", "hämeenlinnan sulkapalloilijat"),
                  "HalSu" = c("halsu", "halikon sulkis"),
                  "HanSu" = c("hansu"),
                  "HYSY" = c("hysy", "helsingin yliopiston sulkapalloyhdistys (hysy)"),
                  "Joen Sulka" = c("joen sulka"),
                  "Juankosken Pyrkivä"=c("juankosken pyrkivä"),
                  "KaaSu" = c("kaasu", "kaarinan sulka"),
                  "Karjalohjan Sulka" = c("karjalohjan sulka"),
                  "Klaki" = c("klaki"),
                  "Kosken Liikuntahalli" = c("kosken liikuntahalli"),
                  "Kyry" = c("kyry"),
                  "Lohjan Teho" = c("lohjan teho"),
                  "Loimaan Seudun Sulkis" = c("loimaan seudun sulkis"),
                  "NBC" = c("nbc"),
                  "ÖIF" = c("öif"),
                  "Onnen Urheilijat" = c("onnen urheiljat"),
                  "OSUKO" = c("osuko"),
                  "ParBa" = c("parba", "paraisten badminton"),
                  "PoPy" = c("porin pyrintö"),
                  "PuiU" = c("puiu", "puistolan urheilijat", "puistolan urheilijat (puiu)"),
                  "RaSu" = c("rasu", "raision sulkapalloilijat", "raision sulkapallolijat"),
                  "SaSu" = c("savon sulka"),
                  "Sjuba" = c("sjuba, sjundeå badmington boys and girls"),
                  "Smash" = c("smash sulkis"),
                  "Sulkaset" = c("sulkaset"),
                  "TS" = c("tapion sulka", "ts"),
                  "ToiVal" = c("toival", "toijalan valpas", "toijalan valpas ry"),
                  "TuPy" = c("pyrkivä", "tupy", "turun pyrkivä"),
                  "TuSu" = c("turun sulka"),
                  "TVS" = c("tvs"),
                  "Vaadin" = c("vaadin", "vaadin oy"),
                  "Vaasan Sulkis" = c("vasaan sulkis", "vaasansulkis"))
  for (new_name in names(mapping)) {
    club[club %in% mapping[[new_name]]] <- new_name
  }
  return(club)
}
