read_member_data <- function() {
  x <- openxlsx::read.xlsx(xlsxFile = "C:/Users/Esa/Downloads/Jasentiedot (2021-04-07).xlsx", detectDates = T) 
  colnames(x) <- stringr::str_replace_all(colnames(x), "ä", "a") 
  colnames(x) <- stringr::str_replace_all(colnames(x), "ö", "o")
  x
}

process_member_data <- function(x, age_ref_date = Sys.Date()) {
  x %>%
    dplyr::mutate(name = paste(Etunimi, Sukunimi)) %>%
    dplyr::filter(Aktiivinen == "X") %>%
    dplyr::mutate(age = floor(as.integer(difftime(age_ref_date, Syntymaaika, units = "days"))/365)) %>%
    dplyr::select(name, gender = Sukupuoli, age, group = Ryhma)
}

check_member_data <- function(x) {
  ## check birth dates
  inds <- which(is.na(x$age))
  ok <- TRUE
  if (length(inds) > 0) {
    cat("Missing birth date:\n", paste0(x$name[inds], collapse = "\n"), sep = "")
    ok <- FALSE
  }
  ## check genders
  inds <- which(is.na(x$gender))
  if (length(inds) > 0) {
    cat("Missing gender:\n", paste0(x$Sukupuoli[inds], collapse = "\n"), sep = "")
    ok <- FALSE
  }
  if (ok) cat("Everything ok")
}


