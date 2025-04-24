# global.R

# Chargement des bibliothèques
library(shiny)
library(shinyjs)
library(readxl)
library(rhandsontable)
library(openxlsx2)
library(reactable)
library(httr)
library(jsonlite)

source("~/work/lereuf/deconstructurateur/helpers/downloadModule.R")


# Fonction : convertir un numéro de colonne en lettre Excel
num_to_excel_col <- function(n) {
  if (is.null(n) || length(n) == 0 || is.na(n)) return("?")
  div <- n
  col <- ""
  while (div > 0) {
    mod <- (div - 1) %% 26
    col <- paste0(LETTERS[mod + 1], col)
    div <- (div - mod - 1) %/% 26
  }
  return(col)
}

# Fonction : convertir une colonne Excel (lettres) en numéro
excel_col_to_num <- function(col_str) {
  chars <- strsplit(col_str, "")[[1]]
  sum((match(chars, LETTERS)) * 26^(rev(seq_along(chars)) - 1))
}