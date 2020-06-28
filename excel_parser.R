#!/usr/bin/env Rscript

#######
# Tested on R 3.6.2
## Written by B. Seelbinder, 2020-02-01
##################


#== Try to load packages. Install if not present ====
# Help output string if something goes wrong
CMDstring <- "
    Rscript excel_parser.R <Eingabe.xlsx> <Ausgabe> [--add_prices]

      <Eingabe.xlsx> Pfad zu einer '.xls' oder '.xlsx' Datei.
      <Ausgabe>      Präfix zur Ausgabe von Dateien. Für jede valide Seite in der
                     Eingabedatei wird eine separate Ausgabedatei angelegt.
      --add_prices   Ist optional. Wenn angegeben, werden Einzelverkaufspreise
                     mit ausgegeben.
"



#' Loads a package. Installs it if necessary.
#'
#' @param pkg Character(n). Name of CRAN library.
try_load_package <- function(pkg) {

  if (!require(pkg, character.only = TRUE, quietly = T)) {

    # pkg not installed
    install.packages(pkg, character.only = TRUE, quiet = T)
    tryCatch({
      library(pkg, character.only = T)
    }, except = function(e) {
      stop("Laden von Paket schlug fehl: ", pkg)
    })
  }
}


try_load_package("dplyr")
try_load_package("tidyr")
try_load_package("magrittr")
try_load_package("stringr")
try_load_package("readxl")
#try_load_package("tidyxl") # only required of formats shall be parsed

message(print("  >> Script erfolgreich geladen."))




#== Read Input Arguments ====
args = commandArgs(trailingOnly=TRUE)

# for testing, args gets overwritten
if ( length(args) == 1 && args[[1]] == "TEST1" ) {
  args <- c("Kontrolle_210-07977-AG_JAB-Hebeanlage_171219.xlsx", "Kontrolle_", "--add_prices")
} else if ( length(args) == 1 && args[[1]] == "TEST2" ) {
  args <- c("210-07088-3-MH_Seume__Burstmaschine_220217.xlsx", "Burst_", "--add_prices")
}

# Positional arguments# Positional arguments
if ( length(args) < 2 ) {
  stop("Es müssen Eingabe- und Ausgabe-Dateien angegeben werden.\n", CMDstring, call. = F)
}


# Optional Arguments
add_prices <- FALSE
if ( length(args) > 2 ) {
  # Currently, there is only one possible input argument
  if ( args[3] == "--add_prices" )
    add_prices <- T
  else
    stop(sprintf("Unbekanntes, optionales argument: '%s'\n", args[3]), CMDstring)
}


input_file <- args[1]  # file path to input Excel file
output_pref <- args[2] # Prefix for file output. Mutliple files are writting, starting with this string

if (!file.exists(input_file))
  stop(sprintf("Eingabedatei nicht gefunden: '%s'", input_file))

if ( output_pref == "" || grepl("^\\.", output_pref) )
  stop("Name für Ausgabedatei ungültig. Der Dateiname sollte nicht mit einem '.' beginnen und nicht leer sein. Dateiname: '", output_file, "'")

dir.create(dirname(output_pref), recursive = T, showWarnings = F)

message(sprintf("  >> Eingabedatei: '%s'", input_file))
message(sprintf("  >> Ausgabepräfix: '%s'", output_pref))



#== Parse Excel Files  ====
# Tasks
# - Main entries do not have numbers, or fabricants
# - remove items with 0 abundance
# - if type is missing, do not output it
# - some pages are irrelevant for the units. detect & ignore them
# - add prices optional
# - parse each sheet into separat text outputs

sheets <- readxl::excel_sheets(input_file)

# iterate all sheets
for (sheet in sheets) {

  datadf <- suppressWarnings(readxl::read_excel(input_file, sheet = sheet))

  # check if these are actual pricing sheets
  if (!("Bezeichnung" %in% colnames(datadf)))
    next()

  # clean data
  datadf %<>%
    select(Bezeichnung, Typ, Hersteller, Anzahl, contains("EP")) %>%
    #drop_na(Bezeichnung) %>%
    filter(is.na(Anzahl) | Anzahl > 0) %>%
    mutate(Titel = NA)

  # remove trailing NA
  last_valid <- datadf$Bezeichnung %>% is.na %>% not %>% which %>% last
  datadf <- datadf[1:last_valid,]

  # find Einzelpreis column (in case it is needed)
  ep_col <- which(grepl("EP", colnames(datadf)))

  # find Bezeichung with 'Klemme' and report the words
  klemme_index  <- grepl(x = datadf$Bezeichnung, "klemme") %>% which
  if ( length(klemme_index) > 0 ) {
    message(sprintf("%s\n%s\n%s",
                    "Mindestens eine Bezeichunng enthält das Wort 'Klemme' und wird zusammengefasst.",
                    "Bitte alle Bezeichnungen prüfen. Wurde ein falscher Begriff gewählt, muss das",
                    "Skript angepasst werden!"))
    print(unique(datadf$Bezeichnung[klemme_index]))
  }


  # '1 Satz Klemmen' should occur only once per title. This switch decides if 'Klemme' needs to be written out for the current title.
  group_klemme <- TRUE
  output_lines <- list()  # to be written into a new file
  for (ir in 1:nrow(datadf)) {
    # A cell can either be:
    #  A title
    #  An item
    #  A blank cell
    cell <- datadf[ir,]

    if ( cell %$% { !is.na(Bezeichnung) & is.na(Typ) & is.na(Hersteller)} ) {
      # Title line
      title <- sprintf("%s:", cell$Bezeichnung)
      output_lines <- c(output_lines, title)
      message(sprintf("Found title line: '%s'", title))
      group_klemme <- TRUE

    } else if ( !is.na(cell$Bezeichnung) ) {
      # Item Line
      # These may further be grouped
      if ( grepl(x = cell$Bezeichnung, "klemme,")[1] ) {
        if (group_klemme) {
          output_lines <- c(output_lines, sprintf("1 Satz Klemmen"))
          group_klemme <- FALSE
        }
      } else {
        output_lines <- c(output_lines, sprintf("%i Stk %s\n  Fabrikat: %s", cell$Anzahl, cell$Bezeichnung, cell$Hersteller)) %>%
          str_split("\n") %>%
          unlist
        if (!is.na(cell$Typ))
          output_lines <- c(output_lines, sprintf("  Typ: %s", cell$Typ))
        if (add_prices)
          output_lines <- c(output_lines, sprintf("  EP: %s", cell[[ep_col]]))
      }

    } else {
      # Blank Line
      output_lines <- c(output_lines, "")
    }
  }

  output_lines %<>% unlist %>% str_remove_all("\\r")
  output_file <- paste0(output_pref, ".", sheet, ".txt")
  message("  >> Writing output to: ", output_file)
  writeLines(output_lines, output_file)
}


