#!/usr/bin/env Rscript

#######
# Tested on R 3.6.2
## Written by B. Seelbinder, 2020-02-01
##################


#== Try to load packages. Install if not present ====
# Help output string if something goes wrong
CMDstring <- "
    Rscript excel_parser.R <Ausgabe.xlsx> <Ausgabe.xlsx> [--add_prices]
    
      --add_prices Ist optional.
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
try_load_package("tidyxl")
try_load_package("readxl")
try_load_package("writexl")

message(print("  >> Script erfolgreich geladen."))



#== Read Input Arguments ====
args = commandArgs(trailingOnly=TRUE)

# FOR TESTING ONLY
args <- c("210-07088-3-MH_Seume__Burstmaschine_220217.xlsx", "steph_text", "add_prices")


# Positional Arguments
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


input_file <- args[1]
output_pref <- args[2]

if (!file.exists(input_file))
  stop(sprintf("Eingabedatei nicht gefunden: '%s'", input_file))

if ( output_file == "" || grepl("^\\.", output_file) )
  stop("Name für Ausgabedatei ungültig. Der Dateiname sollte nicht mit einem '.' beginnen und nicht leer sein. Dateiname: '", output_file, "'")

dir.create(dirname(output_file), recursive = T, showWarnings = F)

message(sprintf("  >> Eingabedatei: '%s'", input_file))
message(sprintf("  >> Ausgabepräfix: '%s'", output_pref))



#== Parse Excel Files  ====
# Tasks
# - bold items make a new "main entry"
# - entries under "Nr." marke a new sub-entry
#   - Encoding: main_entry sub[n] sub[n-1] ...
# - remove items with 0 abundance
# - some pages are irrelevant for the units. detect & ignore them
# - add prices optional
# - parse each sheet into separat text outputs


### Pt1. Determine Bold Cells

# Load Excel File
exdat <- tidyxl::xlsx_cells(input_file)
# Note: tidyxl fails to parse german letters correctly...
#   So we fix that here by loading the sheet names with 'readxl'
exdat %<>% mutate(sheet = factor(sheet,
                                 levels = tidyxl::xlsx_sheet_names(input_file),
                                 labels = readxl::excel_sheets(input_file)) %>%
                    as.character())

# get cell indices at which bold names are used
formats <- xlsx_formats(input_file)
bold_cells <- exdat[exdat$local_format_id %in% which(formats$local$font$bold), c("sheet", "address", "character")]
bold_cells %<>% mutate(is_bold = TRUE)
exdat %<>%
  left_join(bold_cells) %>%
  mutate(is_bold = ifelse(data_type == "blank", F, is_bold)) %>%
  select(sheet, address, data_type, character, style_format, is_bold)

# give is more informative names
exdat %<>% dplyr::rename(main_unit = is_bold)


## translate address to actual indices
# split address into x & y
exdat %<>% 
  mutate(coladd = str_extract(address, "^[A-Z]+"),
         rowadd = str_extract(address, "\\d+$") %>% as.integer())

# make proper index mapping
colorder <- tibble(cols = unique(exdat$coladd)) %>%
  mutate(collen = nchar(cols)) %>%
  arrange(collen, cols)

# translate letters into coordinate
exdat %<>% mutate(coladd = factor(coladd, levels = colorder[["cols"]]) %>% as.integer())

## remove everything that is not relevant to find the sections
# only bold text in column 2 (Bezeichnung) are relevant
exdat %<>% 
  filter(main_unit & coladd == 2)




### Pt2. Parse content & format output
# Strategy:
#   - we now the main unit entries & their indices
#   - hence, jump from entry to entry, remove NA blocks at a time, etc.
sheets <- readxl::excel_sheets(input_file)

# iterate all sheets
for (sheet in sheets) {

  datadf <- suppressWarnings(readxl::read_excel(input_file, sheet = sheet))
  # check if these are actual pricing sheets
  if (!("Bezeichnung" %in% colnames(datadf)))
    next()
  
  datadf %<>%
    select(`Nr.`, Bezeichnung, Typ, Hersteller, Anzahl, `VK\r\nEP`)
  unitinfo <- exdat %>% filter(sheet == !!sheet)
  rowindices <- c(unitinfo$rowadd - 1, nrow(datadf)) # first index to start, last index to end
  
  
  output_lines <- list()
  
  
  # Iterate all chunks
  for (i in 1:(length(rowindices)-1)) {
    chunk <- datadf[ (rowindices[i]):(rowindices[i+1]-1), ]
    
    main_unit <- chunk$Bezeichnung[[1]]
    chunk <- chunk[-1,]
    chunk %<>% filter(is.na(Anzahl) | Anzahl > 0)
    
    # If "Nr." column contains additional strings, we use a slightly different formatting
    # Note: this code could be compressed significantly by forcing two more columns for long-formatting & grouping
    if (F %in% is.na(chunk$`Nr.`)) {
      sub_idx <- c(1, which(is.na(chunk$Bezeichnung)), nrow(chunk)) %>% unique()
      first <- T
      
      for(j in 1:(length(sub_idx)-1)) {
        sub_chunk <- chunk[ sub_idx[j]:(sub_idx[j+1]-1), ] %>% drop_na(Bezeichnung)
        sub_unit <- sub_chunk[, 1 ] %>% 
          drop_na() %>%
          .[["Nr."]] %>%
          rev %>%
          paste(collapse = " ")
        
        # if the string is empty, print a new header only if it is the first entry of this chunk
        if (first || sub_unit != "") {
          # print header
          first <- F
          output_lines <- c(output_lines, paste(main_unit, sub_unit) %>% trimws %>% paste0(":"))
        } else {
          output_lines <- c(output_lines, "")
        }
        
        if (!add_prices)
          s <- sprintf("%i Stk %s\n  Fabrikat: %s\n  Typ: %s\n", sub_chunk$Anzahl, sub_chunk$Bezeichnung, sub_chunk$Hersteller, sub_chunk$Typ) %>% str_split("\n") %>% unlist
        else
          s <- sprintf("%i Stk %s\n  Fabrikat: %s\n  Typ: %s\n  EP: %.2f€\n", sub_chunk$Anzahl, sub_chunk$Bezeichnung, sub_chunk$Hersteller, sub_chunk$Typ, sub_chunk[["VK\r\nEP"]]) %>% str_split("\n") %>% unlist
        output_lines <- c(output_lines, s)
      }
  
    } else {
      chunk %<>% drop_na(Bezeichnung)
      output_lines <- c(output_lines, paste0(main_unit, ":"))
      if (!add_prices)
        s <- sprintf("%i Stk %s\n  Fabrikat: %s\n  Typ: %s\n", chunk$Anzahl, chunk$Bezeichnung, chunk$Hersteller, chunk$Typ) %>% str_split("\n") %>% unlist
      else
        s <- sprintf("%i Stk %s\n  Fabrikat: %s\n  Typ: %s\n  EP: %.2f€\n", chunk$Anzahl, chunk$Bezeichnung, chunk$Hersteller, chunk$Typ, chunk[["VK\r\nEP"]]) %>% str_split("\n") %>% unlist
      output_lines <- c(output_lines, s)
    }
  }
  
  output_lines %<>% unlist
  output_file <- paste0(output_pref, ".", sheet, ".txt")
  message("  >> Writing output to: ", output_file)
  writeLines(output_lines, output_file)
}


