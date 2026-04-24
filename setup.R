required_packages <- c("sandwich", "lmtest", "stringr", "readxl", "dplyr",
                       "purrr", "tibble", "stringdist", "jsonlite", "ggplot2", "stringr")
missing_packages <- required_packages[!required_packages %in% installed.packages()[, "Package"]]
if (length(missing_packages) > 0) {
  message("Installing missing packages: ", paste(missing_packages, collapse = ", "))
  install.packages(missing_packages)
}
library(sandwich)
library(lmtest)
library(stringr)
library(readxl)
library(dplyr)
library(purrr)
library(tibble)
library(stringdist)
library(jsonlite)
library(stringr)

repo_zip <- "https://github.com/DimitrisVallis/CHI-Dashboard/archive/refs/heads/main.zip"
dest_zip <- file.path(tempdir(), "CHI-Dashboard.zip")
dest_dir <- normalizePath(file.path(tempdir(), "CHI-Dashboard-main"), winslash = "/", mustWork = FALSE)

message("Downloading project...")
download.file(repo_zip, destfile = dest_zip, mode = "wb")
message("Unzipping...")
unzip(dest_zip, exdir = tempdir())
message("Project ready.")

run <- str_trim(toupper(readline(prompt = "Run the analysis now? (Y/N): ")))
while (!run %in% c("Y", "N")) {
  message("Invalid input. Please type Y or N.")
  run <- str_trim(toupper(readline(prompt = "Run the analysis now? (Y/N): ")))
}
if (run == "Y") {
  old_wd <- setwd(dest_dir)
  on.exit(setwd(old_wd))
  source(file.path(dest_dir, "analysis.R"))
} else {
  analysis_path <- normalizePath(file.path(dest_dir, "analysis.R"), winslash = "/")
  later_cmd <- sprintf("setwd('%s'); source('%s')", dest_dir, analysis_path)
  message("To run the analysis later, paste this into the console and press Enter:")
  message(later_cmd)
}
