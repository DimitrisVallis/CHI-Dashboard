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
library(ggplot2)

repo_zip <- "https://github.com/DimitrisVallis/CHI-Dashboard/archive/refs/heads/main.zip"
dest_zip <- file.path(tempdir(), "CHI-Dashboard.zip")
dest_dir <- normalizePath(file.path(tempdir(), "CHI-Dashboard-main"), winslash = "/", mustWork = FALSE)

message("Cleaning up old files...")
if (file.exists(dest_zip)) file.remove(dest_zip)
if (dir.exists(dest_dir)) unlink(dest_dir, recursive = TRUE)
message("Downloading project...")
download.file(repo_zip, destfile = dest_zip, mode = "wb")
message("Unzipping...")
unzip(dest_zip, exdir = tempdir())
message("Project ready.")

plot_type <- ""
while (!plot_type %in% c("means", "estimates")) {
  plot_type <- str_trim(tolower(readline(prompt = "Enter plot type - type 'means' or 'estimates' and press Enter: ")))
  if (!plot_type %in% c("means", "estimates")) message("Invalid choice. Please type 'means' or 'estimates'.")
}
message("Running analysis with plot type: ", plot_type)

old_wd <- setwd(dest_dir)
on.exit(setwd(old_wd))
source(file.path(dest_dir, "analysis.R"))
