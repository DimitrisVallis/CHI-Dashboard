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
  source(file.path(dest_dir, "analysis.R"))
} else {
  message("To run the analysis later, paste this into the console and press Enter:")
  message('source("', normalizePath(file.path(dest_dir, "analysis.R"), winslash = "/"), '")')
}
