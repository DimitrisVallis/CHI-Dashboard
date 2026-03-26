repo_zip <- "https://github.com/DimitrisVallis/CHI-Dashboard/archive/refs/heads/main.zip"
dest_zip <- file.path(tempdir(), "CHI-Dashboard.zip")
dest_dir <- normalizePath(file.path(tempdir(), "CHI-Dashboard-main"), winslash = "/", mustWork = FALSE)

message("Downloading project...")
download.file(repo_zip, destfile = dest_zip, mode = "wb")

message("Unzipping...")
unzip(dest_zip, exdir = tempdir())

analysis_path <- normalizePath(file.path(dest_dir, "analysis.R"), winslash = "/", mustWork = FALSE)

message("Project ready.")
message("To run the analysis, paste this into the console and press Enter:")
message('source("', analysis_path, '")')
