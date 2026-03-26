repo_zip <- "https://github.com/DimitrisVallis/CHI-Dashboard/archive/refs/heads/main.zip"
dest_zip <- file.path(tempdir(), "CHI-Dashboard.zip")
dest_dir <- file.path(Sys.getenv("USERPROFILE"), "Documents", "CHI-Dashboard")

message("Downloading project...")
download.file(repo_zip, destfile = dest_zip, mode = "wb")

message("Unzipping...")
unzip(dest_zip, exdir = file.path(Sys.getenv("USERPROFILE"), "Documents"))

old_name <- file.path(Sys.getenv("USERPROFILE"), "Documents", "CHI-Dashboard-main")
if (file.exists(old_name)) file.rename(old_name, dest_dir)

message("Project downloaded and ready in: ", dest_dir)
message("To run the analysis, paste this into the console and press Enter:")
message('source("', file.path(dest_dir, "analysis.R"), '")')