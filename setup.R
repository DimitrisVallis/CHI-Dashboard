repo_zip <- "https://github.com/DimitrisVallis/CHI-Dashboard/archive/refs/heads/main.zip"
dest_zip <- file.path(tempdir(), "CHI-Dashboard.zip")
docs_dir <- normalizePath(file.path(path.expand("~"), "Documents"), winslash = "/")

if (!dir.exists(docs_dir)) {
  docs_dir <- normalizePath(path.expand("~"), winslash = "/")
}

dest_dir <- file.path(docs_dir, "CHI-Dashboard")

message("Downloading project...")
download.file(repo_zip, destfile = dest_zip, mode = "wb")

message("Unzipping...")
unzip(dest_zip, exdir = docs_dir)

old_name <- file.path(docs_dir, "CHI-Dashboard-main")

if (file.exists(dest_dir)) {
  message("Folder already exists, updating files...")
  file.copy(list.files(old_name, full.names = TRUE), dest_dir, overwrite = TRUE)
  unlink(old_name, recursive = TRUE)
} else {
  file.rename(old_name, dest_dir)
}

analysis_path <- normalizePath(file.path(dest_dir, "analysis.R"), winslash = "/")

message("Project ready in: ", dest_dir)
message("To run the analysis, paste this into the console and press Enter:")
message('source("', analysis_path, '")')
