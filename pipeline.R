

#matching
min_score <- 0.58

#Purpose: convert column numbers → Excel letters (1→A, 26→Z, 27→AA)
excel_col_letters <- function(n) {
  vapply(n, function(x) {
    s <- ""
    while (x > 0) {
      r <- (x - 1) %% 26
      s <- paste0(LETTERS[r + 1], s)
      x <- (x - 1) %/% 26
    }
    s
  }, FUN.VALUE = character(1))
}


# Clean header names safely

clean_names_safe <- function(x) {
  x <- as.character(x)
  x <- str_trim(x)
  x <- str_replace_all(x, "[\r\n\t]", " ")
  x <- str_replace_all(x, "(\\d+)\\+", "\\1_plus")   # 75+ -> 75_plus
  x <- str_replace_all(x, "[^A-Za-z0-9]+", "_")
  x <- str_replace_all(x, "_+", "_")
  x <- str_replace_all(x, "^_+|_+$", "")
  x[x == "" | is.na(x)] <- "V"
  x <- ifelse(str_detect(x, "^[0-9]"), paste0("V_", x), x)
  base::make.unique(x, sep = "_")
}

# N-gram scoring functions with normalization

extract_ngrams <- function(text, n = 2) {
  words <- unique(tokenize_words(text))  # deduplicate words before building ngrams
  if (length(words) < n) return(character(0))
  
  ngrams <- character(length(words) - n + 1)
  for (i in seq_along(ngrams)) {
    ngrams[i] <- paste(words[i:(i + n - 1)], collapse = " ")
  }
  ngrams
}

# Score based on n-gram overlap, normalized by candidate length
# Token overlap: penalise both unmatched candidate words AND unmatched key words
score_overlap_normalized <- function(phrase, cand) {
  kw <- tokenize_words(phrase)
  cw <- tokenize_words(cand)
  
  if (length(kw) == 0 || length(cw) == 0) return(0)
  
  matches <- sum(kw %in% cw)
  
  # Geometric mean of (matches/candidate_length) and (matches/key_length)
  # This penalises both a verbose candidate and a key with unmatched words
  precision <- matches / length(cw)   # what fraction of candidate is covered
  recall    <- matches / length(kw)   # what fraction of key is covered
  
  if (precision + recall == 0) return(0)
  2 * precision * recall / (precision + recall)  # F1
}

# N-gram score: bidirectional (F1 between phrase ngrams and candidate ngrams)
score_ngram <- function(phrase, cand, n = 2) {
  phrase_ngrams <- unique(extract_ngrams(phrase, n))
  cand_ngrams   <- unique(extract_ngrams(cand, n))
  
  # Fallback: if either side is too short for bigrams, compare single tokens directly
  if (length(phrase_ngrams) == 0 || length(cand_ngrams) == 0) {
    pw <- unique(tokenize_words(phrase))
    cw <- unique(tokenize_words(cand))
    if (length(pw) == 0 || length(cw) == 0) return(0)
    matches <- sum(pw %in% cw)
    precision <- matches / length(cw)
    recall    <- matches / length(pw)
    if (precision + recall == 0) return(0)
    return(2 * precision * recall / (precision + recall))
  }
  
  intersection <- sum(phrase_ngrams %in% cand_ngrams)
  precision <- intersection / length(cand_ngrams)
  recall    <- intersection / length(phrase_ngrams)
  if (precision + recall == 0) return(0)
  2 * precision * recall / (precision + recall)
}
# Combined score: 50% token F1 + 50% bigram F1
score_combined <- function(phrase, cand) {
  token_score <- score_overlap_normalized(phrase, cand)
  ngram_score <- score_ngram(phrase, cand, n = 2)
  (token_score * 0.5) + (ngram_score * 0.5)
}

# Numeric detection (robust)

# optional sign
# either 123 or 1,234 or 12,345,678
# optional decimals .5, .00, etc.
# This is important later because:
# it’s used to identify the first data row

is_fully_numeric <- function(x) {
  x <- str_trim(as.character(x))
  ok <- !is.na(x) & x != ""
  pat <- "^[+-]?((\\d{1,3}(,\\d{3})+)|\\d+)(\\.\\d+)?$"
  ok & str_detect(x, pat)
}

find_marker_col_in_row <- function(mat, data_row, marker = "E92000001") {
  if (is.na(data_row) || data_row < 1 || data_row > nrow(mat)) return(NA_integer_)
  r <- str_trim(as.character(mat[data_row, ]))
  which(!is.na(r) & r == marker)[1]
}


# Dataset name from row 1

get_dataset_name <- function(mat, sheet_name) {
  # Always extract dataset name from row 1
  r1 <- as.character(mat[1, ])
  r1 <- r1[!is.na(r1) & str_trim(r1) != ""]
  if (length(r1) == 0) {
    paste0("unidentified_sheet_", sheet_name)
  } else {
    r1[[1]]
  }
}

# Build colnames from header rows
#hdr is a header block
# collects all non-empty header cells in that column (across header rows) and concatenates them with _
#This is how you handle merged Excel headers like:
#row 3: “Rent arrears”
#row 4: “Increase in rent”
# → becomes Rent_arrears_Increase_in_rent (after cleaning)
build_colnames <- function(mat, header_rows, col_idx) {
  if (length(header_rows) == 0) {
    v <- paste0("V", col_idx)
    cleaned <- clean_names_safe(v)
    return(list(
      clean            = cleaned,
      raw              = v,
      clean_pre_unique = cleaned  # no duplicates possible here
    ))
  }
  
  hdr <- mat[header_rows, col_idx, drop = FALSE]
  if (!is.matrix(hdr)) hdr <- matrix(hdr, nrow = 1)
  
  clean_number_context <- function(txt) {
    txt <- gsub("(\\d+)\\+", "\\1_plus", txt)                          # 75+ -> 75_plus
    txt <- gsub("(\\d+)-(\\d+)", "\\1_\\2", txt)                       # 16-17 -> 16_17 (already was there)
    txt <- gsub("(?<=[a-zA-Z)])\\d+|\\d+(?=[a-zA-Z])", "", txt, perl = TRUE)  # strip footnote digits only
    txt
  }
  
  raw_cn <- vapply(seq_len(ncol(hdr)), function(j) {
    vals <- str_trim(as.character(hdr[, j]))
    vals <- vals[!is.na(vals) & vals != ""]
    vals <- unique(vals)
    if (length(vals) == 0) return("")
    vals_clean <- gsub("(?<=[a-zA-Z)])\\d+|\\d+(?=[a-zA-Z])", "", vals, perl = TRUE)
    vals_clean <- str_trim(vals_clean)
    # After stripping footnote digits, some vals may become identical — deduplicate
    keep_idx <- !duplicated(tolower(vals_clean))
    vals <- vals[keep_idx]
    vals_clean <- vals_clean[keep_idx]
    if (length(vals) == 0) return("")
    # Drop vals whose words are all already present in another val (redundant parent labels)
    is_redundant <- vapply(seq_along(vals_clean), function(a) {
      wa <- tolower(strsplit(vals_clean[a], "\\s+")[[1]])
      wa <- wa[wa != ""]
      any(vapply(seq_along(vals_clean)[-a], function(b) {
        wb <- tolower(strsplit(vals_clean[b], "\\s+")[[1]])
        wb <- wb[wb != ""]
        all(wa %in% wb)
      }, logical(1)))
    }, logical(1))
    vals <- vals[!is_redundant]
    if (length(vals) == 0) return("")
    txt <- paste(vals, collapse = " ")
    str_trim(clean_number_context(txt))
  }, FUN.VALUE = character(1))
  
  clean_cn <- vapply(seq_len(ncol(hdr)), function(j) {
    vals <- str_trim(as.character(hdr[, j]))
    vals <- vals[!is.na(vals) & vals != ""]
    vals <- unique(vals)
    if (length(vals) == 0) return("")
    vals_clean <- gsub("(?<=[a-zA-Z)])\\d+|\\d+(?=[a-zA-Z])", "", vals, perl = TRUE)
    vals_clean <- str_trim(vals_clean)
    # After stripping footnote digits, some vals may become identical — deduplicate
    keep_idx <- !duplicated(tolower(vals_clean))
    vals <- vals[keep_idx]
    vals_clean <- vals_clean[keep_idx]
    if (length(vals) == 0) return("")
    # Drop vals whose words are all already present in another val (redundant parent labels)
    is_redundant <- vapply(seq_along(vals_clean), function(a) {
      wa <- tolower(strsplit(vals_clean[a], "\\s+")[[1]])
      wa <- wa[wa != ""]
      any(vapply(seq_along(vals_clean)[-a], function(b) {
        wb <- tolower(strsplit(vals_clean[b], "\\s+")[[1]])
        wb <- wb[wb != ""]
        all(wa %in% wb)
      }, logical(1)))
    }, logical(1))
    vals <- vals[!is_redundant]
    if (length(vals) == 0) return("")
    txt <- paste(vals, collapse = "_")
    str_trim(clean_number_context(txt))
  }, FUN.VALUE = character(1))
  
  blank <- which(clean_cn == "" | is.na(clean_cn))
  if (length(blank) > 0) {
    clean_cn[blank] <- paste0("V", col_idx[blank])
    raw_cn[blank]   <- paste0("V", col_idx[blank])
  }
  
  # Apply the character-cleaning steps from clean_names_safe MANUALLY,
  # but WITHOUT the final make.unique() call — so duplicates are preserved
  safe_no_unique <- function(x) {
    x <- as.character(x)
    x <- str_trim(x)
    x <- str_replace_all(x, "[\r\n\t]", " ")
    x <- str_replace_all(x, "(\\d+)\\+", "\\1_plus")   # 75+ -> 75_plus
    x <- str_replace_all(x, "[^A-Za-z0-9]+", "_")
    x <- str_replace_all(x, "_+", "_")
    x <- str_replace_all(x, "^_+|_+$", "")
    x[x == "" | is.na(x)] <- "V"
    x <- ifelse(str_detect(x, "^[0-9]"), paste0("V_", x), x)
    x
  }
  
  pre_unique <- safe_no_unique(clean_cn)
  
  list(
    clean            = base::make.unique(safe_no_unique(clean_cn), sep = "_"),
    raw              = raw_cn,
    clean_pre_unique = pre_unique   # identical strings for count+pct pairs
  )
}


# Fix first 2 columns if blank

fix_region_columns <- function(col_names, col_idx) {
  out <- col_names
  
  if (1 %in% col_idx) {
    j <- which(col_idx == 1)
    if (length(j) == 1 && str_detect(out[j], "^V")) out[j] <- "Region_code"
  }
  if (2 %in% col_idx) {
    j <- which(col_idx == 2)
    if (length(j) == 1 && str_detect(out[j], "^V")) out[j] <- "Region_name"
  }
  
  base::make.unique(out, sep = "_")
}


# Excel ref like "AX12" to row/col

excel_ref_to_rc <- function(ref) {
  m <- str_match(ref, "^([A-Z]+)([0-9]+)$")
  if (any(is.na(m))) return(c(NA_integer_, NA_integer_))
  letters <- m[2]
  row <- as.integer(m[3])
  
  col <- 0L
  for (ch in strsplit(letters, "")[[1]]) {
    col <- col * 26L + match(ch, LETTERS)
  }
  c(row, col)
}


# Parse merge ranges from the xlsx XML
#Excel merges are stored in the sheet XML.
# Then for each sheet:
# find its target XML file
# read it
# extract merge ranges:
#So you end up with a list like:
#merges_by_sheet[["A2P"]] = c("A1:C1", "D3:D5", ...)

get_merge_ranges_from_xlsx <- function(path) {
  td <- tempfile("xlsx_unzip_")
  dir.create(td)
  unzip(path, exdir = td)
  
  wb_path   <- file.path(td, "xl", "workbook.xml")
  rels_path <- file.path(td, "xl", "_rels", "workbook.xml.rels")
  
  wb_txt   <- paste(readLines(wb_path, warn = FALSE), collapse = "")
  rels_txt <- paste(readLines(rels_path, warn = FALSE), collapse = "")
  
  sheet_nodes <- str_match_all(wb_txt, "<sheet[^>]+/>")[[1]]
  sheet_name  <- str_match(sheet_nodes, 'name="([^"]+)"')[, 2]
  rid         <- str_match(sheet_nodes, 'r:id="([^"]+)"')[, 2]
  
  rel_nodes  <- str_match_all(rels_txt, "<Relationship[^>]+/>")[[1]]
  rel_id     <- str_match(rel_nodes, 'Id="([^"]+)"')[, 2]
  rel_target <- str_match(rel_nodes, 'Target="([^"]+)"')[, 2]
  
  merges_by_sheet <- setNames(vector("list", length(sheet_name)), sheet_name)
  
  for (i in seq_along(sheet_name)) {
    targ <- rel_target[match(rid[i], rel_id)]
    if (is.na(targ) || !str_detect(targ, "^worksheets/")) {
      merges_by_sheet[[sheet_name[i]]] <- character(0)
      next
    }
    
    sheet_xml_path <- file.path(td, "xl", targ)
    if (!file.exists(sheet_xml_path)) {
      merges_by_sheet[[sheet_name[i]]] <- character(0)
      next
    }
    
    s_txt <- paste(readLines(sheet_xml_path, warn = FALSE), collapse = "")
    refs <- str_match_all(s_txt, 'mergeCell[^>]+ref="([^"]+)"')[[1]][, 2]
    if (length(refs) == 0) refs <- character(0)
    
    merges_by_sheet[[sheet_name[i]]] <- refs
  }
  
  merges_by_sheet
}

# Add this NEW function right after get_merge_ranges_from_xlsx
# REPLACE the entire get_hidden_columns_from_xlsx function with this expanded version:
get_hidden_cols_and_rows_from_xlsx <- function(path) {
  td <- tempfile("xlsx_unzip_")
  dir.create(td)
  unzip(path, exdir = td)
  
  wb_path   <- file.path(td, "xl", "workbook.xml")
  rels_path <- file.path(td, "xl", "_rels", "workbook.xml.rels")
  
  wb_txt   <- paste(readLines(wb_path, warn = FALSE), collapse = "")
  rels_txt <- paste(readLines(rels_path, warn = FALSE), collapse = "")
  
  sheet_nodes <- str_match_all(wb_txt, "<sheet[^>]+/>")[[1]]
  sheet_name  <- str_match(sheet_nodes, 'name="([^"]+)"')[, 2]
  rid         <- str_match(sheet_nodes, 'r:id="([^"]+)"')[, 2]
  
  rel_nodes  <- str_match_all(rels_txt, "<Relationship[^>]+/>")[[1]]
  rel_id     <- str_match(rel_nodes, 'Id="([^"]+)"')[, 2]
  rel_target <- str_match(rel_nodes, 'Target="([^"]+)"')[, 2]
  
  hidden_by_sheet <- setNames(vector("list", length(sheet_name)), sheet_name)
  
  for (i in seq_along(sheet_name)) {
    targ <- rel_target[match(rid[i], rel_id)]
    if (is.na(targ) || !str_detect(targ, "^worksheets/")) {
      hidden_by_sheet[[sheet_name[i]]] <- list(cols = integer(0), rows = integer(0))
      next
    }
    
    sheet_xml_path <- file.path(td, "xl", targ)
    if (!file.exists(sheet_xml_path)) {
      hidden_by_sheet[[sheet_name[i]]] <- list(cols = integer(0), rows = integer(0))
      next
    }
    
    s_txt <- paste(readLines(sheet_xml_path, warn = FALSE), collapse = "")
    
    # HIDDEN COLUMNS
    col_nodes <- str_match_all(s_txt, '<col[^>]+/>')[[1]]
    hidden_cols <- integer(0)
    for (col_node in col_nodes) {
      is_hidden <- str_detect(col_node, 'hidden="1"')
      if (is_hidden) {
        min_col <- as.integer(str_match(col_node, 'min="([0-9]+)"')[2])
        max_col <- as.integer(str_match(col_node, 'max="([0-9]+)"')[2])
        if (!is.na(min_col) && !is.na(max_col)) {
          hidden_cols <- c(hidden_cols, min_col:max_col)
        }
      }
    }
    
    # HIDDEN ROWS
    row_nodes <- str_match_all(s_txt, '<row[^>]+>')[[1]]
    hidden_rows <- integer(0)
    for (row_node in row_nodes) {
      is_hidden <- str_detect(row_node, 'hidden="1"')
      if (is_hidden) {
        row_num <- as.integer(str_match(row_node, 'r="([0-9]+)"')[2])
        if (!is.na(row_num)) {
          hidden_rows <- c(hidden_rows, row_num)
        }
      }
    }
    
    hidden_by_sheet[[sheet_name[i]]] <- list(
      cols = unique(hidden_cols),
      rows = unique(hidden_rows)
    )
  }
  
  hidden_by_sheet
}

# Pad df to cover merge ranges

pad_to_merges <- function(df, merges) {
  if (is.null(merges) || length(merges) == 0) return(df)
  
  max_r <- nrow(df)
  max_c <- ncol(df)
  
  for (rng in merges) {
    parts <- strsplit(rng, ":", fixed = TRUE)[[1]]
    rc1 <- excel_ref_to_rc(parts[1])
    rc2 <- excel_ref_to_rc(ifelse(length(parts) == 2, parts[2], parts[1]))
    if (any(is.na(c(rc1, rc2)))) next
    max_r <- max(max_r, rc1[1], rc2[1])
    max_c <- max(max_c, rc1[2], rc2[2])
  }
  
  if (nrow(df) < max_r) {
    add <- max_r - nrow(df)
    df <- bind_rows(df, as.data.frame(matrix(NA, nrow = add, ncol = ncol(df))))
  }
  if (ncol(df) < max_c) {
    add <- max_c - ncol(df)
    for (k in seq_len(add)) df[[ncol(df) + 1]] <- NA
  }
  df
}


# Fill merged cells with top-left value
#Handling merges in a data frame: pad_to_merges() + fill_merged_cells_all()

# Merges can reference cells outside the currently-read rectangular area (esp. if readxl stops early).
# So this block ensures df has enough rows/cols to cover the largest merge ref.
# It converts refs like AX12 to (row=12, col=50) via excel_ref_to_rc()
# Tracks max_r, max_c
# Pads with NA rows/columns to that size
fill_merged_cells_all <- function(df, merges) {
  if (is.null(merges) || length(merges) == 0) return(df)
  
  df <- pad_to_merges(df, merges)
  
  for (rng in merges) {
    parts <- strsplit(rng, ":", fixed = TRUE)[[1]]
    rc1 <- excel_ref_to_rc(parts[1])
    rc2 <- excel_ref_to_rc(ifelse(length(parts) == 2, parts[2], parts[1]))
    if (any(is.na(c(rc1, rc2)))) next
    
    r1 <- min(rc1[1], rc2[1]); r2 <- max(rc1[1], rc2[1])
    c1 <- min(rc1[2], rc2[2]); c2 <- max(rc1[2], rc2[2])
    
    val <- df[r1, c1]
    
    for (r in r1:r2) for (c in c1:c2) {
      cur <- df[r, c]
      if (is.na(cur) || str_trim(as.character(cur)) == "") df[r, c] <- val
    }
  }
  df
}


# First data row = row with >=2 numeric entries
# Heuristic:
# For each row, count how many cells look numeric
# First row with at least 2 numeric cells is treated as the start of the table

find_first_data_row <- function(mat, marker = "E92000001", search_cols = 3) {
  if (is.null(mat) || nrow(mat) == 0 || ncol(mat) == 0) return(NA_integer_)
  
  k <- min(search_cols, ncol(mat))
  
  hit <- apply(mat[, seq_len(k), drop = FALSE], 1, function(r) {
    any(!is.na(r) & str_trim(as.character(r)) == marker)
  })
  
  which(hit)[1]
}


# Last relevant column = numeric anywhere below OR header nonblank

find_last_col <- function(mat2, data_row, header_rows) {
  num_any <- apply(mat2[data_row:nrow(mat2), , drop = FALSE], 2,
                   function(col) any(is_fully_numeric(col)))
  
  hdr_any <- rep(FALSE, ncol(mat2))
  if (length(header_rows) > 0) {
    hdr_any <- apply(mat2[header_rows, , drop = FALSE], 2, function(col) {
      any(!is.na(col) & str_trim(as.character(col)) != "")
    })
  }
  
  keep <- which(num_any | hdr_any)
  if (length(keep) == 0) return(2L)
  max(2L, max(keep))
}


# Process a single sheet
# Process a single sheet
# CHANGE THE FUNCTION SIGNATURE:
process_sheet <- function(path, sheet_name, merges_for_sheet, hidden_info = list(cols = integer(0), rows = integer(0))) {
  raw <- suppressMessages(
    readxl::read_excel(
      path,
      sheet = sheet_name,
      col_names = FALSE,
      col_types = "text",
      .name_repair = "minimal"
    )
  )
  if (nrow(raw) < 3 || ncol(raw) < 2) return(NULL)
  
  raw <- fill_merged_cells_all(raw, merges_for_sheet)
  
  names(raw) <- paste0("X", seq_len(ncol(raw)))
  raw <- raw %>% mutate(across(everything(), as.character))
  mat <- as.matrix(raw)
  
  hidden_cols <- hidden_info$cols
  hidden_rows <- hidden_info$rows
  
  data_row <- find_first_data_row(mat, marker = "E92000001", search_cols = 3)
  if (is.na(data_row)) return(NULL)
  
  n_above <- data_row - 1L
  if (n_above <= 0) return(NULL)
  
  dataset_name <- get_dataset_name(mat, sheet_name)
  
  if (n_above == 1L) {
    header_rows <- integer(0)
  } else {
    header_rows <- 2L:(data_row - 1L)
  }
  
  if (length(hidden_rows) > 0 && length(header_rows) > 0) {
    header_rows <- header_rows[!header_rows %in% hidden_rows]
  }
  
  row_vals <- str_trim(as.character(mat[data_row, ]))
  start_col <- which(!is.na(row_vals) & row_vals == "E92000001")[1]
  if (is.na(start_col)) return(NULL)
  if (start_col >= ncol(mat)) return(NULL)
  
  last_col <- find_last_col(mat, data_row, header_rows)
  if (last_col < start_col + 1) return(NULL)
  col_idx <- start_col:last_col
  
  if (length(hidden_cols) > 0) {
    col_idx <- col_idx[!col_idx %in% hidden_cols]
    if (length(col_idx) < 2) return(NULL)
  }
  
  # Drop columns where ALL data rows are empty/NA
  data_mat <- mat[data_row:nrow(mat), col_idx, drop = FALSE]
  col_has_data <- apply(data_mat, 2, function(col) {
    any(!is.na(col) & str_trim(as.character(col)) != "")
  })
  col_idx <- col_idx[col_has_data]
  if (length(col_idx) < 2) return(NULL)
  
  # Drop sparse header rows: if only one unique non-empty value exists across
  # col_idx and it spans <= 2 columns, it's a sheet title / section label — remove it
  if (length(header_rows) > 0) {
    header_rows <- header_rows[vapply(header_rows, function(r) {
      vals <- str_trim(as.character(mat[r, col_idx]))
      vals_filled <- vals[!is.na(vals) & vals != ""]
      n_distinct <- length(unique(vals_filled))
      n_filled   <- length(vals_filled)
      !(n_distinct <= 1 && n_filled <= 2)
    }, logical(1))]
  }
  
  cn <- build_colnames(mat, header_rows, col_idx)
  col_names          <- cn$clean
  raw_col_names      <- cn$raw
  pre_unique_names   <- cn$clean_pre_unique
  
  if (length(col_names) >= 1) { col_names[1] <- "Region_code"; raw_col_names[1] <- "Region_code"; pre_unique_names[1] <- "Region_code" }
  if (length(col_names) >= 2) { col_names[2] <- "Region_name"; raw_col_names[2] <- "Region_name"; pre_unique_names[2] <- "Region_name" }
  col_names <- make.unique(col_names, sep = "_")
  
  col_map <- tibble(
    sheet              = sheet_name,
    dataset_name       = dataset_name,
    excel_col_n        = col_idx,
    excel_col          = excel_col_letters(col_idx),
    column_name        = col_names,
    raw_column_name    = raw_col_names,
    score_column_name  = pre_unique_names
  )
  
  df <- raw[data_row:nrow(raw), col_idx, drop = FALSE]
  colnames(df) <- col_names
  
  if (length(hidden_rows) > 0) {
    relative_hidden <- hidden_rows - data_row + 1
    relative_hidden <- relative_hidden[relative_hidden > 0 & relative_hidden <= nrow(df)]
    if (length(relative_hidden) > 0) {
      df <- df[-relative_hidden, , drop = FALSE]
    }
  }
  
  empty_rows <- (is.na(df[[1]]) | str_trim(as.character(df[[1]])) == "") &
    (is.na(df[[2]]) | str_trim(as.character(df[[2]])) == "")
  df <- df[!empty_rows, , drop = FALSE]
  
  df <- as_tibble(df)
  
  df <- df %>%
    mutate(across(everything(),
                  ~ ifelse(is_fully_numeric(.x), as.numeric(gsub(",", "", .x)), .x)))
  
  attr(df, "dataset_name") <- dataset_name
  attr(df, "sheet") <- sheet_name
  attr(df, "col_map") <- col_map
  
  df
}
read_excel_to_dataset_list <- function(path) {
  sheets <- readxl::excel_sheets(path)
  
  sheets_keep <- sheets[!(
    str_starts(tolower(sheets), "contents") |
      str_detect(tolower(sheets), "la\\s*dropdown")
  )]
  wanted <- unique(norm_sheet(keys$sheet))
  sheets_keep <- sheets_keep[norm_sheet(sheets_keep) %in% wanted]
  
  merges_by_sheet <- get_merge_ranges_from_xlsx(path)
  hidden_by_sheet <- get_hidden_cols_and_rows_from_xlsx(path)
  bad <- character(0)
  
  results <- map(sheets_keep, function(sh) {
    tryCatch(
      process_sheet(path, sh, merges_by_sheet[[sh]], hidden_by_sheet[[sh]]),
      error = function(e) {
        bad <<- c(bad, paste0(sh, "  -->  ", conditionMessage(e)))
        NULL
      }
    )
  })
  
  results <- compact(results)
  
  if (length(results) == 0) {
    col_map_all <- tibble(
      sheet = character(0), dataset_name = character(0),
      excel_col_n = integer(0), excel_col = character(0),
      column_name = character(0), raw_column_name = character(0)
    )
    attr(results, "col_map_all") <- col_map_all
    if (length(bad) > 0) message("Skipped sheets due to errors:\n", paste(bad, collapse = "\n"))
    return(results)
  }
  
  nm <- map_chr(results, ~ attr(.x, "sheet"))
  names(results) <- base::make.unique(nm, sep = "_")
  
  col_map_all <- dplyr::bind_rows(purrr::map(results, ~ attr(.x, "col_map"))) %>%
    dplyr::mutate(
      column_name     = as.character(.data$column_name),
      raw_column_name = as.character(.data$raw_column_name)
    )
  
  attr(results, "col_map_all") <- col_map_all
  if (length(bad) > 0) message("Skipped sheets due to errors:\n", paste(bad, collapse = "\n"))
  
  results
}

norm_sheet <- function(x) {
  x <- tolower(as.character(x))
  x <- str_trim(x)
  # remove all non-alphanumeric characters (dashes, spaces, brackets, etc., like "TA1_" instead of "TA1")
  x <- str_replace_all(x, "[^a-z0-9]+", "")
  x
}

norm_txt <- function(x) {
  x <- paste(as.character(x), collapse = " ")
  x <- tolower(x)
  x <- str_replace_all(x, "[_\\-]", " ")
  x <- str_replace_all(x, "[\r\n\t]", " ")
  x <- str_replace_all(x, "\\+", " + ")
  x <- str_replace_all(x, "[^a-z0-9+/#]+", " ")
  str_squish(x)
}

tokenize_words <- function(x) {
  # Strip footnote digits from raw input before normalising
  x <- str_replace_all(paste(as.character(x), collapse = " "), 
                       "(?<=[a-zA-Z])(\\d+)(?=[^a-zA-Z0-9]|$)", "")
  w <- unlist(strsplit(norm_txt(x), " ", fixed = TRUE))
  w <- w[w != ""]
  stop <- c("the","and","or","of","to","with","as","per","days","day",
            "in","on","for","a","an")
  w <- w[!w %in% stop]
  unique(w)
}

strip_footnote_digits <- function(x) {
  # Match: letter immediately followed by digits, where digits are NOT followed by a letter or hyphen
  str_replace_all(x, "(?<=[a-zA-Z])(\\d+)(?=[^a-zA-Z0-9]|$)", "")
}


score_overlap <- function(phrase, cand) {
  kw <- tokenize_words(phrase)
  cw <- tokenize_words(cand)
  if (length(kw) == 0) return(0)
  sum(kw %in% cw) / length(kw)
}

score_jw <- function(phrase, cand) {
  stringdist::stringsim(norm_txt(phrase), norm_txt(cand), method = "jw", p = 0.1)
}

best_phrase_score <- function(phrases, cand) {
  scores <- vapply(phrases, function(ph) {
    score_combined(ph, cand)  # Token overlap + n-gram overlap
  }, FUN.VALUE = numeric(1))
  max(scores, na.rm = TRUE)
}


match_keys_to_colmap_names_only <- function(keys, col_map_all,
                                            min_score = min_score,
                                            return_debug = FALSE) {
  req <- c("key_id", "sheet", "phrases", "expected_excel_cols")
  if (!all(req %in% names(keys))) stop("keys must contain columns: ", paste(req, collapse = ", "))
  stopifnot(all(c("sheet", "excel_col", "column_name", "raw_column_name") %in% names(col_map_all)))
  
  col_map_all <- col_map_all %>%
    mutate(
      sheet             = as.character(.data$sheet),
      sheet_norm        = norm_sheet(.data$sheet),
      excel_col         = as.character(.data$excel_col),
      column_name       = as.character(.data$column_name),
      raw_column_name   = as.character(.data$raw_column_name),
      score_column_name = if ("score_column_name" %in% names(.))
        as.character(.data$score_column_name)
      else
        as.character(.data$raw_column_name)
    )
  
  purrr::map_dfr(seq_len(nrow(keys)), function(i) {
    k <- keys[i, ]
    
    sheet_cands <- col_map_all %>%
      filter(.data$sheet_norm == norm_sheet(k$sheet))
    
    if (nrow(sheet_cands) == 0) {
      return(tibble(
        key_id = k$key_id, sheet = as.character(k$sheet),
        expected_excel_col = paste(unlist(k$expected_excel_cols), collapse = "+"),
        matched_excel_col = NA_character_, matched_column_name = NA_character_,
        matched_phrase = NA_character_, score = NA_real_, col_ok = NA
      ))
    }
    
    best_phrase_hit <- function(phrases, cand) {
      scores <- vapply(phrases, function(ph) score_combined(ph, cand), FUN.VALUE = numeric(1))
      j <- which.max(scores)
      list(best_phrase = phrases[[j]], best_score = scores[[j]])
    }
    
    # Score against score_column_name (pre-uniquified)
    scored <- sheet_cands %>%
      mutate(
        key_id         = k$key_id,
        hit            = lapply(.data$score_column_name, function(cand) best_phrase_hit(unlist(k$phrases), cand)),
        score          = vapply(.data$hit, function(h) h$best_score, numeric(1)),
        matched_phrase = vapply(.data$hit, function(h) h$best_phrase, character(1))
      ) %>%
      select(-.data$hit) %>%
      arrange(desc(.data$score))
    
    # Get the best score
    top_score <- scored %>%
      filter(!is.na(.data$score), .data$score >= (min_score %||% 0)) %>%
      pull(score) %>%
      max(na.rm = TRUE)
    
    if (is.infinite(top_score) || is.na(top_score)) {
      return(tibble(
        key_id = k$key_id, sheet = as.character(k$sheet),
        expected_excel_col = paste(unlist(k$expected_excel_cols), collapse = "+"),
        matched_excel_col = NA_character_, matched_column_name = NA_character_,
        matched_phrase = NA_character_, score = NA_real_, col_ok = NA
      ))
    }
    
    # Keep ALL rows that share the top score AND the same score_column_name
    # as the top scorer — these are the count+pct sibling columns
    top_score_name <- scored %>%
      filter(score == top_score) %>%
      pull(score_column_name) %>%
      first()
    
    best <- scored %>%
      filter(score == top_score, score_column_name == top_score_name)
    
    expected_set <- unlist(k$expected_excel_cols)
    best %>%
      transmute(
        key_id              = k$key_id,
        sheet               = as.character(k$sheet),
        expected_excel_col  = paste(unlist(k$expected_excel_cols), collapse = "+"),
        matched_excel_col   = .data$excel_col,
        matched_column_name = .data$column_name,
        matched_phrase      = .data$matched_phrase,
        score               = .data$score,
        col_ok              = matched_excel_col %in% expected_set
      )
  })
}

# Keys
key_row <- function(key_id, sheet, excel_col, header_patterns, label = NA_character_) {
  tibble(
    key_id              = key_id,
    sheet               = sheet,
    expected_excel_cols = list(excel_col),
    header_patterns     = list(header_patterns),
    phrases             = list(header_patterns),
    label               = label
  )
}

source(file.path(base_dir, "keys.R"))  #<-- change to actual filename

keys <- keys %>% mutate(sheet = norm_sheet(sheet))


# Run matching PER FILE (within-file / within-sheet)
file_id_from_path <- function(path) {
  basename(path) |> str_replace("\\.xlsx$", "")
}

# GitHub API URL to list directory contents
github_api_url <- "https://api.github.com/repos/homelessnessimpact/England/contents/MHCLG%20Raw%20Data/Statutory%20Homelessness/Quarterly%20Releases"

# Fetch directory listing from GitHub API
repo_contents <- fromJSON(github_api_url)

# Filter to only .xlsx files matching the expected naming pattern
xlsx_filenames <- repo_contents %>%
  filter(type == "file") %>%
  pull(name) %>%
  keep(~ grepl("^\\d{2} Q[1-4].*\\.xlsx$", .x, ignore.case = TRUE))

# Build download URLs from the filtered filenames
github_base <- "https://raw.githubusercontent.com/homelessnessimpact/England/main/MHCLG%20Raw%20Data/Statutory%20Homelessness/Quarterly%20Releases"


# Download each file to a temp directory and build file_paths
temp_dir <- file.path(tempdir(), "chi_xlsx")
dir.create(temp_dir, showWarnings = FALSE)

file_paths <- vapply(xlsx_filenames, function(fname) {
  dest <- file.path(temp_dir, fname)
  if (!file.exists(dest)) {
    url <- paste0(github_base, "/", utils::URLencode(fname, repeated = TRUE))
    utils::download.file(url, destfile = dest, mode = "wb", quiet = TRUE)
  }
  dest
}, FUN.VALUE = character(1))



# Two-iteration matching with claiming

run_matching_for_file <- function(path, keys, min_score = min_score) {
  datasets <- read_excel_to_dataset_list(path)
  col_map_all <- attr(datasets, "col_map_all")
  file_id <- file_id_from_path(path)
  
  # --- ITERATION 1: match all keys to all columns ---
  iter1 <- match_keys_to_colmap_names_only(
    keys = keys,
    col_map_all = col_map_all,
    min_score = min_score,
    return_debug = FALSE
  ) %>%
    mutate(file_path = path, file_id = file_id, iteration = 1L)
  
  # For each COLUMN that matched more than one key, keep only the top-scoring key.
  # If a column has tied top scores across multiple keys, keep all (flagged as tied).
  iter1_resolved <- iter1 %>%
    filter(!is.na(matched_column_name), !is.na(score), score >= min_score) %>%
    group_by(file_id, sheet, matched_column_name) %>%
    mutate(
      col_max_score = max(score, na.rm = TRUE),
      col_n_at_max  = sum(score == col_max_score, na.rm = TRUE)
    ) %>%
    # Keep only the top-scoring key(s) per column
    filter(score == col_max_score) %>%
    ungroup()
  
  # Columns that were claimed (by exactly one key, or tied — both are claimed)
  claimed_cols <- iter1_resolved %>%
    distinct(sheet, matched_column_name)
  
  # Keys that got a match in iter1 (don't re-run in iter2)
  claimed_keys <- iter1_resolved %>%
    distinct(sheet, key_id)
  
  # NA rows from iter1 (keys with no match at all) — carry forward as-is
  iter1_na <- iter1 %>%
    filter(is.na(matched_column_name) | is.na(score) | score < min_score) %>%
    group_by(sheet, key_id) %>%
    slice(1) %>%
    ungroup()
  
  # --- ITERATION 2: remaining keys against unclaimed columns ---
  remaining_keys <- keys %>%
    anti_join(claimed_keys, by = c("sheet", "key_id"))
  
  remaining_col_map <- col_map_all %>%
    mutate(sheet_norm = norm_sheet(sheet)) %>%
    anti_join(
      claimed_cols %>% mutate(sheet_norm = norm_sheet(sheet)),
      by = c("sheet_norm" = "sheet", "column_name" = "matched_column_name")
    ) %>%
    select(-sheet_norm)
  
  if (nrow(remaining_keys) > 0 && nrow(remaining_col_map) > 0) {
    iter2 <- match_keys_to_colmap_names_only(
      keys = remaining_keys,
      col_map_all = remaining_col_map,
      min_score = min_score,
      return_debug = FALSE
    ) %>%
      mutate(file_path = path, file_id = file_id, iteration = 2L)
  } else {
    iter2 <- tibble(
      key_id = character(0), sheet = character(0),
      expected_excel_col = character(0), matched_excel_col = character(0),
      matched_column_name = character(0), matched_phrase = character(0),
      score = numeric(0), col_ok = logical(0),
      file_path = character(0), file_id = character(0), iteration = integer(0)
    )
  }
  
  bind_rows(iter1_resolved, iter1_na, iter2) %>%
    arrange(sheet, key_id, iteration)
}

matches_by_file <- setNames(
  lapply(file_paths, run_matching_for_file, keys = keys, min_score = min_score),
  file_id_from_path(file_paths)
)


# Deduplicate: keep highest score per (file_id, sheet, key_id), flag ties
all_key_matches <- bind_rows(matches_by_file, .id = "file_id_dup") %>%
  select(-file_id_dup) %>%
  relocate(file_id, file_path, iteration, .before = 1) %>%
  group_by(file_id, sheet, matched_column_name) %>%
  mutate(
    col_claimed_by_n_keys = sum(!is.na(score) & score >= min_score, na.rm = TRUE),
    col_tie = col_claimed_by_n_keys > 1
  ) %>%
  ungroup() %>%
  arrange(file_id, sheet, key_id, iteration, desc(score))

all_key_matches

problems_by_file <- all_key_matches %>%
  filter(
    is.na(matched_column_name) |
      isFALSE(col_ok) |
      col_tie == TRUE
  ) %>%
  arrange(file_id, sheet, key_id, iteration)

problems_by_file


####################################################################################################################################
####################################################################################################################################
####################################################################################################################################
####################################################################################################################################



# Rename column names using the matched keys
datasets_by_file <- setNames(
  lapply(file_paths, read_excel_to_dataset_list),
  file_id_from_path(file_paths)
)

# 1) plan: old column name -> new key_id
rename_plan <- all_key_matches %>%
  filter(!is.na(matched_column_name), !is.na(score), score >= min_score) %>%
  group_by(file_id, sheet, matched_column_name) %>%
  slice_max(order_by = score, n = 1, with_ties = TRUE) %>%  # keep ties (count+pct siblings)
  ungroup() %>%
  transmute(file_id, sheet, old = matched_column_name, new = key_id)


# Duplicate key handling: _count / _percentage / _duplicate3...

is_pct_like <- function(x) {
  x_char <- str_trim(as.character(x))
  has_pct_sign <- any(str_detect(x_char[!is.na(x_char) & x_char != ""], "%"), na.rm = TRUE)
  if (has_pct_sign) return(TRUE)
  suppressWarnings(nx <- as.numeric(gsub("[,%]", "", x_char)))
  nx <- nx[!is.na(nx)]
  if (length(nx) < 3) return(FALSE)
  max_val <- max(nx, na.rm = TRUE)
  # All values between 0 and 1 inclusive → must be a proportion column
  if (max_val <= 1 && all(nx >= 0)) return(TRUE)
  # Majority 0-1 with decimals → proportion
  in_0_1  <- mean(nx >= 0 & nx <= 1)
  has_dec <- mean(abs(nx - round(nx)) > 1e-6)
  (in_0_1 > 0.85) && (has_dec > 0.15)
}

# Storage for non-pct ties found during renaming (populated below)
non_pct_ties <- list()


# Remove duplicate words from column names


# 3) apply across files/sheets - use CLEANED column names, not key_ids
datasets_renamed_by_file <- imap(datasets_by_file, function(file_datasets, fid) {
  
  old_names <- names(file_datasets)
  names(file_datasets) <- norm_sheet(old_names)
  
  plan_file <- rename_plan %>%
    filter(file_id == fid) %>%
    mutate(sheet = norm_sheet(sheet))
  
  if (nrow(plan_file) == 0) return(file_datasets)
  
  imap(file_datasets, function(df, sh) {
    if (is.null(df) || nrow(df) == 0) return(df)
    
    plan_sheet <- plan_file %>% filter(sheet == sh)
    if (nrow(plan_sheet) == 0) return(df)
    
    # Step 1: Keep only Region_code, Region_name, and matched columns
    keep_cols <- c("Region_code", "Region_name", plan_sheet$old)
    keep_cols <- intersect(keep_cols, names(df))
    df_subset <- df[, keep_cols, drop = FALSE]
    
    # Step 2: Rename matched columns old -> key_id
    current_names <- names(df_subset)
    new_col_names <- current_names
    for (i in seq_along(current_names)) {
      if (current_names[i] %in% c("Region_code", "Region_name")) next
      match_idx <- which(plan_sheet$old == current_names[i])
      if (length(match_idx) > 0) {
        new_col_names[i] <- plan_sheet$new[match_idx[1]]
      }
    }
    names(df_subset) <- new_col_names
    
# Step 3: Handle duplicate key_ids
    # - Pct columns (% sign or 0-1 proportions): rename as key_id_pct1, key_id_pct2, etc.
    # - Count duplicates: keep most-complete, drop rest
    # - Non-pct ties that aren't explained by pct detection: flag for review
    final_names <- new_col_names
    
    for (name in unique(new_col_names)) {
      if (name %in% c("Region_code", "Region_name")) next
      
      idx <- which(final_names == name)
      if (length(idx) <= 1) next
      
      pct_flags <- vapply(idx, function(j) is_pct_like(df_subset[[j]]), logical(1))
      
      pct_idx   <- idx[pct_flags]
      count_idx <- idx[!pct_flags]
      
      # For count duplicates: keep most-complete, drop rest
      if (length(count_idx) > 1) {
        non_na <- vapply(count_idx, function(j) {
          sum(!is.na(df_subset[[j]]) & str_trim(as.character(df_subset[[j]])) != "")
        }, integer(1))
        keep <- count_idx[which.max(non_na)]
        drop <- count_idx[count_idx != keep]
        final_names[drop] <- paste0(".DROP.", seq_along(drop), ".", name)
        
        # Flag: multiple non-pct columns tied to same key_id — may need review
        non_pct_ties[[length(non_pct_ties) + 1]] <<- tibble(
          file_id    = fid,
          sheet      = sh,
          key_id     = name,
          n_count_dupes = length(count_idx),
          kept_col   = names(df_subset)[keep],
          dropped_cols = paste(names(df_subset)[drop], collapse = ", ")
        )
      }
      
      # For pct columns: rename each as key_id_pct1, key_id_pct2, etc.
      if (length(pct_idx) >= 1) {
        for (k in seq_along(pct_idx)) {
          final_names[pct_idx[k]] <- paste0(name, "_pct", k)
        }
      }
    }
    
    names(df_subset) <- final_names
    df_subset <- df_subset[, !str_detect(names(df_subset), "^\\.DROP\\."), drop = FALSE]
    
    # Post-rename sanity check: non-_pct columns containing proportion values
    # indicate a wrong match (pct column matched to count key) — replace with NA
    for (cn in names(df_subset)) {
      if (cn %in% c("Region_code", "Region_name")) next
      if (str_detect(cn, "_pct\\d*$")) next
      
      vals <- suppressWarnings(as.numeric(gsub("[,%]", "", as.character(df_subset[[cn]]))))
      vals_nn <- vals[!is.na(vals)]
      if (length(vals_nn) < 3) next
      
      in_0_1  <- mean(vals_nn > 0 & vals_nn <= 1)
      has_dec <- mean(abs(vals_nn - round(vals_nn)) > 1e-6)
      
      if (in_0_1 > 0.85 && has_dec > 0.15) {
        message("Proportion values in count column '", cn, "' [", fid, " / ", sh, "] — replacing with NA")
        df_subset[[cn]] <- NA_real_
      }
    }
    
    df_subset
  })
})
  
non_pct_ties_df <- bind_rows(non_pct_ties)
if (nrow(non_pct_ties_df) > 0) {
  message("Non-pct ties found (count columns dropped):")
  print(non_pct_ties_df, n = Inf)
} else {
  message("No non-pct ties found.")
}

# Append data
file_id_to_year_quarter <- function(file_id) {
  # New format: "19 Q1 Jan-Mar Statutory homelessness in England"
  m_new <- str_match(file_id, "^(\\d{2})\\s+Q([1-4])\\b")
  if (!any(is.na(m_new))) {
    yy <- as.integer(m_new[2])
    # Assume 2000s: 19 -> 2019, etc.
    yyyy <- ifelse(yy >= 0 & yy <= 99, 2000L + yy, yy)
    q <- as.integer(m_new[3])
    return(sprintf("%d_Q%d", yyyy, q))
  }
  
  # Old format: "Detailed_LA_202409"
  m_old <- str_match(file_id, "(\\d{4})(\\d{2})$")
  if (!any(is.na(m_old))) {
    yy <- as.integer(m_old[2])
    mm <- as.integer(m_old[3])
    q <- dplyr::case_when(
      mm == 3  ~ 1L,
      mm == 6  ~ 2L,
      mm == 9  ~ 3L,
      mm == 12 ~ 4L,
      TRUE     ~ NA_integer_
    )
    if (is.na(q)) return(NA_character_)
    return(sprintf("%d_Q%d", yy, q))
  }
  
  NA_character_
}

all_sheets <- sort(unique(unlist(map(datasets_renamed_by_file, names))))

combined_by_sheet <- setNames(vector("list", length(all_sheets)), all_sheets)

for (sh in all_sheets) {
  
  dfs <- imap(datasets_renamed_by_file, function(file_datasets, fid) {
    df <- file_datasets[[sh]]
    if (is.null(df)) return(NULL)
    
    df %>%
      mutate(
        year_quarter = file_id_to_year_quarter(fid),
        file_id = fid
      ) %>%
      relocate(year_quarter, file_id)
  })
  
  combined_by_sheet[[sh]] <- bind_rows(compact(dfs))
}
kid_cols  <- c("year_quarter", "file_id", "Region_code", "Region_name")

# Keep all columns (no pattern matching needed since we already filtered to matched columns)
combined_by_sheet <- map(combined_by_sheet, ~ .x %>%
                           filter(
                             !is.na(Region_code),
                             str_detect(Region_code, "^E[0-9]+$")
                           ) %>%
                           select(any_of(kid_cols), everything()) %>%  # CHANGED: id_cols -> kid_cols
                           arrange(year_quarter, Region_code, Region_name)
)



# POST-PROCESSING: Derived columns (sums and percentages)

a2p_rent_arrears_cols <- c(
  "rent_arrears_tenant_difficulty_budgeting",
  "rent_arrears_increase_in_rent",
  "rent_arrears_reduction_employment_income",
  "rent_arrears_changes_benefit_entitlement",
  "rent_arrears_change_personal_circumstances"
)

# --- A2P: Sum of Hospital columns ---
a2p_hospital_cols <- c(
  "hospital_psychiatric",
  "hospital_general"
)

if (!is.null(combined_by_sheet[["a2p"]])) {
  combined_by_sheet[["a2p"]] <- combined_by_sheet[["a2p"]] %>%
    mutate(
      SUM_A2P_rent_arrears = rowSums(
        across(any_of(a2p_rent_arrears_cols), ~ suppressWarnings(as.numeric(.x))),
        na.rm = TRUE
      ),
      SUM_A2P_hospital = rowSums(
        across(any_of(a2p_hospital_cols), ~ suppressWarnings(as.numeric(.x))),
        na.rm = TRUE
      )
    ) %>%
    # Replace 0 sums with NA where ALL contributing cols were NA
    mutate(
      SUM_A2P_rent_arrears = ifelse(
        rowSums(!is.na(across(any_of(a2p_rent_arrears_cols)))) == 0,
        NA_real_, SUM_A2P_rent_arrears
      ),
      SUM_A2P_hospital = ifelse(
        rowSums(!is.na(across(any_of(a2p_hospital_cols)))) == 0,
        NA_real_, SUM_A2P_hospital
      )
    )
}

# --- A2R: Sum of Rent Arrears columns ---

a2r_rent_arrears_cols <- c(
  "relief_rent_arrears_budgeting",
  "relief_rent_arrears_increase",
  "relief_rent_arrears_income_reduction",
  "relief_rent_arrears_benefit_change",
  "relief_rent_arrears_personal_circs"
)

# --- A2R: Sum of Hospital columns ---
a2r_hospital_cols <- c(
  "relief_hospital_psych",
  "relief_hospital_gen"
)

if (!is.null(combined_by_sheet[["a2r"]])) {
  combined_by_sheet[["a2r"]] <- combined_by_sheet[["a2r"]] %>%
    mutate(
      SUM_A2R_rent_arrears = rowSums(
        across(any_of(a2r_rent_arrears_cols), ~ suppressWarnings(as.numeric(.x))),
        na.rm = TRUE
      ),
      SUM_A2R_hospital = rowSums(
        across(any_of(a2r_hospital_cols), ~ suppressWarnings(as.numeric(.x))),
        na.rm = TRUE
      )
    ) %>%
    mutate(
      SUM_A2R_rent_arrears = ifelse(
        rowSums(!is.na(across(any_of(a2r_rent_arrears_cols)))) == 0,
        NA_real_, SUM_A2R_rent_arrears
      ),
      SUM_A2R_hospital = ifelse(
        rowSums(!is.na(across(any_of(a2r_hospital_cols)))) == 0,
        NA_real_, SUM_A2R_hospital
      )
    )
}

# --- R1: Sum of "Total other" columns ---

r1_other_cols <- c(
  "r1_intentional_homeless",
  "r1_refused_final_acc",
  "r1_refused_cooperate",
  "r1_contact_lost",
  "r1_withdrew_deceased",
  "r1_no_longer_eligible",
  "r1_not_known"
)

if (!is.null(combined_by_sheet[["r1"]])) {
  combined_by_sheet[["r1"]] <- combined_by_sheet[["r1"]] %>%
    mutate(
      SUM_R1_total_other = rowSums(
        across(any_of(r1_other_cols), ~ suppressWarnings(as.numeric(.x))),
        na.rm = TRUE
      )
    ) %>%
    mutate(
      SUM_R1_total_other = ifelse(
        rowSums(!is.na(across(any_of(r1_other_cols)))) == 0,
        NA_real_, SUM_R1_total_other
      )
    )
}

# --- A6: Percentage columns (all cols / Total owed a prevention or relief duty) ---

a6_total_col  <- "a6_total_owed_relief_prevention_duty"
a6_other_cols <- c(
  "a6_16_17", "a6_18_24", "a6_25_34", "a6_35_44",
  "a6_45_54", "a6_55_64", "a6_65_74", "a6_75_plus", "a6_not_known"
)

if (!is.null(combined_by_sheet[["a6"]]) && a6_total_col %in% names(combined_by_sheet[["a6"]])) {
  
  a6_cols_present <- intersect(a6_other_cols, names(combined_by_sheet[["a6"]]))
  
  combined_by_sheet[["a6"]] <- combined_by_sheet[["a6"]] %>%
    mutate(across(
      all_of(a6_cols_present),
      ~ suppressWarnings(as.numeric(.x)) / suppressWarnings(as.numeric(.data[[a6_total_col]])),
      .names = "PCT_{.col}"
    ))
}
