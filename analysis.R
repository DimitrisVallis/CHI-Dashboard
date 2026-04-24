
#1.SOURCE THE PIPELINE
base_dir <- dirname(sys.frame(1)$ofile)
source(file.path(base_dir, "pipeline.R"))

message("Data pipeline complete.")

# Set options
extra_covariate <- NULL
reference_period <- NULL
plot_type <- readline(prompt = "Enter plot type - type 'predicted' or 'estimates' and press Enter: ")
while (!plot_type %in% c("predicted", "estimates")) {
  message("Invalid choice. Please type 'predicted' or 'estimates'.")
  plot_type <- readline(prompt = "Enter plot type - type 'predicted' or 'estimates' and press Enter: ")
}
message("Running analysis with plot type: ", plot_type)


source(file.path(base_dir, "keys.R"))  # <-- change to actual filename
keys <- keys %>% mutate(sheet = norm_sheet(sheet))


#2.DEFINE OUTCOMES TO ANALYSE
# Derived outcomes not in keys (computed columns: SUM_, PCT_, etc.)
derived_outcomes <- list(
  list(sheet = "a2p", outcome = "SUM_A2P_rent_arrears", label = "Sum rent arrears"),
  list(sheet = "a2p", outcome = "SUM_A2P_hospital",     label = "Sum hospital"),
  list(sheet = "a2r", outcome = "SUM_A2R_rent_arrears", label = "Sum rent arrears"),
  list(sheet = "a2r", outcome = "SUM_A2R_hospital",     label = "Sum hospital"),
  list(sheet = "a6",  outcome = "PCT_a6_16_17",         label = "% Age 16-17"),
  list(sheet = "a6",  outcome = "PCT_a6_18_24",         label = "% Age 18-24"),
  list(sheet = "a6",  outcome = "PCT_a6_25_34",         label = "% Age 25-34"),
  list(sheet = "a6",  outcome = "PCT_a6_35_44",         label = "% Age 35-44"),
  list(sheet = "a6",  outcome = "PCT_a6_45_54",         label = "% Age 45-54"),
  list(sheet = "a6",  outcome = "PCT_a6_55_64",         label = "% Age 55-64"),
  list(sheet = "a6",  outcome = "PCT_a6_65_74",         label = "% Age 65-74"),
  list(sheet = "a6",  outcome = "PCT_a6_75_plus",       label = "% Age 75+"),
  list(sheet = "a6",  outcome = "PCT_a6_not_known",     label = "% Not known"),
  list(sheet = "r1",  outcome = "SUM_R1_total_other",   label = "Sum total other")
)

outcomes_to_run <- keys %>%
  filter(!is.na(label)) %>%
  mutate(sheet = tolower(sheet)) %>%
  pmap(function(key_id, sheet, label, ...) {
    list(sheet = sheet, outcome = key_id, label = label)
  }) %>%
  c(derived_outcomes)
  



#HELPER: cluster-robust SE
cluster_se <- function(model, cluster_var) {
  vcov_cl <- sandwich::vcovCL(model, cluster = cluster_var)
  lmtest::coeftest(model, vcov = vcov_cl)
}


#HELPER: predicted means regression
run_one_regression <- function(sheet_name, outcome_col, extra_cov = NULL,
                               ref_period = NULL, data_list = combined_by_sheet) {
  df <- data_list[[sheet_name]]
  if (is.null(df) || nrow(df) == 0) { message("Sheet not found or empty: ", sheet_name); return(NULL) }
  if (!outcome_col %in% names(df)) { message("Outcome '", outcome_col, "' not found in '", sheet_name, "'"); return(NULL) }
  if (!is.null(extra_cov) && !extra_cov %in% names(df)) { message("Covariate not found, dropping."); extra_cov <- NULL }
  
  excluded_region_codes <- c(
    "E92000001", "E12000007", "-",
    "E12000001", "E12000002", "E12000003", "E12000004", "E12000005",
    "E12000006", "E12000007", "E12000008", "E12000009"
  )
  
  df <- df %>%
    mutate(
      across(all_of(outcome_col), ~ suppressWarnings(as.numeric(gsub(",", "", as.character(.x))))),
      across(any_of(extra_cov),   ~ suppressWarnings(as.numeric(gsub(",", "", as.character(.x)))))
    ) %>%
    filter(!is.na(.data[[outcome_col]]), !is.infinite(.data[[outcome_col]]),
           !is.na(Region_code), !is.na(year_quarter),
           !Region_code %in% excluded_region_codes)
  
  if (nrow(df) < 10) { message("Too few obs: ", sheet_name, "/", outcome_col); return(NULL) }
  
  all_periods <- sort(unique(df$year_quarter))
  ref <- if (!is.null(ref_period) && ref_period %in% all_periods) ref_period else all_periods[1]
  df  <- df %>% mutate(time = relevel(factor(year_quarter), ref = ref))
  
  rhs   <- if (!is.null(extra_cov)) paste("time +", extra_cov) else "time"
  fml   <- as.formula(paste0("`", outcome_col, "` ~ ", rhs))
  model <- tryCatch(lm(fml, data = df), error = function(e) { message("Model failed: ", e$message); NULL })
  if (is.null(model)) return(NULL)
  
  vcov_cl <- tryCatch(
    sandwich::vcovCL(model, cluster = df$Region_code),
    error = function(e) { message("Clustering failed, using OLS vcov"); vcov(model) }
  )
  
  pred_grid <- data.frame(time = factor(all_periods, levels = levels(df$time)))
  if (!is.null(extra_cov)) pred_grid[[extra_cov]] <- mean(df[[extra_cov]], na.rm = TRUE)
  
  mm <- model.matrix(
    as.formula(paste0("~ time", if (!is.null(extra_cov)) paste0(" + ", extra_cov) else "")),
    data = pred_grid
  )
  
  common_cols <- intersect(colnames(mm), colnames(vcov_cl))
  mm       <- mm[, common_cols, drop = FALSE]
  vcov_sub <- vcov_cl[common_cols, common_cols]
  coef_vec <- coef(model)[common_cols]
  
  pred_mean <- as.vector(mm %*% coef_vec)
  pred_se   <- sqrt(pmax(diag(mm %*% vcov_sub %*% t(mm)), 0))
  
  tibble(
    period    = factor(all_periods, levels = all_periods),
    estimate  = pred_mean,
    std_error = pred_se,
    ci_low    = pred_mean - 1.96 * pred_se,
    ci_high   = pred_mean + 1.96 * pred_se,
    p_value   = NA_real_,
    sheet     = sheet_name,
    outcome   = outcome_col
  )
}


#HELPER: coefficient estimates regression
run_one_regression_estimates <- function(sheet_name, outcome_col, extra_cov = NULL,
                                         ref_period = NULL, data_list = combined_by_sheet) {
  df <- data_list[[sheet_name]]
  if (is.null(df) || nrow(df) == 0) { message("Sheet not found or empty: ", sheet_name); return(NULL) }
  if (!outcome_col %in% names(df)) { message("Outcome '", outcome_col, "' not found in '", sheet_name, "'"); return(NULL) }
  if (!is.null(extra_cov) && !extra_cov %in% names(df)) { message("Covariate not found, dropping."); extra_cov <- NULL }
  
  excluded_region_codes <- c(
    "E92000001", "E12000007", "-",
    "E12000001", "E12000002", "E12000003", "E12000004", "E12000005",
    "E12000006", "E12000007", "E12000008", "E12000009"
  )
  
  df <- df %>%
    mutate(
      across(all_of(outcome_col), ~ suppressWarnings(as.numeric(gsub(",", "", as.character(.x))))),
      across(any_of(extra_cov),   ~ suppressWarnings(as.numeric(gsub(",", "", as.character(.x)))))
    ) %>%
    filter(!is.na(.data[[outcome_col]]), !is.infinite(.data[[outcome_col]]),
           !is.na(Region_code), !is.na(year_quarter),
           !Region_code %in% excluded_region_codes)
  
  if (nrow(df) < 10) { message("Too few obs: ", sheet_name, "/", outcome_col); return(NULL) }
  
  all_periods <- sort(unique(df$year_quarter))
  ref <- if (!is.null(ref_period) && ref_period %in% all_periods) ref_period else all_periods[1]
  df  <- df %>% mutate(time = relevel(factor(year_quarter), ref = ref))
  
  rhs   <- if (!is.null(extra_cov)) paste("time +", extra_cov) else "time"
  fml   <- as.formula(paste0("`", outcome_col, "` ~ ", rhs))
  model <- tryCatch(lm(fml, data = df), error = function(e) { message("Model failed: ", e$message); NULL })
  if (is.null(model)) return(NULL)
  
  vcov_cl <- tryCatch(
    sandwich::vcovCL(model, cluster = df$Region_code),
    error = function(e) { message("Clustering failed, using OLS vcov"); vcov(model) }
  )
  
  ct <- lmtest::coeftest(model, vcov = vcov_cl)
  
  coef_df <- data.frame(
    term      = rownames(ct),
    estimate  = ct[, 1],
    std_error = ct[, 2],
    t_stat    = ct[, 3],
    p_value   = ct[, 4],
    row.names = NULL,
    stringsAsFactors = FALSE
  )
  
  time_rows <- coef_df %>%
    filter(str_detect(term, "^time")) %>%
    mutate(
      period  = str_replace(term, "^time", ""),
      ci_low  = estimate - 1.96 * std_error,
      ci_high = estimate + 1.96 * std_error,
      sheet   = sheet_name,
      outcome = outcome_col
    ) %>%
    select(term, estimate, std_error, p_value, period, ci_low, ci_high, sheet, outcome)
  
  ref_row <- tibble(
    term = paste0("time", ref), estimate = 0, std_error = 0,
    p_value = NA_real_, period = ref,
    ci_low = 0, ci_high = 0, sheet = sheet_name, outcome = outcome_col
  )
  
  bind_rows(ref_row, time_rows) %>%
    mutate(period = factor(period, levels = all_periods))
}

#RUN REGRESSIONS
run_fn <- if (plot_type == "predicted") run_one_regression else run_one_regression_estimates

results_list <- map(outcomes_to_run, function(x) {
  run_fn(sheet_name  = x$sheet,
         outcome_col = x$outcome,
         extra_cov   = extra_covariate,
         ref_period  = reference_period)
})

results_labelled <- map2(results_list, outcomes_to_run, function(res, spec) {
  if (is.null(res)) return(NULL)
  res %>% mutate(label = spec$label, sheet = spec$sheet)
})

marginal_effects <- bind_rows(compact(results_labelled))
if (nrow(marginal_effects) == 0) stop("No results to plot.")


#PLOT: one faceted grid per sheet, estimates annotated at each point

y_label <- if (plot_type == "predicted") "Predicted mean (95% CI)" else "Coefficient vs. reference period (95% CI)"

subtitle_txt <- paste0(
  if (plot_type == "estimates") paste0("Reference: ", if (!is.null(reference_period)) reference_period else "earliest available", "  |  ") else "",
  if (!is.null(extra_covariate)) paste0("Controlling for: ", extra_covariate) else ""
)

sheets_in_results <- unique(marginal_effects$sheet)

plots <- setNames(map(sheets_in_results, function(sh) {
  
  df_plot <- marginal_effects %>%
    filter(sheet == sh) %>%
    mutate(
      point_label = formatC(round(estimate, 2), format = "f", digits = 2),
      significant = if (plot_type == "estimates") (!is.na(p_value) & p_value < 0.05) else FALSE
    )
  
  ggplot(df_plot, aes(x = period, y = estimate, group = label)) +
    { if (plot_type == "estimates") geom_hline(yintercept = 0, linetype = "dashed", colour = "grey50", linewidth = 0.4) } +
    geom_ribbon(aes(ymin = ci_low, ymax = ci_high), alpha = 0.15, fill = "#2166ac") +
    geom_line(colour = "#2166ac", linewidth = 0.7) +
    geom_point(aes(colour = significant), size = 2.2, shape = 21, fill = "white", stroke = 1.2) +
    scale_colour_manual(values = c("TRUE" = "#d73027", "FALSE" = "#2166ac"), guide = "none") +
    geom_text(aes(label = point_label), vjust = -0.9, size = 2.5, colour = "#333333") +
    facet_wrap(~ label, scales = "free_y", ncol = 2) +
    labs(
      title    = paste0("Sheet: ", toupper(sh), " — ",
                        if (plot_type == "predicted") "Predicted Means" else "Regression Estimates",
                        " by Time Period"),
      subtitle = subtitle_txt,
      x        = "Quarter",
      y        = y_label,
      caption  = if (plot_type == "predicted") {
        "95% CIs via delta method. SEs clustered on Region_code."
      } else {
        "95% CIs. SEs clustered on Region_code. Red points = p < 0.05 vs reference period."
      }
    ) +
    theme_minimal(base_size = 11) +
    theme(
      axis.text.x      = element_text(angle = 45, hjust = 1, size = 7),
      strip.text       = element_text(face = "bold", size = 9),
      panel.grid.minor = element_blank()
    )
  
}), sheets_in_results)


for (sh in sheets_in_results) print(plots[[sh]])

message("Analysis complete.")
