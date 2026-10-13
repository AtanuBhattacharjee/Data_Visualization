# =============================================================================
#  Displaying and Visualising Data
#  Interactive classroom demonstration (MPH - Fundamentals of Statistics)
#  Frequency distributions (Session 2) and presentation and visualisation:
#  building a plot in layers, comparing groups, scatter plots and overplotting,
#  honest axes, colour accessibility, small multiples, tables versus figures,
#  the figure checklist and captions. Every display shows its ggplot2 code.
#  "Build your slides" takes the user's own data, recommends the displays that
#  suit each variable's measurement scale, and exports the chosen figures as a
#  PowerPoint or PDF presentation.
#  Community Hypertension Control Programme, n = 40
#
#  Run locally:   shiny::runApp()      (from the folder containing app.R)
#  Packages:      shiny, bslib, ggplot2, readxl, writexl, officer
#  The class dataset is built in. It can be downloaded as Excel (or CSV) from the
#  sidebar, edited or kept as is, and uploaded back (.xlsx, .xls or .csv); every
#  tab then runs on the uploaded data.
# =============================================================================

library(shiny)
library(bslib)
library(ggplot2)
library(readxl)
library(writexl)

# -----------------------------------------------------------------------------
# 1. Class dataset (built in, so the app runs with no external files)
# -----------------------------------------------------------------------------
hyp_builtin <- read.csv(text = trimws('
ID,Sex,Arm,Education,Smoker,Age,BMI,SBP_Base,SBP_6mo,Adherence,Visits,BP_Controlled,Responded
1,F,Usual care,No formal,No,54,26.4,159,160,3,2,No,No
2,M,Usual care,Primary,Yes,61,25.5,150,150,1,3,No,No
3,F,Usual care,Primary,No,47,22.4,148,139,4,1,Yes,No
4,F,Usual care,Secondary,No,58,27.3,145,146,2,3,No,No
5,M,Usual care,No formal,Yes,66,34.6,174,179,2,1,No,No
6,F,Usual care,Secondary,No,52,23.0,141,147,3,1,No,No
7,M,Usual care,Tertiary,Yes,44,24.8,140,139,4,7,Yes,No
8,F,Usual care,Primary,No,63,30.1,165,160,1,2,No,No
9,M,Usual care,Secondary,Yes,49,24.8,152,168,1,2,No,No
10,F,Usual care,No formal,No,57,29.1,149,156,1,2,No,No
11,F,Usual care,Primary,No,71,29.8,167,163,3,12,No,No
12,M,Usual care,Secondary,Yes,45,21.3,141,150,4,1,No,No
13,F,Usual care,Tertiary,No,50,24.1,147,159,3,3,No,No
14,M,Usual care,Primary,No,59,30.4,151,151,4,3,No,No
15,F,Usual care,No formal,No,68,30.5,174,171,4,4,No,No
16,M,Usual care,Secondary,Yes,42,29.6,148,142,2,4,No,No
17,F,Usual care,Primary,No,55,29.8,163,154,3,4,No,No
18,F,Usual care,Tertiary,No,48,29.2,145,147,1,5,No,No
19,M,Usual care,Secondary,Yes,64,25.2,153,138,3,3,Yes,Yes
20,F,Usual care,Primary,No,53,27.3,171,169,3,2,No,No
21,M,Intervention,No formal,Yes,60,27.0,160,146,5,9,No,Yes
22,F,Intervention,Primary,No,46,30.2,157,134,5,5,Yes,Yes
23,F,Intervention,Secondary,No,56,24.0,153,155,4,6,No,No
24,M,Intervention,Primary,Yes,51,28.0,152,141,5,9,No,Yes
25,F,Intervention,Tertiary,No,43,26.6,140,131,4,4,Yes,No
26,F,Intervention,No formal,No,82,27.2,159,148,4,15,No,Yes
27,M,Intervention,Secondary,Yes,62,26.2,149,131,5,5,Yes,Yes
28,F,Intervention,Primary,No,49,25.7,160,145,5,2,No,Yes
29,M,Intervention,Secondary,No,57,26.3,149,125,5,2,Yes,Yes
30,F,Intervention,Primary,No,44,24.3,157,143,5,5,No,Yes
31,M,Intervention,No formal,Yes,65,22.4,165,151,5,7,No,Yes
32,F,Intervention,Tertiary,No,41,27.6,140,131,4,8,Yes,No
33,F,Intervention,Primary,No,53,31.2,163,152,5,4,No,Yes
34,M,Intervention,Secondary,Yes,59,27.4,152,138,5,4,Yes,Yes
35,F,Intervention,Primary,No,47,21.1,147,127,4,5,Yes,Yes
36,M,Intervention,No formal,Yes,69,26.5,161,144,3,5,No,Yes
37,F,Intervention,Secondary,No,50,26.7,155,142,3,4,No,Yes
38,M,Intervention,Tertiary,No,45,26.7,150,144,5,7,No,No
39,F,Intervention,Primary,No,58,18.5,154,132,4,6,Yes,Yes
40,M,Intervention,Secondary,Yes,54,28.6,162,153,4,6,No,No
'), stringsAsFactors = FALSE, strip.white = TRUE)

# Tidy any loaded table: data frame, syntactic names, no empty rows/columns,
# whole-number columns stored as integers, character values trimmed
clean_df <- function(df) {
  df <- as.data.frame(df, stringsAsFactors = FALSE)
  names(df) <- make.names(trimws(names(df)), unique = TRUE)
  for (v in names(df)) {
    x <- df[[v]]
    if (is.character(x)) { x <- trimws(x); x[x == ""] <- NA }
    if (inherits(x, c("POSIXct", "Date"))) x <- as.character(x)
    if (is.numeric(x) && all(is.na(x) | abs(x - round(x)) < 1e-9)) x <- as.integer(round(x))
    df[[v]] <- x
  }
  keep_c <- vapply(df, function(x) any(!is.na(x)), logical(1))
  df <- df[, keep_c, drop = FALSE]
  df[rowSums(!is.na(df)) > 0, , drop = FALSE]
}

read_any <- function(path, name) {
  ext <- tolower(tools::file_ext(name))
  if (ext %in% c("xlsx", "xls")) {
    sh <- readxl::excel_sheets(path)
    df <- readxl::read_excel(path, sheet = if ("Data" %in% sh) "Data" else sh[1])
  } else {
    df <- read.csv(path, stringsAsFactors = FALSE, strip.white = TRUE)
  }
  clean_df(df)
}

hyp_builtin <- clean_df(hyp_builtin)

# -----------------------------------------------------------------------------
# 2. Variable metadata
# -----------------------------------------------------------------------------
pal <- c("#2A9D8F", "#E76F51", "#4C8BF5", "#E9A93A", "#9B72CF",
         "#43AA8B", "#F28C4E", "#E0609A")
col_mean   <- "#E76F51"
col_median <- "#4C8BF5"
col_fence  <- "#9B72CF"
ramp_cols  <- function(n) grDevices::colorRampPalette(
  c("#2A9D8F", "#43AA8B", "#4C8BF5", "#9B72CF", "#E0609A", "#E76F51", "#E9A93A"))(max(1, n))
quarter_cols <- c("Lowest quarter" = "#4C8BF5", "Second quarter" = "#2A9D8F",
                  "Third quarter" = "#E9A93A", "Top quarter" = "#E76F51", "Flagged" = "#E0609A")

share_bar <- function(p, col) sprintf("<span class='share'><span style='width:%.1f%%;background:%s'></span></span>", p, col)

tile <- function(val, lab, k) div(class = paste0("tile tile-", k), div(class = "tile-val", val), div(class = "tile-lab", lab))

hdr <- function(txt, k) card_header(txt, class = paste0("hdr hdr-", k))

known_scales <- c(ID = "identifier", Sex = "nominal", Arm = "nominal",
                  Education = "ordinal", Smoker = "nominal", Age = "ratio",
                  BMI = "ratio", SBP_Base = "ratio", SBP_6mo = "ratio",
                  Adherence = "ordinal", Visits = "ratio",
                  BP_Controlled = "nominal", Responded = "nominal",
                  AgeGroup = "ordinal")

known_levels <- list(
  Sex = c("F", "M"),
  Arm = c("Usual care", "Intervention"),
  Education = c("No formal", "Primary", "Secondary", "Tertiary"),
  Smoker = c("No", "Yes"),
  Adherence = as.character(1:5),
  BP_Controlled = c("No", "Yes"),
  Responded = c("No", "Yes"),
  AgeGroup = c("<50", "50\u201359", "60+")
)

level_labels <- list(
  Adherence = c("1" = "1 Never", "2" = "2 Rarely", "3" = "3 Sometimes",
                "4" = "4 Usually", "5" = "5 Always"),
  Sex = c(F = "Female", M = "Male")
)

var_labels <- c(ID = "Patient ID", Sex = "Sex", Arm = "Treatment arm",
                Education = "Education", Smoker = "Current smoker",
                Age = "Age (years)", BMI = "BMI (kg/m\u00b2)",
                SBP_Base = "Baseline systolic BP (mmHg)",
                SBP_6mo = "Systolic BP at 6 months (mmHg)",
                Adherence = "Medication adherence (1\u20135)",
                Visits = "Clinic visits",
                BP_Controlled = "BP controlled at 6 months",
                Responded = "Responded",
                AgeGroup = "Age group (years)")

vlab <- function(v) if (v %in% names(var_labels)) unname(var_labels[v]) else v

scale_of <- function(df, v) {
  if (v %in% names(known_scales)) return(unname(known_scales[v]))
  if (tolower(v) %in% c("id", "patient_id", "subject_id", "pid")) return("identifier")
  if (is.numeric(df[[v]])) "ratio" else "nominal"
}

vars_of <- function(df, types) {
  v <- names(df)
  v[vapply(v, function(z) scale_of(df, z) %in% types, logical(1))]
}

lab_choices <- function(v) setNames(v, vapply(v, vlab, character(1)))

pick <- function(ch, pref) if (length(ch) == 0) character(0) else if (pref %in% ch) pref else ch[1]

get_levels <- function(df, v) {
  x <- df[[v]]
  obs <- unique(as.character(x[!is.na(x)]))
  kl <- known_levels[[v]]
  if (!is.null(kl) && all(obs %in% kl)) return(kl)
  if (is.numeric(x)) return(as.character(sort(unique(x[!is.na(x)]))))
  sort(obs)
}

as_cat <- function(df, v) {
  lv <- get_levels(df, v)
  f <- factor(as.character(df[[v]]), levels = lv)
  lab <- level_labels[[v]]
  if (!is.null(lab) && all(lv %in% names(lab))) levels(f) <- unname(lab[lv])
  f
}

# -----------------------------------------------------------------------------
# 3. Statistical helpers
# -----------------------------------------------------------------------------
dec_places <- function(x) {
  x <- x[!is.na(x)]
  for (d in 0:4) if (all(abs(round(x, d) - x) < 1e-8)) return(d)
  4
}

fmt <- function(v, d) formatC(v, format = "f", digits = d)
fmtn <- function(v) vapply(v, function(z) if (is.na(z)) "\u2014" else format(round(z, 2), nsmall = 0), character(1))

# Adjusted Fisher-Pearson skewness (G1), as reported on the slides
skew_G1 <- function(x) {
  x <- x[!is.na(x)]; n <- length(x)
  if (n < 3 || sd(x) == 0) return(NA_real_)
  m <- mean(x)
  g1 <- mean((x - m)^3) / mean((x - m)^2)^1.5
  g1 * sqrt(n * (n - 1)) / (n - 2)
}

shape_verdict <- function(s) {
  if (is.na(s)) return("\u2014")
  if (abs(s) < 0.5) "Approximately symmetric"
  else if (s > 0) "Positively (right) skewed"
  else "Negatively (left) skewed"
}

# Sturges' rule
sturges <- function(x) {
  x <- x[!is.na(x)]; n <- length(x)
  k <- 1 + 3.322 * log10(n)
  rng <- max(x) - min(x)
  list(n = n, k = k, rng = rng, w_raw = rng / k, min = min(x), max = max(x))
}

nice_width <- function(w_raw, u) {
  if (!is.finite(w_raw) || w_raw <= 0) return(u)
  p <- 10^floor(log10(w_raw))
  cand <- c(1, 2, 5, 10) * p
  max(cand[which.min(abs(cand - w_raw))], u)
}

default_w <- function(x) {
  x <- x[!is.na(x)]
  u <- 10^-dec_places(x)
  nice_width(sturges(x)$w_raw, u)
}

default_start <- function(x, w) floor(min(x, na.rm = TRUE) / w + 1e-9) * w

# Grouped frequency distribution with stated limits, boundaries and midpoints
grouped <- function(x, w, start = NULL) {
  x <- x[!is.na(x)]
  d <- dec_places(x); u <- 10^-d
  w <- max(w, u)
  if (is.null(start) || is.na(start)) start <- default_start(x, w)
  K <- floor((max(x) - start) / w + 1e-9) + 1
  K <- max(1, min(K, 80))
  lower <- start + (0:(K - 1)) * w
  upper <- lower + w - u
  brks  <- start - u / 2 + (0:K) * w
  idx <- findInterval(x, brks)
  inside <- idx >= 1 & idx <= K
  f <- tabulate(idx[inside], nbins = K)
  list(lower = lower, upper = upper, lb = brks[-(K + 1)], ub = brks[-1],
       mid = (lower + upper) / 2, f = f, cf = cumsum(f), n = length(x),
       K = K, d = d, u = u, w = w, start = start, brks = brks,
       uncovered = sum(!inside), min = min(x), max = max(x))
}

grouped_df <- function(g) {
  d <- g$d
  out <- data.frame(
    "Class (stated limits)" = paste0(fmt(g$lower, d), " \u2013 ", fmt(g$upper, d)),
    "Class boundaries" = paste0(fmt(g$lb, d + 1), " \u2013 ", fmt(g$ub, d + 1)),
    "Midpoint" = fmt(g$mid, d + 1),
    "f" = as.character(g$f),
    "rf (%)" = fmt(100 * g$f / g$n, 1),
    "cf" = as.character(g$cf),
    "c%" = fmt(100 * g$cf / g$n, 1),
    "Share" = share_bar(100 * g$f / g$n, ramp_cols(g$K)),
    check.names = FALSE)
  rbind(out, data.frame("Class (stated limits)" = "Total", "Class boundaries" = "",
                        "Midpoint" = "", "f" = as.character(sum(g$f)),
                        "rf (%)" = fmt(100 * sum(g$f) / g$n, 1), "cf" = "\u2014",
                        "c%" = "\u2014", "Share" = "", check.names = FALSE))
}

tally_html <- function(k) {
  if (k == 0) return("")
  five <- k %/% 5; rest <- k %% 5
  paste0(paste(rep("<span class='tally5'>||||</span>", five), collapse = " "),
         if (five > 0 && rest > 0) " " else "", strrep("|", rest))
}

integer_breaks <- function(l) { b <- pretty(l); b[b == round(b)] }

theme_class <- function(fs) {
  theme_minimal(base_size = fs) +
    theme(panel.grid.minor = element_blank(),
          panel.grid.major.x = element_blank(),
          plot.subtitle = element_text(colour = "#55616B"),
          axis.title = element_text(colour = "#33414D"),
          plot.title = element_text(face = "bold", size = rel(1), colour = "#1E6F65"),
          panel.grid.major.y = element_line(colour = "#E6ECF2"),
          panel.background = element_rect(fill = "#FBFCFF", colour = NA),
          legend.position = "top",
          plot.background = element_rect(fill = "white", colour = NA))
}

hist_plot <- function(g, xlab, fs, fill = NULL) {
  d <- data.frame(xmin = g$lb, xmax = g$ub, f = g$f, cls = factor(seq_len(g$K)))
  fills <- if (is.null(fill)) ramp_cols(g$K) else rep(fill, g$K)
  p <- ggplot(d) +
    geom_rect(aes(xmin = xmin, xmax = xmax, ymin = 0, ymax = f, fill = cls),
              colour = "white", linewidth = 0.7) +
    scale_fill_manual(values = fills, guide = "none") +
    scale_y_continuous(expand = expansion(mult = c(0, 0.08)), breaks = integer_breaks) +
    labs(x = xlab, y = "Frequency (f)") +
    theme_class(fs)
  if (g$K <= 16) {
    p <- p + scale_x_continuous(breaks = g$brks, labels = fmt(g$brks, g$d + 1))
    if (g$K > 7) p <- p + theme(axis.text.x = element_text(angle = 40, hjust = 1))
  }
  p
}

cat_counts <- function(f) {
  f <- f[!is.na(f)]
  tb <- table(f)
  data.frame(level = factor(names(tb), levels = names(tb)),
             n = as.integer(tb), pct = 100 * as.integer(tb) / sum(tb))
}

legit_displays <- function(scale) {
  switch(scale,
         nominal = "Frequency table without cumulative columns; bar chart (any order); pie for 2\u20133 categories",
         ordinal = "Frequency table with cumulative columns; bar chart kept in natural order",
         ratio = , interval = "Grouped frequency table; histogram; stem-and-leaf; polygon; ogive; boxplot",
         identifier = "None \u2014 a label, not a measurement")
}

dict_df <- function(df) {
  data.frame(Variable = names(df),
             Description = vapply(names(df), vlab, ""),
             Scale = vapply(names(df), function(v) scale_of(df, v), ""),
             "Distinct values" = vapply(df, function(x) length(unique(x[!is.na(x)])), integer(1)),
             "Legitimate summaries and displays" = vapply(names(df), function(v) legit_displays(scale_of(df, v)), ""),
             check.names = FALSE, row.names = NULL)
}

note <- function(...) conditionalPanel("input.show_notes", div(class = "teach-note", ...))

# -----------------------------------------------------------------------------
# 3b. Plot tools, code panels, colour-vision simulation (visualisation tabs)
# -----------------------------------------------------------------------------
# A plot with a PNG download link and, optionally, a fold-out panel showing the
# ggplot2 code that reproduces it on the current settings
plotx <- function(id, height = "400px", code = TRUE) {
  tagList(
    plotOutput(id, height = height),
    div(class = "plot-tools",
        downloadLink(paste0(id, "_png"), HTML("&#8681;&nbsp;Download PNG"), class = "tool-link"),
        if (code) tags$details(class = "code-box",
                               tags$summary("Show the R code"),
                               verbatimTextOutput(paste0(id, "_code"))))
  )
}

code_head <- paste0("library(ggplot2)\n",
                    "hyp <- readxl::read_excel(\"hypertension-dataset.xlsx\", sheet = \"Data\")\n")

# A line declaring a categorical variable's level order, when it has one
lev_code <- function(df, v) {
  lv <- get_levels(df, v)
  if (is.numeric(df[[v]]))
    return(sprintf("hyp$%s <- factor(hyp$%s, levels = c(%s))\n", v, v, paste(lv, collapse = ", ")))
  sprintf("hyp$%s <- factor(hyp$%s, levels = c(%s))\n", v, v, paste0("\"", lv, "\"", collapse = ", "))
}
qq <- function(s) paste0("\"", s, "\"")

# Colour-vision deficiency simulation (Machado, Oliveira & Fernandes 2009, severity 1),
# applied in linear RGB; greyscale uses relative luminance
cvd_mats <- list(
  Protanopia   = matrix(c(0.152286, 1.052583, -0.204868, 0.114503, 0.786281, 0.099216,
                          -0.003882, -0.048116, 1.051998), 3, byrow = TRUE),
  Deuteranopia = matrix(c(0.367322, 0.860646, -0.227968, 0.280085, 0.672501, 0.047413,
                          -0.011820, 0.042940, 0.968881), 3, byrow = TRUE),
  Tritanopia   = matrix(c(1.255528, -0.076749, -0.178779, -0.078411, 0.930809, 0.147602,
                          0.004733, 0.691367, 0.303900), 3, byrow = TRUE))

to_lin <- function(hex) {
  rgb <- t(grDevices::col2rgb(hex)) / 255
  ifelse(rgb <= 0.04045, rgb / 12.92, ((rgb + 0.055) / 1.055)^2.4)
}
simulate_cvd <- function(hex, type) {
  if (type == "Normal vision") return(hex)
  lin <- to_lin(hex)
  out <- if (type == "Greyscale") {
    y <- lin %*% c(0.2126, 0.7152, 0.0722); cbind(y, y, y)
  } else lin %*% t(cvd_mats[[type]])
  out <- pmin(pmax(out, 0), 1)
  s <- ifelse(out <= 0.0031308, 12.92 * out, 1.055 * out^(1 / 2.4) - 0.055)
  grDevices::rgb(s[, 1], s[, 2], s[, 3])
}
# CIE lightness L* (0 = black, 100 = white)
lightness <- function(hex) {
  y <- as.vector(to_lin(hex) %*% c(0.2126, 0.7152, 0.0722))
  ifelse(y > 216 / 24389, 116 * y^(1 / 3) - 16, y * 24389 / 27)
}

acc_pal <- function(name, n) switch(name,
  redgreen  = rep_len(c("#D7301F", "#1A9850", "#F46D43", "#66BD63", "#A50026", "#006837"), n),
  lightness = grDevices::colorRampPalette(c("#C6DBEF", "#2171B5", "#08306B"))(n),
  okabe     = rep_len(c("#E69F00", "#0072B2", "#009E73", "#CC79A7", "#56B4E9", "#D55E00"), n))
acc_pal_code <- c(redgreen = "c(\"#D7301F\", \"#1A9850\")  # red and green of similar lightness",
                  lightness = "colorRampPalette(c(\"#C6DBEF\", \"#2171B5\", \"#08306B\"))(n)  # light to dark",
                  okabe = "c(\"#E69F00\", \"#0072B2\", \"#009E73\", \"#CC79A7\")  # Okabe-Ito")

# Age group derived for faceting (only when Age is present and numeric)
add_derived <- function(df) {
  if ("Age" %in% names(df) && is.numeric(df$Age) && !"AgeGroup" %in% names(df))
    df$AgeGroup <- as.character(cut(df$Age, c(-Inf, 49, 59, Inf), labels = c("<50", "50\u201359", "60+")))
  df
}
facet_choices <- function(df) {
  d2 <- add_derived(df)
  lab_choices(unique(c(if ("AgeGroup" %in% names(d2)) "AgeGroup", vars_of(d2, c("nominal", "ordinal")))))
}

# Narrow axis labels: break the first line of a label at its spaces
wrap_first <- function(z) vapply(strsplit(z, "\n", fixed = TRUE),
                                 function(p) paste(c(gsub(" ", "\n", p[1]), p[-1]), collapse = "\n"), "")

# Group labels carrying their n, e.g. "Usual care (n = 20)"
with_n <- function(f, sep = " ") {
  tb <- table(f)
  factor(as.character(f), levels = names(tb), labels = sprintf("%s%s(n = %d)", names(tb), sep, as.integer(tb)))
}

# Baseline characteristics in the class dataset (the rest are follow-up measures or outcomes)
followup_vars <- c("SBP_6mo", "BP_Controlled", "Responded", "Adherence", "Visits")

# The eight checks before a figure leaves your desk
ck_items <- c(
  scale   = "The display matches the measurement scale (an ordinal variable keeps its order)",
  labels  = "Every axis is labelled, with units",
  zero    = "A bar\u2019s axis starts at zero, or any truncation is visible",
  n       = "The n is given for every group or panel",
  data    = "The data are shown, not only a mean and an error bar",
  grey    = "The message survives in greyscale: colour is not the only cue",
  caption = "The caption stands alone: what is plotted, what each mark means, the group sizes",
  one     = "One vehicle per result: the figure does not duplicate a table")

# -----------------------------------------------------------------------------
# 4. Checkpoint questions (from the session slides)
# -----------------------------------------------------------------------------
cps <- list(
  list(q = "A cumulative frequency column would be meaningless for which of these variables?",
       opts = c(A = "Adherence score (1\u20135)", B = "Baseline systolic blood pressure",
                C = "Current smoker (Yes / No)", D = "Number of clinic visits"),
       ans = "C",
       why = "Cumulating means counting everything \u201cat or below\u201d a value, which requires ordered categories. Smoker is nominal \u2014 Yes is not above or below No. The other three are ordinal or ratio and cumulate legitimately."),
  list(q = "A class of \u201c150 \u2013 159 mmHg\u201d is recorded to the nearest whole mmHg. What is its class width?",
       opts = c(A = "9", B = "9.5", C = "10", D = "It cannot be determined"),
       ans = "C",
       why = "Width is measured between boundaries: 159.5 \u2212 149.5 = 10, or equivalently 160 \u2212 150. Answer A subtracts the stated limits and forgets that both 150 and 159 are inside the class."),
  list(q = "You want to compare the education profile of 20 intervention patients against 20 usual-care patients. Which display is best?",
       opts = c(A = "Two separate pie charts, one per arm", B = "A clustered bar chart of counts",
                C = "A clustered bar chart of percentages within each arm", D = "A single histogram of education"),
       ans = "C",
       why = "Percentages within each arm keep the comparison valid whatever the group sizes, and clustering puts the two profiles side by side. B works only because both arms happen to have n = 20. D is wrong outright \u2014 education is ordinal."),
  list(q = "For a boxplot of clinic visits, Q1 = 2 and Q3 = 6. Which value is flagged as an outlier?",
       opts = c(A = "9", B = "11", C = "12", D = "15"),
       ans = "D",
       why = "IQR = 6 \u2212 2 = 4, so the upper fence is 6 + 1.5 \u00d7 4 = 12. Only values beyond the fence are flagged: 15 qualifies, 12 sits exactly on the fence and does not."),
  list(q = "A histogram of clinic visits has a long tail stretching to the right. How is this distribution described, and what follows?",
       opts = c(A = "Negatively skewed; report the mean", B = "Positively skewed; the mean will exceed the median",
                C = "Positively skewed; the mean will fall below the median", D = "Bimodal; report both modes"),
       ans = "B",
       why = "Skew is named for the direction of the tail. Extreme high values drag the mean upward and leave the median almost untouched \u2014 mean 4.58 against a median of 4.0 in our data.")
)


# Visualisation checkpoints (6 to 10), on the class dataset
cps <- c(cps, list(
  list(q = "Medication adherence runs 1 Never, 2 Rarely, 3 Sometimes, 4 Usually, 5 Always. How should the bars of a bar chart be ordered?",
       opts = c(A = "By frequency, tallest first", B = "In the order of the scale",
                C = "Alphabetically", D = "It makes no difference"),
       ans = "B",
       why = "An ordinal scale carries one piece of information beyond the category names: their order. Sorting by frequency throws it away, so the reader can no longer see whether patients sit at the adherent end or the non-adherent end. Frequency order is right for a nominal variable, where there is no order to lose."),
  list(q = "A figure shows two bars with standard-error whiskers comparing mean systolic BP at 6 months between the arms. What is the main objection?",
       opts = c(A = "Bar charts are always wrong",
                B = "It shows two summary numbers and hides the 40 patients",
                C = "The standard error is miscalculated", D = "The bars should be horizontal"),
       ans = "B",
       why = "With 20 patients per arm there is no reason to show only two numbers. The bars hide the range (125 to 155 mmHg against 138 to 179 mmHg), the overlap and any outlying patient, and standard-error whiskers make the arms look more separated than the patients are. Bars are right for counts of categories; a mean is not a count."),
  list(q = "A bar chart of BP controlled at 6 months (15% usual care, 40% intervention) uses a y-axis from 12% to 45%. What is the effect?",
       opts = c(A = "It clarifies a small difference", B = "It exaggerates the difference about 3.5-fold",
                C = "It has no effect on interpretation", D = "It is standard practice"),
       ans = "B",
       why = "The bar encodes its value by length, so the reader compares 40 with 15, a ratio of 2.7. From a base of 12 the visible heights are 28 and 3, an apparent ratio of 9.3: about 3.5 times the real one, for a difference that was plain on a full axis."),
  list(q = "A facetted figure of SBP at 6 months by arm has one panel per education level. The Tertiary panel holds 6 patients, 3 per arm. What should you conclude from that panel?",
       opts = c(A = "The effect reverses in tertiary-educated patients",
                B = "Very little: six patients cannot support a comparison",
                C = "It confirms the overall finding", D = "The panel should be deleted"),
       ans = "B",
       why = "Three patients per arm cannot distinguish anything, which is why the n belongs in every panel: without it, four panels of equal size look like four equal pieces of evidence. Deleting the panel would be selective reporting. Show it, label its n, and describe the trend across panels rather than any one panel."),
  list(q = "You need to report the mean and standard deviation of six baseline characteristics by arm. Table or figure?",
       opts = c(A = "A figure: it is more engaging", B = "A table: exact values matter and readers may re-use them",
                C = "Both", D = "Neither; describe it in text"),
       ans = "B",
       why = "This is Table 1. The exact values matter and other researchers may extract them, for example for a meta-analysis. A figure would trade the numbers for the pattern, which is the wrong trade here; twelve numbers in prose are unreadable, and printing both wastes space on a duplicate.")
))
cp_sets <- list("Frequency distributions" = 1:5, "Visualisation" = 6:10)


# -----------------------------------------------------------------------------
# 3c. Slide builder: recommended displays for the user's own data
# -----------------------------------------------------------------------------
okabe <- c("#E69F00", "#0072B2", "#009E73", "#CC79A7", "#56B4E9", "#D55E00", "#F0E442", "#999999")

# Default scale for the builder: quantitative / ordinal / nominal / identifier
b_default_scale <- function(df, v) {
  sc <- scale_of(df, v)
  if (sc %in% c("ratio", "interval")) {
    x <- df[[v]]; ux <- unique(x[!is.na(x)])
    if (!v %in% names(known_scales) && length(ux) <= 2) return("nominal")
    return("quantitative")
  }
  sc
}
b_label <- function(v) if (v %in% names(var_labels)) unname(var_labels[v]) else gsub("[._]+", " ", v)
# lower-case the first letter only when it starts an ordinary word ("Systolic BP" -> "systolic BP", "BMI" stays)
low1 <- function(s) ifelse(grepl("^[A-Z][a-z]", s), paste0(tolower(substr(s, 1, 1)), substring(s, 2)), s)

# Recommended displays for a combination of measurement scales
b_recommend <- function(kind, info) {
  r <- function(type, label, verdict, why) data.frame(type = type, label = label, verdict = verdict, why = why)
  switch(kind,
    nominal1 = rbind(
      r("bar_freq", "Bar chart, sorted by frequency", "best", "Nominal categories have no order, so sorting by frequency makes the ranking easy to read."),
      r("dot_pct", "Lollipop chart of percentages", "good", "Lighter than bars when there are many categories."),
      if (info$nlev <= 3) r("pie", "Pie chart", "good", sprintf("Acceptable with %d categories that make up a whole.", info$nlev))
      else r("pie", "Pie chart", "avoid", sprintf("%d slices: people judge angles badly, so the slices cannot be ranked.", info$nlev))),
    ordinal1 = rbind(
      r("bar_order", "Bar chart, in the order of the scale", "best", "Keeps the one extra piece of information an ordinal scale carries: its order."),
      r("dot_pct", "Lollipop chart of percentages, in scale order", "good", "Same information with less ink."),
      r("bar_freq", "Bar chart sorted by frequency", "avoid", "Sorting by frequency destroys the order of the scale."),
      r("pie", "Pie chart", "avoid", "A pie has no order, so the scale is lost.")),
    quant1 = rbind(
      if (info$n < 30) r("strip", "Every value, with the median", "best", sprintf("With n = %d, every observation can be shown.", info$n))
      else r("hist", "Histogram", "best", "Shows the shape, the centre, the spread and any tail."),
      if (info$n < 30) r("hist", "Histogram", "good", "Works, but with few values the shape depends heavily on the bin width.")
      else r("strip", "Every value, with the median", "good", "Shows every observation; best when n is modest."),
      r("box", "Boxplot with the points", "good", "Five-number summary plus flagged values; compact for reports."),
      r("ecdf", "Cumulative distribution (ECDF)", "good", "Reads off any percentile directly; no bin width to choose."),
      r("bar_mean", "A single bar for the mean", "avoid", "A mean is not an amount, and the bar hides the whole distribution.")),
    qc = rbind(
      r(if (info$maxg <= 150) "pts_med" else "box_pts",
        if (info$maxg <= 150) "Every value by group, with the median" else "Boxplot with the points", "best",
        if (info$maxg <= 150) sprintf("At most %d per group: show everyone and let the reader see the overlap.", info$maxg)
        else "Large groups: a box summarises, the points keep the outliers honest."),
      r(if (info$maxg <= 150) "box_pts" else "pts_med",
        if (info$maxg <= 150) "Boxplot with the points" else "Every value by group, with the median", "good",
        "Summary and data together."),
      r("box", "Boxplot", "good", "Compact when there are many groups."),
      r("hist_facet", "Histogram for each group", "good", "Compares shapes; keep the axes shared."),
      r("dynamite", "Bar of the mean with an error bar", "avoid", "Two numbers per group stand in for every observation, and SE whiskers exaggerate separation.")),
    cc = rbind(
      r("cluster_pct", "Clustered bars of % within each group", "best", "Percentages within each group stay comparable whatever the group sizes."),
      r("stack_pct", "100% stacked bars", "good", "Shows each group's whole profile; middle segments are harder to compare."),
      if (info$equal) r("cluster_n", "Clustered bars of counts", "good", "Acceptable here because the groups are the same size.")
      else r("cluster_n", "Clustered bars of counts", "avoid", "The groups differ in size, so counts mislead; use percentages."),
      r("pies", "One pie per group", "avoid", "Angles across separate pies cannot be compared.")),
    qq = rbind(
      if (info$ties > 0.2) r("jitter", "Scatter plot with jitter and transparency", "best",
                             sprintf("%.0f%% of points sit on top of another; jitter reveals them.", 100 * info$ties))
      else r("scatter", "Scatter plot", "best", "Each point is one observation; the relationship is visible at once."),
      if (info$ties > 0.2) r("scatter", "Plain scatter plot", "good", "Hides the overlapping points.")
      else r("jitter", "Scatter plot with jitter", "good", "Only needed when values coincide."),
      r("count", "Scatter plot sized by count", "good", "Makes stacked points visible without moving them."),
      r("dual", "Two lines on two y-axes", "avoid", "Two axes let the scales be tuned to show any relationship.")))
}

b_caption_default <- function(sp) {
  u <- sp$unit; us <- paste0(u, "s")
  nn <- function(f) paste(sprintf("%s n = %d", names(table(f)), as.integer(table(f))), collapse = "; ")
  what <- switch(sp$type,
    bar_freq = , bar_order = , pie = sprintf("Number and percentage of %s in each category of %s.", us, low1(sp$lab1)),
    dot_pct = sprintf("Percentage of %s in each category of %s.", us, low1(sp$lab1)),
    hist = sprintf("Distribution of %s; bars span the class boundaries (width %s).", low1(sp$lab1), format(sp$bw)),
    strip = sprintf("%s. Each point is one %s; the horizontal line is the median.", sp$lab1, u),
    box = if (!is.null(sp$d$g)) sprintf("%s by %s. Box = quartiles, line = median, whiskers to the most extreme values within 1.5 \u00d7 IQR; dots are flagged values.", sp$laby, low1(sp$labg))
          else sprintf("%s. Box = quartiles, line = median, whiskers to the most extreme values within 1.5 \u00d7 IQR; each point is one %s.", sp$lab1, u),
    ecdf = sprintf("Cumulative distribution of %s: the proportion of %s at or below each value.", low1(sp$lab1), us),
    pts_med = sprintf("%s by %s. Each point is one %s; the horizontal line is the group median.", sp$laby, low1(sp$labg), u),
    box_pts = sprintf("%s by %s. Box = quartiles, line = median; each point is one %s.", sp$laby, low1(sp$labg), u),
    hist_facet = sprintf("Distribution of %s by %s, on shared axes.", low1(sp$laby), low1(sp$labg)),
    cluster_pct = sprintf("%s by %s, as the percentage of %s within each group.", sp$lab1, low1(sp$labg), us),
    stack_pct = sprintf("%s by %s: each bar is one group and sums to 100%%.", sp$lab1, low1(sp$labg)),
    cluster_n = sprintf("%s by %s, number of %s.", sp$lab1, low1(sp$labg), us),
    scatter = , jitter = , count = sprintf("%s against %s. Each point is one %s%s%s.", sp$laby, low1(sp$labx), u,
                                           if (sp$type == "jitter") ", displaced slightly to reveal overlapping values" else "",
                                           if (isTRUE(sp$lm)) "; the line is the least-squares fit with its 95% confidence band" else ""),
    "")
  ns <- if (!is.null(sp$d$g)) nn(sp$d$g) else sprintf("n = %d", nrow(sp$d))
  if (!is.null(sp$d$f)) ns <- paste0(ns, ". Panels: ", nn(sp$d$f))
  paste0(what, " ", ns, ".")
}

# Draw one figure from a specification
b_draw <- function(sp, fs) {
  d <- sp$d; type <- sp$type; th <- theme_class(fs)
  cols <- function(n) rep_len(okabe, n)
  pctlab <- function(n) sprintf("%d (%.0f%%)", n, 100 * n / sum(n))
  if (type %in% c("bar_freq", "bar_order", "dot_pct", "pie")) {
    tb <- as.data.frame(table(level = d$cat)); tb$pct <- 100 * tb$Freq / sum(tb$Freq)
    if (type == "bar_freq" || (type == "dot_pct" && sp$scale1 == "nominal"))
      tb$level <- factor(tb$level, levels = tb$level[order(-tb$Freq)])
    if (type == "pie")
      return(ggplot(tb, aes(x = 1, y = Freq, fill = level)) + geom_col(width = 1, colour = "white") +
               coord_polar(theta = "y", direction = -1) + scale_fill_manual(values = cols(nrow(tb)), name = NULL) +
               geom_text(aes(label = sprintf("%.0f%%", pct)), position = position_stack(vjust = 0.5), size = fs / 3) +
               theme_void(base_size = fs) + theme(legend.position = "right"))
    yv <- if (isTRUE(sp$pct)) "pct" else "Freq"
    if (type == "dot_pct")
      return(ggplot(tb, aes(level, pct)) + geom_segment(aes(xend = level, y = 0, yend = pct), colour = "#9AA7B2", linewidth = 1.2) +
               geom_point(size = 5, colour = okabe[2]) + geom_text(aes(label = sprintf("%.0f%%", pct)), vjust = -1, size = fs / 3.3) +
               scale_y_continuous(limits = c(0, max(tb$pct) * 1.18), expand = expansion(mult = c(0, 0))) +
               labs(x = sp$xlab, y = sp$ylab) + th)
    return(ggplot(tb, aes(level, .data[[yv]])) + geom_col(width = 0.62, fill = okabe[2]) +
             geom_text(aes(label = pctlab(Freq)), vjust = -0.45, size = fs / 3.3) +
             scale_y_continuous(expand = expansion(mult = c(0, 0.14))) + labs(x = sp$xlab, y = sp$ylab) + th)
  }
  if (type %in% c("hist", "strip", "box", "ecdf", "bar_mean")) {
    x <- d$val
    p <- switch(type,
      hist = { g <- grouped(x, sp$bw)
        ggplot(data.frame(xmin = g$lb, xmax = g$ub, f = g$f)) +
          geom_rect(aes(xmin = xmin, xmax = xmax, ymin = 0, ymax = f), fill = okabe[2], colour = "white", linewidth = 0.6) +
          scale_y_continuous(expand = expansion(mult = c(0, 0.08)), breaks = integer_breaks) },
      strip = { set.seed(2)
        ggplot(d, aes(x = "", y = val)) + geom_jitter(width = 0.12, height = 0, size = 3, alpha = 0.75, colour = okabe[2]) +
          stat_summary(fun = median, fun.min = median, fun.max = median, geom = "crossbar", width = 0.4, colour = "#33414D") },
      box = { set.seed(2)
        ggplot(d, aes(x = "", y = val)) + geom_boxplot(width = 0.4, fill = "#CFE3F3", outlier.shape = NA) +
          geom_jitter(width = 0.1, height = 0, alpha = 0.6, size = 2.4, colour = okabe[2]) },
      ecdf = ggplot(d, aes(val)) + stat_ecdf(geom = "step", linewidth = 1.2, colour = okabe[2]) +
        scale_y_continuous(labels = function(z) paste0(100 * z, "%")))
    p <- p + labs(x = sp$xlab, y = sp$ylab) + th
    if (type %in% c("strip", "box")) p <- p + theme(axis.text.x = element_blank(), axis.ticks.x = element_blank())
  } else if (type %in% c("pts_med", "box_pts", "box", "hist_facet")) {
    d$gn <- with_n(d$g, "\n"); set.seed(2)
    p <- switch(type,
      pts_med = ggplot(d, aes(gn, val, colour = gn)) + geom_jitter(width = 0.15, height = 0, size = 2.8, alpha = 0.75) +
        stat_summary(fun = median, fun.min = median, fun.max = median, geom = "crossbar", width = 0.45, colour = "#33414D"),
      box_pts = ggplot(d, aes(gn, val)) + geom_boxplot(aes(fill = gn), width = 0.5, alpha = 0.3, outlier.shape = NA) +
        geom_jitter(aes(colour = gn), width = 0.15, height = 0, size = 2.3, alpha = 0.7),
      box = ggplot(d, aes(gn, val, fill = gn)) + geom_boxplot(width = 0.5, alpha = 0.45, outlier.colour = "#D55E00"),
      hist_facet = { g0 <- grouped(d$val, sp$bw)
        rr <- do.call(rbind, lapply(levels(d$gn), function(l) { xi <- d$val[d$gn == l]; idx <- findInterval(xi, g0$brks)
          data.frame(xmin = g0$lb, xmax = g0$ub, f = tabulate(idx[idx >= 1 & idx <= g0$K], g0$K), gn = l) }))
        rr$gn <- factor(rr$gn, levels = levels(d$gn))
        ggplot(rr) + geom_rect(aes(xmin = xmin, xmax = xmax, ymin = 0, ymax = f, fill = gn), colour = "white") +
          facet_wrap(~gn, ncol = 1) + scale_y_continuous(expand = expansion(mult = c(0, 0.08)), breaks = integer_breaks) })
    p <- p + scale_colour_manual(values = cols(nlevels(d$gn)), guide = "none") +
      scale_fill_manual(values = cols(nlevels(d$gn)), guide = "none") + labs(x = sp$xlab, y = sp$ylab) + th
  } else if (type %in% c("cluster_pct", "stack_pct", "cluster_n")) {
    d$gn <- with_n(d$g)
    if (!is.null(d$f)) {
      tb <- as.data.frame(table(cat = d$cat, gn = d$gn, fn = with_n(d$f)))
      tb$pct <- 100 * tb$Freq / pmax(1, ave(tb$Freq, tb$gn, tb$fn, FUN = sum))
    } else {
      tb <- as.data.frame(table(cat = d$cat, gn = d$gn)); tb$pct <- 100 * tb$Freq / ave(tb$Freq, tb$gn, FUN = sum)
    }
    if (type == "stack_pct") {
      p <- ggplot(tb, aes(gn, pct, fill = cat)) + geom_col(width = 0.6, colour = "white") +
        geom_text(aes(label = ifelse(pct >= 6, sprintf("%.0f%%", pct), "")), position = position_stack(vjust = 0.5), size = fs / 3.5) +
        scale_fill_manual(values = cols(nlevels(tb$cat)), name = sp$lab1) + scale_y_continuous(expand = expansion(mult = c(0, 0.02)))
    } else {
      yv <- if (type == "cluster_pct") "pct" else "Freq"
      p <- ggplot(tb, aes(cat, .data[[yv]], fill = gn)) + geom_col(position = position_dodge(0.8), width = 0.75) +
        geom_text(aes(label = if (type == "cluster_pct") sprintf("%.0f%%", pct) else Freq), position = position_dodge(0.8),
                  vjust = -0.4, size = fs / 3.6) +
        scale_fill_manual(values = cols(nlevels(tb$gn)), name = sp$labg) + scale_y_continuous(expand = expansion(mult = c(0, 0.14)))
    }
    p <- p + labs(x = sp$xlab, y = sp$ylab) + th
    if (!is.null(d$f)) p <- p + facet_wrap(~fn, nrow = 1)
  } else {
    p <- ggplot(d, aes(x, val)) + labs(x = sp$xlab, y = sp$ylab) + th
    p <- p + switch(type,
      scatter = geom_point(size = 3, alpha = 0.8, colour = okabe[2]),
      jitter = { set.seed(2); geom_jitter(width = sp$jw, height = sp$jh, size = 3, alpha = 0.5, colour = okabe[2]) },
      count = geom_count(colour = okabe[2], alpha = 0.8))
    if (type == "count") p <- p + scale_size_area(max_size = 11, name = paste0(toupper(substr(sp$unit, 1, 1)), substring(sp$unit, 2), "s"))
    if (isTRUE(sp$lm)) p <- p + geom_smooth(method = "lm", formula = y ~ x, colour = okabe[6], fill = "#F6DCD3", linewidth = 1.1)
  }
  if (!is.null(d$f) && type != "pie") {
    fn <- with_n(d$f)
    if (type == "hist_facet") p <- p + facet_grid(gn ~ fn)
    else if (!type %in% c("cluster_pct", "stack_pct", "cluster_n")) { p$data$fn <- fn; p <- p + facet_wrap(~fn, nrow = 1) }
  }
  if (!is.null(sp$title_on_plot) && nzchar(sp$title_on_plot)) p <- p + labs(title = sp$title_on_plot)
  p + theme(strip.text = element_text(face = "bold"))
}

# Automatic checks on a specification (the eight checks that can be judged by the app)
b_checks <- function(sp) {
  raw <- function(l) is.null(l) || !nzchar(trimws(l)) || grepl("_", l)
  bars <- sp$type %in% c("bar_freq", "bar_order", "cluster_pct", "cluster_n", "stack_pct", "hist", "hist_facet")
  list(
    list(ok = sp$verdict != "avoid", txt = "The display matches the measurement scales"),
    list(ok = sp$type == "pie" || ((sp$type %in% c("strip", "box") || !raw(sp$xlab)) && !raw(sp$ylab)),
         txt = "The axes carry readable labels (add units in brackets)"),
    list(ok = TRUE, txt = if (bars) "Bars start at zero" else "Positions, not lengths: zero is not required"),
    list(ok = grepl("n = ", sp$caption), txt = "The n is given (in the labels and the caption)"),
    list(ok = !sp$type %in% c("dynamite", "bar_mean"), txt = "The data are shown, not only a summary"),
    list(ok = TRUE, txt = "Colour-blind-safe palette; colour is backed by labels or position"),
    list(ok = nchar(sp$caption) >= 60, txt = "The caption says what is plotted, what each mark means and the group sizes"))
}

# Export: PowerPoint (officer) and PDF (base graphics), both light backgrounds
b_slide_png <- function(p, w = 9, h = 4.7) {
  f <- tempfile(fileext = ".png"); ggsave(f, p, width = w, height = h, dpi = 200, bg = "white"); f
}
b_write_pptx <- function(deck, slides, file) {
  doc <- officer::read_pptx()
  doc <- officer::add_slide(doc, layout = "Title Slide", master = "Office Theme")
  doc <- officer::ph_with(doc, deck$title, location = officer::ph_location_type("ctrTitle"))
  doc <- officer::ph_with(doc, deck$sub, location = officer::ph_location_type("subTitle"))
  for (s in slides) {
    doc <- officer::add_slide(doc, layout = "Title Only", master = "Office Theme")
    doc <- officer::ph_with(doc, officer::fpar(officer::ftext(s$title, officer::fp_text(font.size = 26, bold = TRUE, color = "#1E6F65"))),
                            location = officer::ph_location_type("title"))
    doc <- officer::ph_with(doc, officer::external_img(b_slide_png(s$plot), width = 9, height = 4.7),
                            location = officer::ph_location(left = 0.5, top = 1.45, width = 9, height = 4.7))
    doc <- officer::ph_with(doc, officer::fpar(officer::ftext(s$caption, officer::fp_text(font.size = 12, italic = TRUE, color = "#33414D"))),
                            location = officer::ph_location(left = 0.5, top = 6.25, width = 9, height = 1))
    if (nzchar(s$notes)) doc <- officer::set_notes(doc, s$notes, location = officer::notes_location_type("body"))
  }
  print(doc, target = file)
}
b_write_pdf <- function(deck, slides, file) {
  grDevices::cairo_pdf(file, width = 11, height = 7.5, onefile = TRUE)
  on.exit(grDevices::dev.off())
  txt <- function(label, y, size, col = "#1F2A33", face = "plain", x = 0.06, just = "left")
    grid::grid.text(label, x = x, y = y, just = just, gp = grid::gpar(fontsize = size, col = col, fontface = face))
  grid::grid.newpage()
  grid::grid.rect(gp = grid::gpar(fill = "#F2FBF8", col = NA))
  txt(deck$title, 0.56, 34, "#1E6F65", "bold", x = 0.5, just = "centre"); txt(deck$sub, 0.44, 18, "#50606E", x = 0.5, just = "centre")
  for (s in slides) {
    grid::grid.newpage()
    txt(s$title, 0.93, 24, "#1E6F65", "bold")
    print(s$plot, vp = grid::viewport(x = 0.5, y = 0.5, width = 0.9, height = 0.72))
    wrapped <- paste(strwrap(s$caption, 150), collapse = "\n")
    grid::grid.text(wrapped, x = 0.06, y = 0.08, just = c("left", "centre"), gp = grid::gpar(fontsize = 11, fontface = "italic", col = "#33414D"))
  }
}

# -----------------------------------------------------------------------------
# 5. Initial choices
# -----------------------------------------------------------------------------
init_all   <- vars_of(hyp_builtin, c("nominal", "ordinal", "ratio", "interval"))
init_cat   <- vars_of(hyp_builtin, c("nominal", "ordinal"))
init_quant <- vars_of(hyp_builtin, c("ratio", "interval"))
init_num   <- init_all[vapply(init_all, function(v) is.numeric(hyp_builtin[[v]]), logical(1))]
grp_choices <- function(catv) c("None" = "", lab_choices(catv))

measure_choices <- function(df) {
  q <- vars_of(df, c("ratio", "interval"))
  cv <- vars_of(df, c("nominal", "ordinal"))
  bin <- cv[vapply(cv, function(v) "Yes" %in% df[[v]], logical(1))]
  c(setNames(paste0("mean:", q), paste("Mean", vapply(q, vlab, ""))),
    setNames(paste0("pct:", bin), paste("% Yes \u2014", vapply(bin, vlab, ""))))
}

# -----------------------------------------------------------------------------
# 6. User interface
# -----------------------------------------------------------------------------
app_css <- "
body { background: linear-gradient(180deg, #F2FBF8 0%, #F8F5FF 45%, #FFF7EE 100%) fixed; }
.navbar { background: linear-gradient(90deg, #DDF4EE 0%, #E6EDFF 35%, #FBE6F0 70%, #FFEFDC 100%) !important;
          border-bottom: 2px solid #CDEBE3; }
.navbar-brand { font-weight: 700; color: #1E6F65 !important; }
.navbar .nav-link { font-weight: 600; border-radius: 999px; padding: .35rem .85rem !important; margin: 0 .1rem; }
.navbar .nav-item:nth-child(7n+1) .nav-link { color: #1E7F73 !important; }
.navbar .nav-item:nth-child(7n+2) .nav-link { color: #C4532F !important; }
.navbar .nav-item:nth-child(7n+3) .nav-link { color: #2F66C8 !important; }
.navbar .nav-item:nth-child(7n+4) .nav-link { color: #A8740F !important; }
.navbar .nav-item:nth-child(7n+5) .nav-link { color: #7447B0 !important; }
.navbar .nav-item:nth-child(7n+6) .nav-link { color: #B53A72 !important; }
.navbar .nav-item:nth-child(7n+7) .nav-link { color: #2B8A5C !important; }
.navbar .nav-link.active { background: rgba(255,255,255,.85); box-shadow: 0 1px 4px rgba(60,80,100,.12); }
.bslib-sidebar-layout > .sidebar { background: #F6F2FF; }
.bslib-page-navbar > .bslib-sidebar-layout > .sidebar { background: linear-gradient(180deg, #F3EEFF 0%, #EAF7F3 100%); }
.card { border-radius: 14px; border: 1px solid #E3E9F0; box-shadow: 0 2px 10px rgba(80,100,130,.07); }
.hdr { font-weight: 700; border-bottom-width: 3px !important; border-bottom-style: solid !important; }
.hdr-1 { background: #E2F5EF; border-bottom-color: #2A9D8F !important; color: #1E6F65; }
.hdr-2 { background: #FDEBE4; border-bottom-color: #E76F51 !important; color: #A9472B; }
.hdr-3 { background: #E7EEFE; border-bottom-color: #4C8BF5 !important; color: #2F5FB8; }
.hdr-4 { background: #FFF3D9; border-bottom-color: #E9A93A !important; color: #8E6210; }
.hdr-5 { background: #F1E9FB; border-bottom-color: #9B72CF !important; color: #6B45A0; }
.hdr-6 { background: #FCE6F0; border-bottom-color: #E0609A !important; color: #A33165; }
.nav-tabs .nav-link { font-weight: 600; color: #50606E; }
.nav-tabs .nav-item:nth-child(6n+1) .nav-link.active { color: #1E7F73; border-top: 3px solid #2A9D8F; background: #F0FAF7; }
.nav-tabs .nav-item:nth-child(6n+2) .nav-link.active { color: #C4532F; border-top: 3px solid #E76F51; background: #FFF4F0; }
.nav-tabs .nav-item:nth-child(6n+3) .nav-link.active { color: #2F66C8; border-top: 3px solid #4C8BF5; background: #F2F6FF; }
.nav-tabs .nav-item:nth-child(6n+4) .nav-link.active { color: #A8740F; border-top: 3px solid #E9A93A; background: #FFF9EC; }
.nav-tabs .nav-item:nth-child(6n+5) .nav-link.active { color: #7447B0; border-top: 3px solid #9B72CF; background: #F7F2FD; }
.nav-tabs .nav-item:nth-child(6n+6) .nav-link.active { color: #B53A72; border-top: 3px solid #E0609A; background: #FDF1F6; }
.btn-outline-primary { border-radius: 999px; }
.btn-dl { border-radius: 999px; font-weight: 600; margin-bottom: .45rem; width: 100%; }
.btn-dl-1 { background: #E2F5EF; border: 1px solid #2A9D8F; color: #1E6F65; }
.btn-dl-2 { background: #E7EEFE; border: 1px solid #4C8BF5; color: #2F5FB8; }
.btn-dl-3 { background: #FDEBE4; border: 1px solid #E76F51; color: #A9472B; }
.btn-dl:hover { filter: brightness(0.97); }
.src-chip { display: inline-block; padding: .25rem .7rem; border-radius: 999px; font-size: .9rem; font-weight: 600;
            margin-top: .3rem; }
.src-builtin { background: #E2F5EF; color: #1E6F65; border: 1px solid #9ED8CC; }
.src-upload  { background: #FFF3D9; color: #8E6210; border: 1px solid #F0CF8A; }
.tiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: .8rem; margin-bottom: 1rem; }
.tile { border-radius: 14px; padding: .8rem 1rem; border: 1px solid; }
.tile-val { font-size: 2rem; font-weight: 800; line-height: 1.1; }
.tile-lab { font-size: .95rem; color: #50606E; }
.tile-1 { background: #E2F5EF; border-color: #A7DDD2; } .tile-1 .tile-val { color: #1E7F73; }
.tile-2 { background: #E7EEFE; border-color: #B5CBFA; } .tile-2 .tile-val { color: #2F66C8; }
.tile-3 { background: #FDEBE4; border-color: #F5BFAE; } .tile-3 .tile-val { color: #C4532F; }
.tile-4 { background: #FFF3D9; border-color: #F0CF8A; } .tile-4 .tile-val { color: #A8740F; }
.tile-5 { background: #F1E9FB; border-color: #D3BFEE; } .tile-5 .tile-val { color: #7447B0; }
.share { display: inline-block; width: 150px; height: 14px; background: #EEF2F5; border-radius: 7px;
         overflow: hidden; vertical-align: middle; }
.share > span { display: block; height: 100%; border-radius: 7px; }
.table-striped > tbody > tr:nth-of-type(odd) > * { --bs-table-accent-bg: #F4FBF9; }
.stem-l { color: #2A9D8F; font-weight: 700; } .stem-b { color: #9AA7B2; } .stem-r { color: #E76F51; font-weight: 600; }
.stem-h { color: #7447B0; }
.teach-note { background: #FFF8E6; border-left: 4px solid #E9A93A; padding: .65rem .95rem;
              border-radius: 6px; margin: .5rem 0 .9rem; font-size: 1.02rem; }
.warn-note  { background: #FDEEE8; border-left: 4px solid #E76F51; padding: .65rem .95rem;
              border-radius: 6px; margin: .4rem 0 .8rem; }
.ok-note    { background: #E9F6F2; border-left: 4px solid #2A9D8F; padding: .65rem .95rem;
              border-radius: 6px; margin: .4rem 0 .8rem; }
.rawvals { font-family: ui-monospace, Menlo, Consolas, monospace; font-size: 1.3rem;
           line-height: 2.1; background: #FBF9FF; padding: 1rem 1.2rem; border: 1px solid #E6DDF7;
           border-radius: 10px; }
.rawvals span { display: inline-block; min-width: 2.6em; text-align: center; margin: .1rem; padding: 0 .25rem;
                border-radius: 6px; font-weight: 700; }
.table { font-size: 1.05rem; }
.table thead th { background: #EEF1F3; }
.table tr:last-child td { font-weight: 600; }
.no-total .table tr:last-child td { font-weight: 400; }
.tally { font-family: ui-monospace, Menlo, Consolas, monospace; letter-spacing: .05em; color: #2A9D8F; font-weight: 700; }
.tally5 { text-decoration: line-through; text-decoration-thickness: 2px; text-decoration-color: #E76F51; }
.stem pre { font-size: 1.3rem; line-height: 1.5; background: #FBF9FF; border: 1px solid #E6DDF7; border-radius: 10px; padding: 1rem 1.2rem; }
.working p { font-family: ui-monospace, Menlo, Consolas, monospace; font-size: 1.05rem; margin-bottom: .35rem; }
.checks li { margin-bottom: .25rem; }
.big-q { font-size: 1.12rem; font-weight: 600; }
.steps { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: .7rem; margin-bottom: 1rem; }
.step { background: #FFFFFF; border: 1px solid #E3E9F0; border-radius: 14px; padding: .7rem .9rem; font-weight: 600; color: #33414D;
        box-shadow: 0 2px 8px rgba(80,100,130,.06); }
.step-n { display: inline-flex; width: 1.7rem; height: 1.7rem; border-radius: 999px; background: #E2F5EF; color: #1E7F73;
          align-items: center; justify-content: center; margin-right: .5rem; font-weight: 800; }
.muted { color: #6B7884; font-weight: 400; }
.scale-grid > div { display: grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap: .6rem; }
.scale-row { background: #FBFCFF; border: 1px solid #E3E9F0; border-radius: 10px; padding: .55rem .7rem; }
.scale-row .form-group { margin-bottom: .3rem; }
.scale-name { margin-bottom: .3rem; font-size: .95rem; }
.kind-chip { background: #E7EEFE; color: #2F5FB8; border-radius: 10px; padding: .5rem .8rem; font-weight: 600; margin-bottom: .7rem; }
.badge-v { display: inline-block; font-size: .78rem; font-weight: 800; border-radius: 999px; padding: .05rem .55rem; margin-right: .45rem; }
.badge-v.best { background: #E2F5EF; color: #1E7F73; border: 1px solid #9ED8CC; }
.badge-v.good { background: #E7EEFE; color: #2F5FB8; border: 1px solid #B5CBFA; }
#b_type .radio { margin-bottom: .45rem; }
#b_type, #b_type .shiny-options-group { width: 100%; }
.avoid-box { background: #FDEEE8; border-left: 4px solid #E76F51; border-radius: 6px; padding: .5rem .8rem; margin-top: .3rem; }
.avoid-box ul { margin: .2rem 0 0; padding-left: 1.1rem; }
.auto-checks { list-style: none; padding-left: 0; margin-top: .8rem; font-size: .95rem; }
.ck-ok { color: #2A9D8F; font-weight: 800; } .ck-no { color: #E76F51; font-weight: 800; }
.slide-list { padding-left: 1.3rem; margin: 0; }
.slide-list li { display: flex; justify-content: space-between; align-items: center; gap: .6rem; padding: .45rem .2rem;
                 border-bottom: 1px solid #EEF1F3; }
.slide-list li.slide-title { display: list-item; color: #50606E; }
.slide-acts { white-space: nowrap; }
.slide-act { display: inline-block; text-decoration: none; padding: .1rem .45rem; margin-left: .2rem; border-radius: 6px;
             background: #F1F4F7; color: #33414D; font-weight: 700; }
.slide-act:hover { background: #E7EEFE; }
.accordion-button { font-weight: 700; color: #1E6F65; }
.plot-tools { display: flex; flex-wrap: wrap; align-items: flex-start; gap: .4rem 1.2rem; margin: .35rem 0 .2rem; font-size: .92rem; }
.tool-link { font-weight: 600; color: #2F66C8; text-decoration: none; padding-top: .1rem; }
.tool-link:hover { text-decoration: underline; }
.code-box { flex: 1 1 100%; }
.code-box summary { cursor: pointer; font-weight: 600; color: #7447B0; display: inline-block; }
.code-box pre, .code-live pre { font-size: .92rem; line-height: 1.45; background: #FBF9FF; border: 1px solid #E6DDF7;
                                border-radius: 10px; padding: .8rem 1rem; white-space: pre; overflow-x: auto; margin-top: .4rem; }
.code-live pre { font-size: 1rem; min-height: 150px; white-space: pre-wrap; word-break: break-word; }
.btn-row { display: flex; flex-wrap: wrap; gap: .5rem; }
.btn-row .btn-dl { width: auto; }
table.layers { width: 100%; margin-top: .6rem; font-size: 1.02rem; }
table.layers td { padding: .3rem .4rem; border-bottom: 1px solid #EEF1F3; vertical-align: top; }
table.layers tr.on td:first-child { color: #2A9D8F; font-weight: 800; }
table.layers tr.off { color: #9AA7B2; }
.rule { border-left: 4px solid #E3E9F0; padding: .45rem .8rem; margin-bottom: .55rem; border-radius: 6px; background: #FBFCFF; }
.rule-on { border-left-color: #E9A93A; background: #FFF8E6; }
.caption-box { font-style: italic; background: #F7F9FB; border: 1px dashed #C9D3DC; border-radius: 8px;
               padding: .6rem .9rem; margin: .4rem 0 .8rem; }
.ck-list { margin: .4rem 0 0; padding-left: 1rem; }
.ck-hit { color: #1E7F73; } .ck-miss { color: #C4532F; } .ck-fa { color: #7A8791; }
.cap-rules li { margin-bottom: .25rem; }
.guide td, .guide th { vertical-align: top; }
.navbar .dropdown-menu { border-radius: 12px; border: 1px solid #E3E9F0; box-shadow: 0 6px 18px rgba(60,80,100,.12); }
.navbar .dropdown-item { font-weight: 600; color: #2F5FB8; }
.navbar .dropdown-item.active, .navbar .dropdown-item:active { background: #E7EEFE; color: #2F5FB8; }
"

cp_clear_js <- "
Shiny.addCustomMessageHandler('cp_clear', function(ids) {
  [].concat(ids).forEach(function(id) {
    document.querySelectorAll('input[name=\"' + id + '\"]').forEach(function(r) { r.checked = false; });
    Shiny.setInputValue(id, null);
  });
});
"

ui <- page_navbar(
  title = "Displaying and visualising data",
  id = "main",
  bg = "#EAF5F1",
  inverse = FALSE,
  fillable = FALSE,
  theme = bs_theme(version = 5, bg = "#FFFFFF", fg = "#1F2A33",
                   primary = "#2A9D8F", secondary = "#E76F51",
                   base_font = font_collection("Segoe UI", "Helvetica Neue", "Arial", "sans-serif"),
                   "font-size-base" = "1.02rem"),
  header = tags$head(tags$style(HTML(app_css)), tags$script(HTML(cp_clear_js))),
  sidebar = sidebar(
    title = "Class controls", width = 270, open = "desktop",
    sliderInput("fs", "Text size on plots", min = 12, max = 28, value = 17, step = 1),
    checkboxInput("show_notes", "Show teaching notes", TRUE),
    hr(),
    tags$strong("1. Download the data"),
    downloadButton("dl_xlsx", "Excel (.xlsx)", class = "btn-sm btn-dl btn-dl-1"),
    downloadButton("dl_csv", "CSV (.csv)", class = "btn-sm btn-dl btn-dl-2"),
    tags$strong("2. Upload it back"),
    fileInput("upload", NULL, accept = c(".xlsx", ".xls", ".csv"),
              buttonLabel = "Browse\u2026", placeholder = "Excel or CSV"),
    helpText("Excel: the sheet named Data is read (otherwise the first sheet)."),
    actionButton("use_builtin", "Back to the class dataset", class = "btn-sm btn-dl btn-dl-3"),
    uiOutput("data_info")
  ),

  # ---- The data --------------------------------------------------------------
  nav_panel(
    "The data",
    uiOutput("tiles"),
    card(
      hdr("Forty values. What can you actually see?", 1),
      layout_sidebar(
        sidebar = sidebar(
          selectInput("raw_var", "Variable", choices = lab_choices(init_all), selected = "Adherence"),
          radioButtons("raw_order", "Show values in",
                       c("Patient-ID order" = "id", "Sorted order" = "sorted"))
        ),
        uiOutput("raw_vals"),
        uiOutput("raw_caption"),
        note("Give the class ten seconds with the raw row. What is the most common value? ",
             "Do most patients do well or badly? Are the two arms different? None of this is readable by eye. ",
             "The move: stop looking at cases and start counting values.")
      )
    ),
    card(hdr("Variables, measurement scales and legitimate displays", 2), tableOutput("dict")),
    card(hdr("The data matrix: rows are patients, columns are variables", 3),
         div(class = "no-total", style = "max-height: 520px; overflow-y: auto;", tableOutput("data_tbl")))
  ),

  # ---- Ungrouped frequency table --------------------------------------------
  nav_panel(
    "Frequency table",
    card(
      hdr("Ungrouped frequency distribution: f, rf, cf and c%", 4),
      layout_sidebar(
        sidebar = sidebar(
          selectInput("ft_var", "Variable", choices = lab_choices(init_all), selected = "Adherence"),
          checkboxInput("ft_tally", "Show tally column", TRUE)
        ),
        uiOutput("ft_warn"),
        tableOutput("ft_tbl"),
        uiOutput("ft_read"),
        note("The tally is a teaching device, not a deliverable. Two checks every time: ",
             "\u03a3f = N, and the last c% is exactly 100.0. ",
             "Ask what 42.5 in the c% column means for adherence: 42.5% scored 3 or below, not 42.5% scored 3.")
      )
    )
  ),

  # ---- Grouped frequency table ----------------------------------------------
  nav_panel(
    "Grouped table",
    layout_sidebar(
      sidebar = sidebar(
        selectInput("gt_var", "Quantitative variable", choices = lab_choices(init_quant), selected = "SBP_Base"),
        numericInput("gt_w", "Class width", value = 5, min = 0.1, step = 1),
        numericInput("gt_start", "Lower limit of the first class", value = 140, step = 1),
        actionButton("gt_reset", "Use the suggested width and start", class = "btn-sm btn-outline-primary")
      ),
      layout_columns(
        col_widths = c(7, 5),
        card(hdr("Sturges\u2019 rule: a starting point, not an answer", 5), div(class = "working", uiOutput("gt_sturges"))),
        card(hdr("Six rules for class intervals", 6), uiOutput("gt_checks"))
      ),
      card(hdr("Grouped frequency distribution", 1), tableOutput("gt_tbl")),
      card(hdr("The same classes as a histogram", 2), plotx("gt_hist", "380px", code = TRUE)),
      note("Width is boundary to boundary, not upper minus lower stated limit: 140\u2013144 has width 5, not 4. ",
           "Try width 7 or a start above the minimum and watch the checks fail. Group for display; keep the raw data for analysis.")
    )
  ),

  # ---- Categorical displays ---------------------------------------------------
  nav_panel(
    "Categorical displays",
    navset_card_tab(
      nav_panel(
        "Bar chart",
        layout_sidebar(
          sidebar = sidebar(
            selectInput("bar_var", "Categorical variable", choices = lab_choices(init_cat), selected = "Education"),
            radioButtons("bar_order", "Bar order", c("Natural order" = "natural", "Sorted by frequency" = "freq")),
            radioButtons("bar_y", "Bar height", c("Count" = "n", "Percentage" = "pct"))
          ),
          uiOutput("bar_warn"),
          plotx("bar_plot", "420px", code = TRUE),
          note("The bar chart and its frequency table are the same object. Gaps between bars say \u201cseparate categories\u201d. ",
               "Never sort an ordinal variable by frequency \u2014 it throws away the order the scale carries.")
        )
      ),
      nav_panel(
        "Pie or bar?",
        layout_sidebar(
          sidebar = sidebar(
            selectInput("pie_var", "Categorical variable", choices = lab_choices(init_cat), selected = "Education"),
            checkboxInput("pie_lab", "Label slices with percentages", FALSE)
          ),
          layout_columns(col_widths = c(6, 6),
                         plotx("pie_plot", "420px", code = FALSE),
                         plotx("pie_bar", "420px", code = TRUE)),
          note("Ask the room which slice is second largest, first from the pie, then from the bars. ",
               "People judge lengths well and angles badly. Pies are acceptable for two or three categories that form a meaningful whole.")
        )
      ),
      nav_panel(
        "Clustered bars",
        layout_sidebar(
          sidebar = sidebar(
            selectInput("cl_var", "Variable on the axis", choices = lab_choices(init_cat), selected = "Adherence"),
            selectInput("cl_grp", "Cluster by", choices = lab_choices(init_cat), selected = "Arm"),
            radioButtons("cl_y", "Bar height", c("Counts" = "n", "% within each group" = "pct"))
          ),
          plotx("cl_plot", "440px", code = TRUE),
          note("Adherence by arm: every patient scoring 5 is in the intervention arm; every patient scoring 1 or 2 is in usual care. ",
               "Percentages within each group keep the comparison fair when group sizes differ.")
        )
      ),
      nav_panel(
        "Cross-tabulation",
        layout_sidebar(
          sidebar = sidebar(
            selectInput("ct_row", "Rows", choices = lab_choices(init_cat), selected = "Education"),
            selectInput("ct_col", "Columns", choices = lab_choices(init_cat), selected = "BP_Controlled"),
            radioButtons("ct_pct", "Show",
                         c("Counts" = "none", "Row %" = "row", "Column %" = "col", "% of grand total" = "total"),
                         selected = "none")
          ),
          tableOutput("ct_tbl"),
          uiOutput("ct_read"),
          note("Every cell has three possible percentages: of the row, of the column and of the grand total. ",
               "They answer different questions. Choose one deliberately and label it. Whether a gradient is real is a Session 10 question.")
        )
      )
    )
  ),

  # ---- Quantitative displays --------------------------------------------------
  nav_panel(
    "Quantitative displays",
    navset_card_tab(
      nav_panel(
        "Bar chart \u2260 histogram",
        layout_sidebar(
          sidebar = sidebar(
            selectInput("bh_var", "Quantitative variable", choices = lab_choices(init_quant), selected = "SBP_Base"),
            numericInput("bh_w", "Class width", value = 5, min = 0.1, step = 1)
          ),
          layout_columns(col_widths = c(6, 6),
                         plotx("bh_bar", "400px", code = FALSE),
                         plotx("bh_hist", "400px", code = FALSE)),
          note("The single most common chart error in student work. Gaps assert that the variable is categorical; ",
               "touching bars say this is one continuous scale, and each bar spans the class boundaries.")
        )
      ),
      nav_panel(
        "Histogram and bin width",
        layout_sidebar(
          sidebar = sidebar(
            selectInput("hw_var", "Quantitative variable", choices = lab_choices(init_quant), selected = "SBP_Base"),
            sliderInput("hw_w", "Bin width", min = 1, max = 17, value = 5, step = 1),
            checkboxInput("hw_compare", "Compare half, current and double width", FALSE)
          ),
          plotx("hw_plot", "430px", code = TRUE),
          note("The histogram is a choice you make, not a fact you find. Drag the slider from 1 to 17 and watch the shape change character. ",
               "Report the bin width, and distrust any histogram whose story changes with it.")
        )
      ),
      nav_panel(
        "Stem-and-leaf",
        layout_sidebar(
          sidebar = sidebar(
            selectInput("sl_var", "Quantitative variable", choices = lab_choices(init_quant), selected = "SBP_Base"),
            sliderInput("sl_scale", "Stem scale (length of the display)", min = 0.5, max = 3, value = 1, step = 0.5)
          ),
          div(class = "stem", uiOutput("sl_out")),
          note("Rotate it 90\u00b0 anticlockwise and you have the histogram \u2014 but every value is still readable. ",
               "Scale 1 splits each stem into two lines (leaves 0\u20134 and 5\u20139); scale 0.5 gives one line per stem.")
        )
      ),
      nav_panel(
        "Polygon and ogive",
        layout_sidebar(
          sidebar = sidebar(
            selectInput("po_var", "Quantitative variable", choices = lab_choices(init_quant), selected = "SBP_Base"),
            numericInput("po_w", "Class width", value = 5, min = 0.1, step = 1),
            selectInput("po_grp", "Overlay polygons by", choices = grp_choices(init_cat), selected = ""),
            sliderInput("po_p", "Read a percentile off the ogive", min = 1, max = 99, value = 50, step = 1)
          ),
          layout_columns(col_widths = c(6, 6),
                         plotx("po_poly", "400px", code = FALSE),
                         plotx("po_ogive", "400px", code = FALSE)),
          uiOutput("po_read"),
          note("The polygon joins class midpoints and is anchored at zero one class beyond each end \u2014 useful for overlaying two groups. ",
               "The ogive plots c% at each upper boundary; read across at 50% and down to find the median.")
        )
      ),
      nav_panel(
        "Boxplot and outliers",
        layout_sidebar(
          sidebar = sidebar(
            selectInput("bx_var", "Quantitative variable", choices = lab_choices(init_quant), selected = "Visits"),
            selectInput("bx_grp", "Side by side by", choices = grp_choices(init_cat), selected = ""),
            checkboxInput("bx_fence", "Show the 1.5 \u00d7 IQR fences", TRUE),
            checkboxInput("bx_pts", "Show individual patients", FALSE)
          ),
          plotx("bx_plot", "auto", code = TRUE),
          tableOutput("bx_tbl"),
          note("Fences: Q1 \u2212 1.5 \u00d7 IQR and Q3 + 1.5 \u00d7 IQR. Whiskers stop at the most extreme value still inside the fences. ",
               "Flagged is not wrong: the 82-year-old and the patient with 15 visits are real. The rule tells you where to look, not what to delete.")
        )
      )
    )
  ),

  # ---- Shape --------------------------------------------------------------------
  nav_panel(
    "Shape",
    layout_sidebar(
      sidebar = sidebar(
        selectInput("sh_var", "Variable", choices = lab_choices(init_num), selected = "Visits"),
        selectInput("sh_grp", "Also show each group of", choices = grp_choices(init_cat), selected = ""),
        numericInput("sh_w", "Bin width", value = 2, min = 0.1, step = 1),
        checkboxInput("sh_dens", "Overlay a smooth outline", FALSE)
      ),
      card(hdr("Shape, mean and median", 3), plotx("sh_plot", "auto", code = TRUE)),
      card(hdr("Numerical companions", 4), tableOutput("sh_tbl")),
      note("Skew is named for the tail, not the bulk. A long right tail pulls the mean above the median. ",
           "Try SBP at 6 months pooled, then split by treatment arm: two peaks nearly always mean two populations have been pooled. ",
           "Rule of thumb used here: |G1| < 0.5 approximately symmetric. The picture comes first; a skewness near zero does not rule out two peaks.")
    )
  ),

  # ---- Misleading graphs --------------------------------------------------------
  nav_panel(
    "Misleading graphs",
    navset_card_tab(
      nav_panel(
        "Truncated axis",
        layout_sidebar(
          sidebar = sidebar(
            selectInput("mx_measure", "Quantity plotted", choices = measure_choices(hyp_builtin), selected = "mean:SBP_6mo"),
            selectInput("mx_grp", "Grouped by", choices = lab_choices(init_cat), selected = "Arm"),
            sliderInput("mx_start", "Truncated axis starts at", min = 0, max = 130, value = 125, step = 1)
          ),
          layout_columns(col_widths = c(6, 6),
                         plotx("mx_honest", "400px", code = FALSE),
                         plotx("mx_trunc", "400px", code = TRUE)),
          uiOutput("mx_read"),
          note("Same data, same bars, one changed axis. A truncated axis is not always dishonest, but it must be declared, ",
               "and bar charts in particular should start at zero because the eye reads bar length.")
        )
      ),
      nav_panel(
        "Unequal class widths",
        layout_sidebar(
          sidebar = sidebar(
            selectInput("uq_var", "Quantitative variable", choices = lab_choices(init_quant), selected = "SBP_Base"),
            numericInput("uq_w", "Class width before merging", value = 5, min = 0.1, step = 1),
            sliderInput("uq_merge", "Merge the top classes into one", min = 2, max = 6, value = 4, step = 1),
            radioButtons("uq_h", "Bar height",
                         c("Count (misleading)" = "count", "Frequency density = f \u00f7 width" = "density"))
          ),
          plotx("uq_plot", "420px", code = FALSE),
          note("The eye reads area, not height. A wide merged class drawn at its full count looks far bigger than it is. ",
               "Either keep widths equal, or plot frequency density so that area is proportional to frequency.")
        )
      )
    )
  ),

  # ---- Visualisation ----------------------------------------------------------------
  nav_menu(
    "Visualisation",

    # -- Building a plot ------------------------------------------------------------
    nav_panel(
      "Building a plot",
      layout_sidebar(
        sidebar = sidebar(
          width = 250,
          selectInput("gg_x", "x", choices = lab_choices(init_quant), selected = "SBP_Base"),
          selectInput("gg_y", "y", choices = lab_choices(init_quant), selected = "SBP_6mo"),
          selectInput("gg_col", "Colour by", choices = grp_choices(init_cat), selected = "Arm"),
          radioButtons("gg_colmode", "Colour is",
                       c("Mapped inside aes(): varies with the data" = "mapped",
                         "Fixed outside aes(): one colour for all" = "fixed")),
          checkboxGroupInput("gg_layers", "Layers",
                             c("geom_point()" = "point", "geom_smooth(method = \"lm\")" = "smooth",
                               "labs(): labels and units" = "labs", "facet_wrap()" = "facet"),
                             selected = character(0)),
          selectInput("gg_facet", "Facet by", choices = facet_choices(hyp_builtin), selected = "Sex"),
          div(class = "btn-row",
              actionButton("gg_next", "Add the next layer", class = "btn-sm btn-dl btn-dl-1"),
              actionButton("gg_reset", "Back to the empty canvas", class = "btn-sm btn-dl btn-dl-3"))
        ),
        layout_columns(
          col_widths = c(7, 5),
          card(hdr("The plot, one layer at a time", 1), plotx("gg_plot", "470px", code = FALSE)),
          card(hdr("The code that draws it", 2),
               div(class = "code-live", verbatimTextOutput("gg_code")),
               uiOutput("gg_explain"))
        ),
        note("Start from the empty canvas and press \u201cAdd the next layer\u201d. Data and aes() alone give axes and nothing else; ",
             "each + adds a layer, and the order of geoms decides what is drawn on top. Without labs() the axes carry raw column names ",
             "and no units: an unlabelled axis is an incomplete figure.")
      )
    ),

    # -- Comparing groups ------------------------------------------------------------
    nav_panel(
      "Comparing groups",
      layout_sidebar(
        sidebar = sidebar(
          width = 250,
          selectInput("cg_y", "Outcome", choices = lab_choices(init_quant), selected = "SBP_6mo"),
          selectInput("cg_g", "Groups", choices = lab_choices(init_cat), selected = "Arm"),
          radioButtons("cg_err", "Whisker on the bar",
                       c("Standard error (SE)" = "se", "Standard deviation (SD)" = "sd")),
          checkboxInput("cg_boxpts", "Middle panel: boxplot with the points on top", FALSE),
          checkboxInput("cg_dup", "\u2026 and leave outlier.shape at its default", FALSE)
        ),
        layout_columns(
          col_widths = c(4, 4, 4),
          card(hdr("Bar and whisker (dynamite)", 2), plotx("cg_bar", "400px")),
          card(hdr("Boxplot", 3), plotx("cg_box", "400px")),
          card(hdr("Every patient, with the median", 1), plotx("cg_pts", "400px"))
        ),
        uiOutput("cg_read"),
        card(
          hdr("Different data, the same dynamite plot", 5),
          layout_columns(col_widths = c(5, 7),
                         plotx("cg_same_bar", "380px", code = FALSE),
                         plotx("cg_same_pts", "380px", code = FALSE)),
          tableOutput("cg_same_tbl"),
          actionButton("cg_reshuffle", "Make another look-alike", class = "btn-sm btn-dl btn-dl-2", style = "width:auto")
        ),
        note("All three top panels plot the same patients. Only the first hides that the arms overlap. ",
             "A mean is not an amount of anything, so a bar misleads; standard error is precision of the mean, not spread. ",
             "The bottom card builds a second group with the same n, mean and SD but two clusters: its dynamite plot is identical.")
      )
    ),

    # -- Two variables -----------------------------------------------------------------
    nav_panel(
      "Two variables",
      navset_card_tab(
        nav_panel(
          "Scatter plot",
          layout_sidebar(
            sidebar = sidebar(
              selectInput("sc_x", "x", choices = lab_choices(init_quant), selected = "Age"),
              selectInput("sc_y", "y", choices = lab_choices(init_quant), selected = "SBP_Base"),
              selectInput("sc_col", "Colour by", choices = grp_choices(init_cat), selected = ""),
              checkboxInput("sc_lm", "Add a straight-line fit", TRUE),
              checkboxInput("sc_log", "Log scale on both axes", FALSE)
            ),
            plotx("sc_plot", "460px"),
            uiOutput("sc_read"),
            note("Each point is one patient. Pearson\u2019s r measures straight-line association only, so it changes when the scale changes. ",
                 "A log scale helps when values are crushed against one end; here the ranges are narrow and it changes little.")
          )
        ),
        nav_panel(
          "Overplotting",
          layout_sidebar(
            sidebar = sidebar(
              selectInput("op_x", "x", choices = lab_choices(init_num), selected = "Visits"),
              selectInput("op_y", "y", choices = lab_choices(init_num), selected = "Adherence"),
              radioButtons("op_mode", "Draw",
                           c("Plain points" = "plain", "Jitter + transparency" = "jitter", "Point size = count" = "count")),
              sliderInput("op_jit", "Jitter width", min = 0.05, max = 0.45, value = 0.25, step = 0.05)
            ),
            plotx("op_plot", "460px"),
            uiOutput("op_read"),
            note("Both variables are whole numbers, so patients land on exactly the same spot. ",
                 "Jitter is a display device, never an analysis one: the displacement exists only in the picture.")
          )
        )
      )
    ),

    # -- Honest axes and colour --------------------------------------------------------
    nav_panel(
      "Axes and colour",
      navset_card_tab(
        nav_panel(
          "When may an axis not start at zero?",
          layout_sidebar(
            sidebar = sidebar(
              width = 250,
              radioButtons("ax_type", "Chart",
                           c("Bar chart of means" = "bar", "Means with 95% CI (points)" = "dot",
                             "Line over time: baseline to 6 months" = "line", "Boxplot" = "box")),
              conditionalPanel("input.ax_type != 'line'",
                               selectInput("ax_var", "Outcome", choices = lab_choices(init_quant), selected = "SBP_6mo")),
              selectInput("ax_g", "Groups", choices = lab_choices(init_cat), selected = "Arm"),
              sliderInput("ax_start", "Axis starts at", min = 0, max = 150, value = 0, step = 1),
              checkboxInput("ax_declare", "Declare the truncation on the figure", TRUE)
            ),
            layout_columns(
              col_widths = c(7, 5),
              card(hdr("The figure", 3), plotx("ax_plot", "430px"), uiOutput("ax_verdict")),
              card(hdr("The rule, and its one exception", 4), uiOutput("ax_rules"))
            ),
            note("The test is whether the visual encoding is length. A bar is read as a length, so its axis starts at zero. ",
                 "Points and lines are read as positions, so zero can be left out, but say so on the figure.")
          )
        ),
        nav_panel(
          "Colour and accessibility",
          layout_sidebar(
            sidebar = sidebar(
              width = 250,
              selectInput("ac_x", "Categories", choices = lab_choices(init_cat), selected = "Adherence"),
              selectInput("ac_g", "Coloured by", choices = lab_choices(init_cat), selected = "Arm"),
              radioButtons("ac_pal", "Palette",
                           c("Red and green hues" = "redgreen", "Light to dark (one hue)" = "lightness",
                             "Okabe\u2013Ito (colour-blind safe)" = "okabe")),
              checkboxInput("ac_lab", "Label the bars directly instead of a legend", FALSE),
              checkboxGroupInput("ac_views", "Show it as seen with",
                                 c("Normal vision", "Deuteranopia", "Protanopia", "Tritanopia", "Greyscale"),
                                 selected = c("Normal vision", "Deuteranopia", "Greyscale"))
            ),
            plotx("ac_plot", "auto"),
            uiOutput("ac_read"),
            note("About 1 man in 12 has a red\u2013green colour-vision deficiency, and most journals still print in greyscale. ",
                 "Vary lightness rather than hue, and never let colour alone carry the message: direct labels solve most problems at once.")
          )
        )
      )
    ),

    # -- Small multiples -----------------------------------------------------------------
    nav_panel(
      "Small multiples",
      layout_sidebar(
        sidebar = sidebar(
          width = 250,
          selectInput("fm_y", "Outcome", choices = lab_choices(init_quant), selected = "SBP_6mo"),
          selectInput("fm_g", "Groups (x-axis)", choices = lab_choices(init_cat), selected = "Arm"),
          selectInput("fm_f", "One panel per", choices = facet_choices(hyp_builtin), selected = "AgeGroup"),
          checkboxInput("fm_n", "Put the n in every panel", TRUE),
          checkboxInput("fm_med", "Mark each group median", TRUE),
          radioButtons("fm_scales", "Axes", c("Shared by every panel" = "fixed", "Free in each panel" = "free_y")),
          sliderInput("fm_min", "Flag panels with fewer patients than", min = 4, max = 20, value = 10, step = 1)
        ),
        card(hdr("Faceting: the same plot repeated across a third variable", 6), plotx("fm_plot", "440px")),
        uiOutput("fm_read"),
        card(hdr("What each panel holds", 1), tableOutput("fm_tbl")),
        note("The pattern across panels is the finding. Keep the axes common so the panels can be compared, and report the n in every panel. ",
             "Faceting invites subgroup reading; a panel of a handful of patients supports no comparison of its own.")
      )
    ),

    # -- Delivering the figure -------------------------------------------------------------
    nav_panel(
      "Delivering the figure",
      navset_card_tab(
        nav_panel(
          "Table, figure or neither",
          layout_columns(
            col_widths = c(5, 7),
            card(hdr("Choosing the vehicle", 4), uiOutput("tf_guide")),
            card(hdr("Baseline characteristics by group: this is a table", 1),
                 selectInput("t1_g", "Columns", choices = lab_choices(init_cat), selected = "Arm", width = "260px"),
                 div(class = "no-total", tableOutput("t1_tbl")),
                 uiOutput("t1_note"))
          ),
          note("A figure that could have been a sentence wastes a page. A table of numbers that could have been a figure wastes the reader. ",
               "Figures earn their place when a shape, a relationship or a comparison across many groups is the point.")
        ),
        nav_panel(
          "Spot the faults",
          layout_columns(
            col_widths = c(7, 5),
            card(hdr("A figure for review", 2),
                 plotOutput("ck_plot", height = "420px"),
                 uiOutput("ck_caption"),
                 actionButton("ck_new", "Draw a new figure", class = "btn-sm btn-dl btn-dl-1", style = "width:auto")),
            card(hdr("Eight checks before a figure leaves your desk", 5),
                 tags$p(class = "big-q", "Tick every check this figure fails."),
                 checkboxGroupInput("ck_marks", NULL, choiceNames = unname(ck_items), choiceValues = names(ck_items)),
                 actionButton("ck_reveal", "Reveal the faults", class = "btn-sm btn-dl btn-dl-3", style = "width:auto"),
                 uiOutput("ck_result"))
          ),
          note("Draw it, break it, fix it: it is much easier to spot a distortion once you have built one on purpose. ",
               "Each new figure carries two or three faults chosen at random.")
        ),
        nav_panel(
          "Captions",
          layout_sidebar(
            sidebar = sidebar(
              radioButtons("cap_fig", "Figure",
                           c("Every patient, with the median (Comparing groups)" = "cg",
                             "Scatter plot (Two variables)" = "sc",
                             "Small multiples" = "fm"))
            ),
            plotOutput("cap_plot", height = "400px"),
            uiOutput("cap_text"),
            note("The caption describes; the text argues. The figure and caption together should make sense to a reader who has not read the paper.")
          )
        )
      )
    )
  ),

  # ---- Build your slides ----------------------------------------------------------------
  nav_panel(
    "Build your slides",
    div(class = "steps",
        div(class = "step", span(class = "step-n", "1"), "Load your data in the sidebar (Excel or CSV) and check the scale of each variable"),
        div(class = "step", span(class = "step-n", "2"), "Choose what to show; the app ranks the displays that fit"),
        div(class = "step", span(class = "step-n", "3"), "Edit the labels and caption, then add the figure as a slide"),
        div(class = "step", span(class = "step-n", "4"), "Download the deck as PowerPoint or PDF")),
    accordion(
      id = "b_acc", open = c("b_a1", "b_a2"), multiple = TRUE,
      accordion_panel(
        "1 \u00b7 Your data and its measurement scales", value = "b_a1",
        uiOutput("b_data_info"),
        tags$p(class = "muted", "The scale decides which displays are legitimate. Correct any the app has guessed wrongly, ",
               "for example a 1\u20135 score stored as numbers is ordinal, and a 0/1 code is nominal. For an ordinal variable, give the order of its levels."),
        div(class = "scale-grid", uiOutput("b_scales"))
      ),
      accordion_panel(
        "2 \u00b7 What do you want to show?", value = "b_a2",
        layout_columns(
          col_widths = c(4, 8),
          div(
            selectInput("b_v1", "Main variable (the outcome)", choices = NULL),
            selectInput("b_v2", "Compare by, or plot against (optional)", choices = NULL),
            selectInput("b_f", "One panel per (optional)", choices = NULL),
            textInput("b_unit", "Each row of the data is one\u2026", value = "patient")
          ),
          div(uiOutput("b_kind"), uiOutput("b_options"))
        )
      )
    ),
    layout_columns(
      col_widths = c(8, 4),
      card(hdr("3 \u00b7 Preview", 1),
           plotOutput("b_plot", height = "460px"),
           uiOutput("b_caption_show")),
      card(hdr("Edit the figure", 2),
           textInput("b_title", "Slide title"),
           textInput("b_xlab", "x-axis label"),
           textInput("b_ylab", "y-axis label"),
           uiOutput("b_extra"),
           textAreaInput("b_caption", "Caption", rows = 4),
           textAreaInput("b_notes", "Speaker notes (optional)", rows = 2),
           actionLink("b_refill", "Reset the labels and caption to the suggested wording"),
           uiOutput("b_checks"),
           actionButton("b_add", "Add this figure as a slide", class = "btn btn-dl btn-dl-1", style = "margin-top:.6rem"))
    ),
    card(hdr("4 \u00b7 Your presentation", 5),
         layout_columns(
           col_widths = c(5, 7),
           div(textInput("b_deck_title", "Presentation title", value = "Displaying the data"),
               textInput("b_deck_sub", "Subtitle (presenter, date)", value = ""),
               div(class = "btn-row",
                   downloadButton("b_pptx", "PowerPoint (.pptx)", class = "btn-dl btn-dl-1"),
                   downloadButton("b_pdf", "PDF", class = "btn-dl btn-dl-2"),
                   actionButton("b_clear", "Remove all slides", class = "btn-dl btn-dl-3"))),
           uiOutput("b_slides")
         )),
    note("Only displays that suit the scales of the chosen variables are offered; those to avoid are listed with the reason. ",
         "Each slide keeps a copy of the figure as it was when added, so later changes to the data do not alter it.")
  ),

  # ---- Checkpoints ----------------------------------------------------------------
  nav_panel(
    "Checkpoints",
    uiOutput("cp_score"),
    div(class = "btn-row", style = "margin-bottom: .8rem;",
        actionButton("cp_reveal", "Reveal all answers", class = "btn-sm btn-dl btn-dl-2", style = "width:auto"),
        actionButton("cp_reset", "Clear all answers", class = "btn-sm btn-dl btn-dl-3", style = "width:auto")),
    do.call(navset_card_tab, unname(lapply(names(cp_sets), function(nm) {
      ids <- cp_sets[[nm]]
      nav_panel(sprintf("%s (%d\u2013%d)", nm, min(ids), max(ids)),
                do.call(layout_columns, c(list(col_widths = c(6, 6)), lapply(ids, function(i) {
                  cp <- cps[[i]]
                  card(
                    class = "cp-card",
                    hdr(paste("Checkpoint", i, "of", length(cps)), (i - 1) %% 6 + 1),
                    div(class = "big-q", cp$q),
                    radioButtons(paste0("cp_", i), NULL,
                                 choiceNames = paste0(names(cp$opts), ".  ", cp$opts),
                                 choiceValues = names(cp$opts), selected = character(0)),
                    actionButton(paste0("cp_go_", i), "Check answer", class = "btn-sm btn-outline-primary"),
                    uiOutput(paste0("cp_fb_", i))
                  )
                }))))
    })))
  )
)

# -----------------------------------------------------------------------------
# 7. Server
# -----------------------------------------------------------------------------
server <- function(input, output, session) {

  dat <- reactiveVal(hyp_builtin)
  src <- reactiveVal("builtin")
  fs  <- reactive(input$fs)

  # renderPlot plus a PNG download of the same figure
  plot_out <- function(id, expr, ..., env = parent.frame()) {
    force(env)
    e <- substitute(expr)
    output[[id]] <- renderPlot(eval(e, env), ...)
    output[[paste0(id, "_png")]] <- downloadHandler(
      filename = function() paste0(id, ".png"),
      content = function(file) {
        p <- isolate(eval(e, env))
        ggsave(file, p, width = 9, height = 5.6, dpi = 200, bg = "white")
      })
  }

  observeEvent(input$upload, {
    f <- input$upload
    df <- tryCatch(read_any(f$datapath, f$name), error = function(e) NULL)
    if (is.null(df) || ncol(df) < 2 || nrow(df) < 3) {
      showNotification("That file could not be read: it needs a header row and at least two columns.", type = "error")
      return()
    }
    dat(df); src(f$name)
    showNotification(sprintf("Loaded %s: %d rows \u00d7 %d columns. Every tab now uses this file.",
                             f$name, nrow(df), ncol(df)), type = "message", duration = 6)
  })
  observeEvent(input$use_builtin, { dat(hyp_builtin); src("builtin") })

  output$dl_xlsx <- downloadHandler(
    filename = function() "hypertension-dataset.xlsx",
    content = function(file) writexl::write_xlsx(list(Data = dat(), Dictionary = dict_df(dat())), file)
  )
  output$dl_csv <- downloadHandler(
    filename = function() "hypertension-dataset.csv",
    content = function(file) write.csv(dat(), file, row.names = FALSE, na = "")
  )

  output$data_info <- renderUI({
    df <- dat()
    tagList(
      if (src() == "builtin") span(class = "src-chip src-builtin", "Class dataset (built in)")
      else span(class = "src-chip src-upload", paste("Uploaded:", src())),
      helpText(sprintf("%d rows \u00d7 %d columns", nrow(df), ncol(df)))
    )
  })

  output$tiles <- renderUI({
    df <- dat()
    allv <- vars_of(df, c("nominal", "ordinal", "ratio", "interval"))
    div(class = "tiles",
        tile(nrow(df), "patients (rows)", 1),
        tile(length(allv), "variables to analyse", 2),
        tile(length(vars_of(df, "nominal")), "nominal", 3),
        tile(length(vars_of(df, "ordinal")), "ordinal", 4),
        tile(length(vars_of(df, c("ratio", "interval"))), "interval / ratio", 5))
  })

  # refresh every selector when the data change
  observeEvent(dat(), {
    df <- dat()
    allv <- vars_of(df, c("nominal", "ordinal", "ratio", "interval"))
    catv <- vars_of(df, c("nominal", "ordinal"))
    qv   <- vars_of(df, c("ratio", "interval"))
    numv <- allv[vapply(allv, function(v) is.numeric(df[[v]]), logical(1))]
    upd <- function(id, ch, pref) updateSelectInput(session, id, choices = lab_choices(ch), selected = pick(ch, pref))
    upd("raw_var", allv, "Adherence"); upd("ft_var", allv, "Adherence")
    upd("gt_var", qv, "SBP_Base"); upd("bh_var", qv, "SBP_Base"); upd("hw_var", qv, "SBP_Base")
    upd("sl_var", qv, "SBP_Base"); upd("po_var", qv, "SBP_Base"); upd("bx_var", qv, "Visits")
    upd("uq_var", qv, "SBP_Base"); upd("sh_var", numv, "Visits")
    upd("bar_var", catv, "Education"); upd("pie_var", catv, "Education")
    upd("cl_var", catv, "Adherence"); upd("cl_grp", catv, "Arm")
    upd("ct_row", catv, "Education"); upd("ct_col", catv, "BP_Controlled"); upd("mx_grp", catv, "Arm")
    for (id in c("po_grp", "bx_grp", "sh_grp")) updateSelectInput(session, id, choices = grp_choices(catv), selected = "")
    upd("gg_x", qv, "SBP_Base"); upd("gg_y", qv, "SBP_6mo"); upd("cg_y", qv, "SBP_6mo")
    upd("sc_x", qv, "Age"); upd("sc_y", qv, "SBP_Base"); upd("op_x", numv, "Visits"); upd("op_y", numv, "Adherence")
    upd("ax_var", qv, "SBP_6mo"); upd("fm_y", qv, "SBP_6mo")
    for (id in c("cg_g", "ax_g", "fm_g", "t1_g", "ac_g")) upd(id, catv, "Arm")
    upd("ac_x", catv, "Adherence")
    for (id in c("gg_col", "sc_col")) updateSelectInput(session, id, choices = grp_choices(catv), selected = pick(catv, "Arm"))
    updateSelectInput(session, "sc_col", selected = "")
    fc <- facet_choices(df)
    updateSelectInput(session, "gg_facet", choices = fc, selected = pick(unname(fc), "Sex"))
    updateSelectInput(session, "fm_f", choices = fc, selected = pick(unname(fc), "AgeGroup"))
    mc <- measure_choices(df)
    updateSelectInput(session, "mx_measure", choices = mc, selected = pick(unname(mc), "mean:SBP_6mo"))
  }, ignoreInit = TRUE)

  numvar <- function(v) {
    req(v, v %in% names(dat()))
    x <- dat()[[v]]
    validate(need(is.numeric(x), "Choose a numeric variable."))
    x[!is.na(x)]
  }

  # ---- The data ---------------------------------------------------------------
  output$raw_vals <- renderUI({
    v <- req(input$raw_var); df <- dat(); x <- df[[v]]
    if (input$raw_order == "sorted") x <- sort(x)
    lv <- if (is.numeric(x)) sort(unique(x[!is.na(x)])) else get_levels(df, v)
    cols <- setNames(ramp_cols(length(lv)), as.character(lv))
    tint <- function(h) grDevices::adjustcolor(h, alpha.f = 0.16)
    div(class = "rawvals", HTML(paste(vapply(as.character(x), function(z) {
      if (is.na(z)) return("<span style='color:#9AA7B2'>NA</span>")
      sprintf("<span style='color:%s;background:%s'>%s</span>", cols[[z]], tint(cols[[z]]), htmltools::htmlEscape(z))
    }, ""), collapse = " ")))
  })
  output$raw_caption <- renderUI({
    v <- req(input$raw_var); df <- dat(); x <- df[[v]]
    tags$p(style = "margin-top:.6rem; color:#55616B;",
           sprintf("%s \u2014 all %d patients, %s. %d distinct values; scale: %s.",
                   vlab(v), sum(!is.na(x)),
                   if (input$raw_order == "sorted") "sorted" else "in patient-ID order",
                   length(unique(x[!is.na(x)])), scale_of(df, v)))
  })
  output$dict <- renderTable({
    d <- dict_df(dat())
    sc_col <- c(nominal = "#C4532F", ordinal = "#A8740F", ratio = "#2F66C8", interval = "#2F66C8", identifier = "#7A8791")
    d$Scale <- sprintf("<span style='color:%s;font-weight:700'>%s</span>", sc_col[d$Scale], d$Scale)
    d
  }, striped = TRUE, spacing = "s", sanitize.text.function = function(x) x)
  output$data_tbl <- renderTable(dat(), striped = TRUE, spacing = "xs", digits = 1)

  # ---- Frequency table -------------------------------------------------------
  output$ft_warn <- renderUI({
    v <- req(input$ft_var); df <- dat(); sc <- scale_of(df, v)
    nd <- length(unique(df[[v]][!is.na(df[[v]])]))
    if (sc == "nominal") {
      div(class = "warn-note", strong(vlab(v)), " is nominal: there is no \u201cat or below\u201d, so the cumulative columns are left empty.")
    } else if (sc %in% c("ratio", "interval") && nd > 15) {
      div(class = "warn-note", sprintf("%d distinct values in %d patients: this table is the raw list wearing a table\u2019s clothes. Use the Grouped table tab.",
                                       nd, sum(!is.na(df[[v]]))))
    }
  })
  output$ft_tbl <- renderTable({
    v <- req(input$ft_var); df <- dat(); sc <- scale_of(df, v)
    f <- as_cat(df, v); tb <- table(f); n <- sum(tb)
    cum_ok <- sc != "nominal"
    out <- data.frame(Value = htmltools::htmlEscape(names(tb)), check.names = FALSE)
    if (isTRUE(input$ft_tally)) out$Tally <- paste0("<span class='tally'>", vapply(as.integer(tb), tally_html, ""), "</span>")
    out$f <- as.character(as.integer(tb))
    out[["rf (%)"]] <- fmt(100 * as.integer(tb) / n, 1)
    out$cf <- if (cum_ok) as.character(cumsum(as.integer(tb))) else ""
    out[["c%"]] <- if (cum_ok) fmt(100 * cumsum(as.integer(tb)) / n, 1) else ""
    out$Share <- share_bar(100 * as.integer(tb) / n, ramp_cols(length(tb)))
    tot <- out[1, ]; tot[1, ] <- ""
    tot$Value <- "Total"; tot$f <- as.character(n); tot[["rf (%)"]] <- fmt(100, 1)
    tot$cf <- "\u2014"; tot[["c%"]] <- "\u2014"
    rbind(out, tot)
  }, sanitize.text.function = function(x) x, align = "l", striped = TRUE)
  output$ft_read <- renderUI({
    v <- req(input$ft_var); df <- dat(); sc <- scale_of(df, v)
    f <- as_cat(df, v); tb <- table(f); n <- sum(tb); p <- 100 * as.integer(tb) / n
    modes <- names(tb)[tb == max(tb)]
    items <- list(tags$li(sprintf("Modal category: %s (f = %d, %.1f%%).", paste(modes, collapse = " and "), max(tb), max(p))))
    if (sc != "nominal" && length(tb) >= 4) {
      k <- length(tb); cp <- cumsum(p)
      items <- c(items, list(
        tags$li(sprintf("%.1f%% are at or below \u201c%s\u201d (c%% column).", cp[2], names(tb)[2])),
        tags$li(sprintf("%.1f%% are in the top two categories (\u201c%s\u201d or \u201c%s\u201d).",
                        p[k - 1] + p[k], names(tb)[k - 1], names(tb)[k]))))
    }
    items <- c(items, list(tags$li(sprintf("Checks: \u03a3f = %d = N; final c%% = %s.", n,
                                           if (sc != "nominal") "100.0" else "not applicable for a nominal variable"))))
    div(class = "ok-note", strong("Reading the table"), tags$ul(style = "margin-bottom:0", items))
  })

  # ---- Grouped table ------------------------------------------------------------
  set_width_start <- function(wid, sid, v) {
    x <- numvar(v); w <- default_w(x); u <- 10^-dec_places(x)
    updateNumericInput(session, wid, value = w, min = u, step = if (u < 1) u * 10 else 1)
    if (!is.null(sid)) updateNumericInput(session, sid, value = default_start(x, w), step = if (u < 1) u * 10 else 1)
  }
  observeEvent(input$gt_var, set_width_start("gt_w", "gt_start", input$gt_var))
  observeEvent(input$gt_reset, set_width_start("gt_w", "gt_start", input$gt_var))
  observeEvent(input$bh_var, set_width_start("bh_w", NULL, input$bh_var))
  observeEvent(input$po_var, set_width_start("po_w", NULL, input$po_var))
  observeEvent(input$uq_var, set_width_start("uq_w", NULL, input$uq_var))
  observeEvent(input$sh_var, set_width_start("sh_w", NULL, input$sh_var))
  observeEvent(input$hw_var, {
    x <- numvar(input$hw_var); u <- 10^-dec_places(x); w <- default_w(x)
    mx <- max(u * 2, ceiling((max(x) - min(x)) / 2))
    updateSliderInput(session, "hw_w", min = u, max = mx, value = min(w, mx), step = u)
  })

  gt <- reactive({
    x <- numvar(input$gt_var)
    w <- input$gt_w; s <- input$gt_start
    validate(need(is.numeric(w) && !is.na(w) && w > 0, "Enter a positive class width."),
             need(is.numeric(s) && !is.na(s), "Enter the lower limit of the first class."))
    grouped(x, w, s)
  })

  output$gt_sturges <- renderUI({
    x <- numvar(input$gt_var); s <- sturges(x); d <- dec_places(x); u <- 10^-d
    tagList(
      tags$p(sprintf("k = 1 + 3.322 \u00d7 log\u2081\u2080(%d) = 1 + 3.322 \u00d7 %.3f = %.2f", s$n, log10(s$n), s$k)),
      tags$p(sprintf("range = %s \u2212 %s = %s", fmt(s$max, d), fmt(s$min, d), fmt(s$rng, d))),
      tags$p(sprintf("w \u2248 %s / %.2f = %.2f  \u2192  convenient width %s", fmt(s$rng, d), s$k, s$w_raw,
                     format(nice_width(s$w_raw, u)))),
      tags$p(sprintf("\u2192 about %d or %d classes", floor(s$k), ceiling(s$k))),
      div(style = "font-family: inherit; color:#55616B; margin-top:.4rem;",
          "Sturges assumes roughly normal data. For a skewed variable such as clinic visits, follow the shape, not the formula.")
    )
  })

  output$gt_checks <- renderUI({
    g <- gt()
    lead <- g$w / 10^floor(log10(g$w))
    conv <- any(abs(lead - c(1, 2, 2.5, 5, 10)) < 1e-9)
    chk <- function(ok, txt) tags$li(HTML(paste0(if (ok) "<span style='color:#2A9D8F;font-weight:700'>\u2713</span> "
                                                  else "<span style='color:#E76F51;font-weight:700'>\u2717</span> ", txt)))
    inner_empty <- if (g$K > 2) sum(g$f[2:(g$K - 1)] == 0) else 0
    tags$ul(class = "checks", style = "list-style:none; padding-left:0;",
            chk(TRUE, "Mutually exclusive: boundaries are half a unit outside the stated limits, so no value can sit on two classes"),
            chk(g$uncovered == 0, sprintf("Exhaustive: %d of %d values fall in a class", g$n - g$uncovered, g$n)),
            chk(TRUE, sprintf("Equal width: every class is %s wide", format(g$w))),
            chk(g$K >= 5 && g$K <= 15, sprintf("Between 5 and 15 classes: %d classes", g$K)),
            chk(conv, sprintf("Convenient width: %s", format(g$w))),
            chk(g$start <= g$min && g$f[1] > 0 && g$f[g$K] > 0,
                sprintf("Starts at or below the minimum (%s) with no empty class at either end", format(g$min))),
            if (inner_empty > 0) tags$li(style = "color:#55616B", sprintf("(%d empty class%s inside the range)", inner_empty, if (inner_empty > 1) "es" else ""))
    )
  })

  output$gt_tbl <- renderTable(grouped_df(gt()), striped = TRUE, align = "l", sanitize.text.function = function(x) x)
  plot_out("gt_hist", {
    g <- gt()
    hist_plot(g, vlab(input$gt_var), fs()) +
      labs(title = sprintf("%d classes of width %s; bars span the class boundaries", g$K, format(g$w)))
  }, res = 96)

  # ---- Bar chart -------------------------------------------------------------------
  cat_bar <- function(d, xlab, y = "n", fs) {
    d$y <- if (y == "pct") d$pct else d$n
    d$lab <- if (y == "pct") sprintf("%.1f%%", d$pct) else as.character(d$n)
    ggplot(d, aes(level, y, fill = level)) +
      geom_col(width = 0.62) +
      geom_text(aes(label = lab), vjust = -0.45, size = fs / 3) +
      scale_fill_manual(values = rep_len(pal, nrow(d)), guide = "none") +
      scale_y_continuous(expand = expansion(mult = c(0, 0.14))) +
      labs(x = xlab, y = if (y == "pct") "Percentage of patients" else "Frequency (f)") +
      theme_class(fs)
  }
  output$bar_warn <- renderUI({
    v <- req(input$bar_var); sc <- scale_of(dat(), v)
    if (input$bar_order == "freq" && sc == "ordinal")
      div(class = "warn-note", strong(vlab(v)), " is ordinal. Sorting by frequency destroys the order the scale carries \u2014 put it back in natural order.")
    else if (input$bar_order == "freq" && sc == "nominal")
      div(class = "ok-note", strong(vlab(v)), " is nominal, so sorting by frequency loses nothing: there is no natural order.")
  })
  plot_out("bar_plot", {
    v <- req(input$bar_var); d <- cat_counts(as_cat(dat(), v))
    if (input$bar_order == "freq") d$level <- factor(d$level, levels = d$level[order(-d$n)])
    cat_bar(d, vlab(v), input$bar_y, fs()) + labs(title = paste(vlab(v), "\u2014 all", sum(d$n), "patients"))
  }, res = 96)

  # ---- Pie vs bar -------------------------------------------------------------------
  plot_out("pie_plot", {
    v <- req(input$pie_var); d <- cat_counts(as_cat(dat(), v))
    p <- ggplot(d, aes(x = 1, y = n, fill = level)) +
      geom_col(width = 1, colour = "white", linewidth = 1) +
      coord_polar(theta = "y", direction = -1) +
      scale_fill_manual(values = rep_len(pal, nrow(d))) +
      labs(fill = NULL, title = "Pie chart: compare the angles") +
      theme_void(base_size = fs()) +
      theme(legend.position = "bottom", plot.title = element_text(face = "bold", hjust = 0.5))
    if (isTRUE(input$pie_lab))
      p <- p + geom_text(aes(label = sprintf("%.1f%%", pct)), position = position_stack(vjust = 0.5), size = fs() / 3)
    p
  }, res = 96)
  plot_out("pie_bar", {
    v <- req(input$pie_var); d <- cat_counts(as_cat(dat(), v))
    cat_bar(d, vlab(v), "n", fs()) + labs(title = "Bar chart: compare the lengths")
  }, res = 96)

  # ---- Clustered bars -----------------------------------------------------------------
  plot_out("cl_plot", {
    v <- req(input$cl_var); g <- req(input$cl_grp)
    validate(need(v != g, "Choose two different variables."))
    df <- dat()
    tb <- as.data.frame(table(x = as_cat(df, v), g = as_cat(df, g)))
    tot <- ave(tb$Freq, tb$g, FUN = sum)
    pct <- input$cl_y == "pct"
    tb$y <- if (pct) 100 * tb$Freq / tot else tb$Freq
    tb$lab <- if (pct) sprintf("%.0f%%", tb$y) else as.character(tb$Freq)
    ggplot(tb, aes(x, y, fill = g)) +
      geom_col(position = position_dodge(width = 0.8), width = 0.75) +
      geom_text(aes(label = lab), position = position_dodge(width = 0.8), vjust = -0.4, size = fs() / 3.4) +
      scale_fill_manual(values = rep_len(pal, nlevels(tb$g))) +
      scale_y_continuous(expand = expansion(mult = c(0, 0.14))) +
      labs(x = vlab(v), y = if (pct) paste("% within each", tolower(vlab(g))) else "Frequency (f)",
           fill = vlab(g), title = paste(vlab(v), "by", tolower(vlab(g)))) +
      theme_class(fs())
  }, res = 96)

  # ---- Cross-tabulation -----------------------------------------------------------------
  ct <- reactive({
    r <- req(input$ct_row); cc <- req(input$ct_col)
    validate(need(r != cc, "Choose two different variables."))
    table(as_cat(dat(), r), as_cat(dat(), cc))
  })
  output$ct_tbl <- renderTable({
    tb <- ct(); mode <- input$ct_pct
    rs <- rowSums(tb); cs <- colSums(tb); N <- sum(tb)
    pc <- function(v) sprintf("%.1f%%", v)
    if (mode == "none") {
      body <- cbind(matrix(as.character(tb), nrow(tb)), as.character(rs))
      last <- c(as.character(cs), as.character(N))
    } else if (mode == "row") {
      body <- cbind(matrix(pc(100 * prop.table(tb, 1)), nrow(tb)), sprintf("100.0%% (n = %d)", rs))
      last <- c(pc(100 * cs / N), sprintf("100.0%% (n = %d)", N))
    } else if (mode == "col") {
      body <- cbind(matrix(pc(100 * prop.table(tb, 2)), nrow(tb)), pc(100 * rs / N))
      last <- c(sprintf("100.0%% (n = %d)", cs), sprintf("100.0%% (n = %d)", N))
    } else {
      body <- cbind(matrix(pc(100 * tb / N), nrow(tb)), pc(100 * rs / N))
      last <- c(pc(100 * cs / N), sprintf("100.0%% (n = %d)", N))
    }
    out <- rbind(body, last)
    out <- data.frame(c(rownames(tb), "Total"), out, check.names = FALSE, stringsAsFactors = FALSE)
    names(out) <- c(vlab(input$ct_row), colnames(tb), "Row total")
    out
  }, striped = TRUE, align = "l")
  output$ct_read <- renderUI({
    tb <- ct(); mode <- input$ct_pct; r <- vlab(input$ct_row); cc <- vlab(input$ct_col)
    q <- switch(mode,
                none = "Counts: the raw material. Percentages come next \u2014 but of what base?",
                row = sprintf("Row %%: what proportion of each %s group falls in each %s category?", tolower(r), tolower(cc)),
                col = sprintf("Column %%: within each %s category, how is %s distributed?", tolower(cc), tolower(r)),
                total = sprintf("%% of grand total: what share of all %d patients sits in each cell?", sum(tb)))
    extra <- NULL
    if (mode == "row" && ncol(tb) >= 2) {
      last <- colnames(tb)[ncol(tb)]; rp <- 100 * prop.table(tb, 1)[, ncol(tb)]; rs <- rowSums(tb)
      lo <- which.min(rp); hi <- which.max(rp)
      extra <- tags$p(style = "margin:.3rem 0 0", sprintf("\u201c%s\u201d ranges from %.1f%% (%s, %d of %d) to %.1f%% (%s, %d of %d). Always report the base.",
                                                        last, rp[lo], rownames(tb)[lo], tb[lo, ncol(tb)], rs[lo],
                                                        rp[hi], rownames(tb)[hi], tb[hi, ncol(tb)], rs[hi]))
    }
    div(class = "ok-note", q, extra)
  })

  # ---- Bar chart vs histogram -----------------------------------------------------
  bh <- reactive({
    x <- numvar(input$bh_var); w <- input$bh_w
    validate(need(is.numeric(w) && !is.na(w) && w > 0, "Enter a positive class width."))
    grouped(x, w)
  })
  plot_out("bh_bar", {
    g <- bh()
    d <- data.frame(cls = factor(paste0(fmt(g$lower, g$d), "\u2013", fmt(g$upper, g$d)),
                                 levels = paste0(fmt(g$lower, g$d), "\u2013", fmt(g$upper, g$d))), f = g$f)
    ggplot(d, aes(cls, f, fill = cls)) + geom_col(width = 0.6) +
      scale_fill_manual(values = grDevices::colorRampPalette(c("#F4A08A", "#E76F51", "#C94F6D"))(nrow(d)), guide = "none") +
      scale_y_continuous(expand = expansion(mult = c(0, 0.08)), breaks = integer_breaks) +
      labs(x = paste(vlab(input$bh_var), "(stated class limits)"), y = "Frequency (f)",
           title = "Wrong: gaps between bars", subtitle = "Gaps assert separate categories") +
      theme_class(fs()) + theme(axis.text.x = element_text(angle = 40, hjust = 1))
  }, res = 96)
  plot_out("bh_hist", {
    hist_plot(bh(), paste(vlab(input$bh_var), "(class boundaries)"), fs()) +
      labs(title = "Right: bars touch", subtitle = "One continuous scale")
  }, res = 96)

  # ---- Histogram and bin width -----------------------------------------------------
  plot_out("hw_plot", {
    x <- numvar(input$hw_var); w <- req(input$hw_w); u <- 10^-dec_places(x)
    ws <- if (isTRUE(input$hw_compare)) unique(c(max(u, w / 2), w, 2 * w)) else w
    rects <- do.call(rbind, lapply(ws, function(wi) {
      g <- grouped(x, wi)
      data.frame(xmin = g$lb, xmax = g$ub, f = g$f, col = ramp_cols(g$K),
                 panel = sprintf("Width %s  (%d classes)", format(signif(wi, 4)), g$K))
    }))
    rects$panel <- factor(rects$panel, levels = unique(rects$panel))
    p <- ggplot(rects) +
      geom_rect(aes(xmin = xmin, xmax = xmax, ymin = 0, ymax = f, fill = col), colour = "white", linewidth = 0.5) +
      scale_fill_identity() +
      scale_y_continuous(expand = expansion(mult = c(0, 0.08)), breaks = integer_breaks) +
      labs(x = vlab(input$hw_var), y = "Frequency (f)") + theme_class(fs())
    if (length(ws) > 1) p + facet_wrap(~panel, nrow = 1) else p + labs(title = levels(rects$panel))
  }, res = 96)

  # ---- Stem-and-leaf -----------------------------------------------------------------------
  output$sl_out <- renderUI({
    x <- numvar(input$sl_var)
    lines <- capture.output(stem(x, scale = input$sl_scale))
    html <- vapply(lines, function(l) {
      m <- regmatches(l, regexec("^(\\s*-?[0-9.]+) \\| ?(.*)$", l))[[1]]
      if (length(m) == 3)
        sprintf("<span class='stem-l'>%s</span> <span class='stem-b'>|</span> <span class='stem-r'>%s</span>",
                htmltools::htmlEscape(m[2]), htmltools::htmlEscape(m[3]))
      else sprintf("<span class='stem-h'>%s</span>", htmltools::htmlEscape(l))
    }, "")
    tags$pre(HTML(paste(c(sprintf("<span class='stem-h'><b>%s</b>  (n = %d)</span>",
                                  htmltools::htmlEscape(vlab(input$sl_var)), length(x)), html), collapse = "\n")))
  })

  # ---- Polygon and ogive ---------------------------------------------------------------------
  po <- reactive({
    x <- numvar(input$po_var); w <- input$po_w
    validate(need(is.numeric(w) && !is.na(w) && w > 0, "Enter a positive class width."))
    grouped(x, w)
  })
  plot_out("po_poly", {
    g <- po(); v <- input$po_var; grp <- input$po_grp
    xs <- c(g$mid[1] - g$w, g$mid, g$mid[g$K] + g$w)
    if (is.null(grp) || grp == "") {
      d <- data.frame(x = xs, y = c(0, g$f, 0))
      hd <- data.frame(xmin = g$lb, xmax = g$ub, f = g$f)
      ggplot() +
        geom_rect(data = hd, aes(xmin = xmin, xmax = xmax, ymin = 0, ymax = f), fill = "#D8EFEA", colour = "white") +
        geom_line(data = d, aes(x, y), colour = pal[1], linewidth = 1.2) +
        geom_point(data = d, aes(x, y), colour = pal[1], size = 3) +
        scale_y_continuous(expand = expansion(mult = c(0, 0.08)), breaks = integer_breaks) +
        labs(x = paste(vlab(v), "(class midpoints)"), y = "Frequency (f)", title = "Frequency polygon") +
        theme_class(fs())
    } else {
      df <- dat(); gf <- as_cat(df, grp); xv <- df[[v]]
      d <- do.call(rbind, lapply(levels(gf), function(l) {
        xi <- xv[gf == l & !is.na(gf) & !is.na(xv)]
        idx <- findInterval(xi, g$brks); f <- tabulate(idx[idx >= 1 & idx <= g$K], g$K)
        data.frame(x = xs, y = 100 * c(0, f, 0) / max(1, length(xi)), grp = sprintf("%s (n = %d)", l, length(xi)))
      }))
      ggplot(d, aes(x, y, colour = grp)) + geom_line(linewidth = 1.2) + geom_point(size = 3) +
        scale_colour_manual(values = rep_len(pal, length(unique(d$grp)))) +
        scale_y_continuous(expand = expansion(mult = c(0, 0.08))) +
        labs(x = paste(vlab(v), "(class midpoints)"), y = "% within group", colour = NULL,
             title = paste("Polygons overlaid by", tolower(vlab(grp)))) +
        theme_class(fs())
    }
  }, res = 96)
  ogive_read <- reactive({
    g <- po(); y <- c(0, 100 * g$cf / g$n)
    xv <- approx(y, g$brks, xout = input$po_p, ties = min)$y
    list(g = g, y = y, xv = xv)
  })
  plot_out("po_ogive", {
    o <- ogive_read(); g <- o$g; p <- input$po_p
    d <- data.frame(x = g$brks, y = o$y)
    ggplot(d, aes(x, y)) +
      geom_area(fill = "#FDEBE4", alpha = 0.9) +
      annotate("segment", x = g$brks[1], xend = o$xv, y = p, yend = p, colour = col_median, linetype = "dashed", linewidth = 0.8) +
      annotate("segment", x = o$xv, xend = o$xv, y = p, yend = 0, colour = col_median, linetype = "dashed", linewidth = 0.8) +
      geom_line(colour = pal[2], linewidth = 1.2) + geom_point(colour = pal[2], size = 3) +
      annotate("point", x = o$xv, y = p, colour = col_median, size = 4) +
      annotate("label", x = o$xv, y = p, label = sprintf("P%d \u2248 %s", p, fmt(o$xv, g$d + 1)),
               hjust = -0.1, vjust = 1.3, size = fs() / 3.2, label.size = 0, fill = "#EEF4FE") +
      scale_y_continuous(limits = c(0, 100), breaks = seq(0, 100, 25), expand = expansion(mult = c(0, 0.02))) +
      labs(x = paste(vlab(input$po_var), "(upper class boundaries)"), y = "Cumulative %", title = "Ogive") +
      theme_class(fs())
  }, res = 96)
  output$po_read <- renderUI({
    o <- ogive_read(); x <- numvar(input$po_var); p <- input$po_p
    div(class = "ok-note",
        sprintf("Reading across at %d%% and down: the %d%s percentile from the grouped ogive is about %s. From the raw values it is %s. The ogive estimate is only as fine as the classes.",
                p, p, if (p %% 10 == 1 && p != 11) "st" else if (p %% 10 == 2 && p != 12) "nd" else if (p %% 10 == 3 && p != 13) "rd" else "th",
                fmt(o$xv, o$g$d + 1), fmtn(quantile(x, p / 100, type = 7))))
  })

  # ---- Boxplot ------------------------------------------------------------------------------
  bx_data <- reactive({
    v <- req(input$bx_var); df <- dat(); x <- df[[v]]
    validate(need(is.numeric(x), "Choose a numeric variable."))
    grp <- input$bx_grp
    g <- if (is.null(grp) || grp == "") factor(rep("All patients", nrow(df))) else as_cat(df, grp)
    keep <- !is.na(x) & !is.na(g)
    data.frame(val = x[keep], g = droplevels(g[keep]))
  })
  bx_stats <- reactive({
    d <- bx_data()
    do.call(rbind, lapply(levels(d$g), function(l) {
      x <- d$val[d$g == l]; q <- quantile(x, c(.25, .5, .75), type = 7); iqr <- q[3] - q[1]
      lf <- q[1] - 1.5 * iqr; uf <- q[3] + 1.5 * iqr
      fl <- sort(x[x < lf | x > uf])
      data.frame(g = l, n = length(x), min = min(x), q1 = q[1], med = q[2], q3 = q[3], max = max(x),
                 iqr = iqr, lf = lf, uf = uf, flagged = if (length(fl)) paste(fl, collapse = ", ") else "none",
                 row.names = NULL)
    }))
  })
  plot_out("bx_plot", {
    d <- bx_data(); s <- bx_stats(); v <- input$bx_var
    s$yid <- seq_len(nrow(s))
    single <- nrow(s) == 1
    p <- ggplot(d, aes(x = val, y = g)) +
      geom_boxplot(aes(fill = g), width = 0.45, coef = 1.5, alpha = 0.35, colour = "#3A4A57",
                   outlier.colour = col_mean, outlier.size = 3.5) +
      scale_fill_manual(values = rep_len(pal, nrow(s)), guide = "none") +
      scale_x_continuous(expand = expansion(mult = 0.14)) +
      labs(x = vlab(v), y = NULL) + theme_class(fs()) +
      theme(panel.grid.major.y = element_blank(), panel.grid.major.x = element_line(colour = "#ECEFF1"))
    if (isTRUE(input$bx_pts)) {
      dq <- merge(d, s[, c("g", "q1", "med", "q3", "lf", "uf")], by = "g")
      dq$quarter <- ifelse(dq$val < dq$lf | dq$val > dq$uf, "Flagged",
                    ifelse(dq$val < dq$q1, "Lowest quarter", ifelse(dq$val < dq$med, "Second quarter",
                    ifelse(dq$val <= dq$q3, "Third quarter", "Top quarter"))))
      dq$quarter <- factor(dq$quarter, levels = names(quarter_cols))
      p <- p + geom_jitter(data = dq, aes(colour = quarter), height = 0.1, width = 0, alpha = 0.85, size = 2.6) +
        scale_colour_manual(values = quarter_cols, name = NULL, drop = TRUE)
    }
    if (isTRUE(input$bx_fence)) {
      p <- p +
        geom_segment(data = s, aes(x = lf, xend = lf, y = yid - 0.38, yend = yid + 0.38), inherit.aes = FALSE,
                     colour = col_fence, linetype = "dashed", linewidth = 0.9) +
        geom_segment(data = s, aes(x = uf, xend = uf, y = yid - 0.38, yend = yid + 0.38), inherit.aes = FALSE,
                     colour = col_fence, linetype = "dashed", linewidth = 0.9) +
        geom_text(data = s, aes(x = uf, y = yid + 0.43, label = paste("upper fence", fmtn(uf))), inherit.aes = FALSE,
                  colour = col_fence, size = fs() / 3.6, vjust = 0, hjust = 1) +
        geom_text(data = s, aes(x = lf, y = if (single) yid + 0.43 else yid - 0.43, label = paste("lower fence", fmtn(lf))),
                  inherit.aes = FALSE, colour = col_fence, size = fs() / 3.6, vjust = if (single) 0 else 1, hjust = 0)
    }
    p <- p + theme(legend.position = "bottom")
    if (single) {
      five <- data.frame(x = c(s$min, s$q1, s$med, s$q3, s$max),
                         lab = c("Min", "Q1", "Median", "Q3", "Max"),
                         y = c(0.64, 0.40, 0.64, 0.40, 0.64))
      five$txt <- paste0(five$lab, "\n", fmtn(five$x))
      p <- p + geom_text(data = five, aes(x = x, y = y, label = txt), inherit.aes = FALSE,
                         size = fs() / 3.4, colour = "#33414D", lineheight = 0.9) +
        scale_y_discrete(expand = expansion(add = c(0.85, 0.6)))
    }
    p
  }, res = 96, height = function() {
    g <- input$bx_grp
    if (is.null(g) || g == "") 420 else 220 + 150 * length(get_levels(dat(), g))
  })
  output$bx_tbl <- renderTable({
    s <- bx_stats()
    data.frame(Group = s$g, n = s$n, Min = fmtn(s$min), Q1 = fmtn(s$q1), Median = fmtn(s$med), Q3 = fmtn(s$q3),
               Max = fmtn(s$max), IQR = fmtn(s$iqr), "Lower fence" = fmtn(s$lf), "Upper fence" = fmtn(s$uf),
               "Flagged values" = s$flagged, check.names = FALSE)
  }, striped = TRUE, align = "l")

  # ---- Shape ----------------------------------------------------------------------------------
  sh_data <- reactive({
    v <- req(input$sh_var); df <- dat(); x <- df[[v]]
    validate(need(is.numeric(x), "Choose a numeric variable."))
    grp <- input$sh_grp
    out <- data.frame(val = x[!is.na(x)], panel = "All patients")
    if (!is.null(grp) && grp != "") {
      gf <- as_cat(df, grp)
      for (l in levels(gf)) {
        xi <- x[gf == l & !is.na(gf) & !is.na(x)]
        if (length(xi)) out <- rbind(out, data.frame(val = xi, panel = paste0(vlab(grp), ": ", l)))
      }
    }
    out$panel <- factor(out$panel, levels = unique(out$panel))
    out
  })
  plot_out("sh_plot", {
    d <- sh_data(); w <- input$sh_w
    validate(need(is.numeric(w) && !is.na(w) && w > 0, "Enter a positive bin width."))
    g0 <- grouped(d$val, w)
    rects <- do.call(rbind, lapply(levels(d$panel), function(pn) {
      xi <- d$val[d$panel == pn]; idx <- findInterval(xi, g0$brks)
      data.frame(xmin = g0$lb, xmax = g0$ub, f = tabulate(idx[idx >= 1 & idx <= g0$K], g0$K), panel = pn)
    }))
    rects$panel <- factor(rects$panel, levels = levels(d$panel))
    cent <- do.call(rbind, lapply(levels(d$panel), function(pn) {
      xi <- d$val[d$panel == pn]
      data.frame(panel = pn, stat = c("Mean", "Median"), value = c(mean(xi), median(xi)))
    }))
    cent$panel <- factor(cent$panel, levels = levels(d$panel))
    fills <- c("#8FCFC3", rep_len(pal[-1], max(0, nlevels(d$panel) - 1)))
    p <- ggplot() +
      geom_rect(data = rects, aes(xmin = xmin, xmax = xmax, ymin = 0, ymax = f, fill = panel), colour = "white", linewidth = 0.5) +
      geom_vline(data = cent, aes(xintercept = value, colour = stat, linetype = stat), linewidth = 1.1) +
      scale_fill_manual(values = setNames(fills, levels(d$panel)), guide = "none") +
      scale_colour_manual(values = c(Mean = col_mean, Median = col_median), name = NULL) +
      scale_linetype_manual(values = c(Mean = "solid", Median = "longdash"), name = NULL) +
      scale_y_continuous(expand = expansion(mult = c(0, 0.08)), breaks = integer_breaks) +
      labs(x = vlab(input$sh_var), y = "Frequency (f)") + theme_class(fs())
    if (isTRUE(input$sh_dens))
      p <- p + geom_density(data = d, aes(x = val, y = after_stat(count) * w), colour = "#3A4A57", linewidth = 0.9,
                            fill = NA, adjust = 1)
    if (nlevels(d$panel) > 1) p <- p + facet_wrap(~panel, ncol = 1)
    p
  }, res = 96, height = function() {
    g <- input$sh_grp
    if (is.null(g) || g == "") 440 else 300 + 230 * length(get_levels(dat(), g))
  })
  output$sh_tbl <- renderTable({
    d <- sh_data()
    do.call(rbind, lapply(levels(d$panel), function(pn) {
      xi <- d$val[d$panel == pn]; s <- skew_G1(xi)
      data.frame(Panel = pn, n = length(xi), Mean = fmt(mean(xi), 2), Median = fmtn(median(xi)),
                 "Mean \u2212 median" = fmt(mean(xi) - median(xi), 2),
                 "Skewness (G1)" = if (is.na(s)) "\u2014" else fmt(s, 2),
                 "Shape (rule of thumb)" = shape_verdict(s), check.names = FALSE)
    }))
  }, striped = TRUE, align = "l")

  # ---- Truncated axis -----------------------------------------------------------------------
  mx_vals <- reactive({
    m <- req(input$mx_measure); grp <- req(input$mx_grp); df <- dat()
    type <- sub(":.*", "", m); v <- sub("^[^:]*:", "", m)
    req(v %in% names(df), grp %in% names(df))
    gf <- as_cat(df, grp)
    vals <- vapply(levels(gf), function(l) {
      xi <- df[[v]][gf == l & !is.na(gf)]
      if (type == "mean") mean(xi, na.rm = TRUE) else 100 * mean(xi == "Yes", na.rm = TRUE)
    }, numeric(1))
    list(d = data.frame(g = factor(levels(gf), levels = levels(gf)), y = vals), type = type, v = v, grp = grp,
         ylab = if (type == "mean") paste("Mean", vlab(v)) else paste("% Yes \u2014", vlab(v)))
  })
  observeEvent(mx_vals(), {
    lo <- min(mx_vals()$d$y, na.rm = TRUE)
    mx <- max(0, floor(lo * 0.98))
    updateSliderInput(session, "mx_start", min = 0, max = mx, value = floor(lo * 0.9))
  })
  mx_plot <- function(start, title) {
    m <- mx_vals(); d <- m$d; top <- max(d$y) * 1.08
    lab <- if (m$type == "mean") fmt(d$y, 1) else sprintf("%.0f%%", d$y)
    ggplot(d, aes(g, y, fill = g)) + geom_col(width = 0.55) +
      geom_text(aes(label = lab), vjust = -0.4, size = fs() / 3) +
      scale_fill_manual(values = rep_len(pal, nrow(d)), guide = "none") +
      coord_cartesian(ylim = c(start, top), expand = FALSE) +
      labs(x = vlab(m$grp), y = m$ylab, title = title) + theme_class(fs())
  }
  plot_out("mx_honest", mx_plot(0, "Axis from zero"), res = 96)
  plot_out("mx_trunc", {
    s <- req(input$mx_start)
    mx_plot(s, sprintf("Axis starting at %s", format(s))) +
      theme(axis.line.y = element_line(colour = col_mean, linewidth = 1.2))
  }, res = 96)
  output$mx_read <- renderUI({
    m <- mx_vals(); y <- m$d$y; s <- input$mx_start
    hi <- max(y); lo <- min(y)
    req(lo > 0, hi > s, lo > s)
    real <- hi / lo; look <- (hi - s) / (lo - s)
    div(class = if (look > 1.5 * real) "warn-note" else "ok-note",
        sprintf("In the data the larger bar is %.2f times the smaller. On the truncated axis it looks %.2f times as tall.", real, look))
  })

  # ---- Unequal class widths ----------------------------------------------------------------
  observeEvent(list(input$uq_var, input$uq_w), {
    x <- numvar(input$uq_var); w <- input$uq_w; req(is.numeric(w), !is.na(w), w > 0)
    K <- grouped(x, w)$K
    if (K >= 3) updateSliderInput(session, "uq_merge", min = 2, max = K - 1, value = min(max(2, input$uq_merge), K - 1))
  })
  plot_out("uq_plot", {
    x <- numvar(input$uq_var); w <- input$uq_w
    validate(need(is.numeric(w) && !is.na(w) && w > 0, "Enter a positive class width."))
    g <- grouped(x, w)
    validate(need(g$K >= 3, "This width leaves fewer than three classes \u2014 choose a smaller width."))
    m <- min(max(2, input$uq_merge), g$K - 1); keep <- g$K - m
    d <- data.frame(xmin = c(g$lb[seq_len(keep)], g$lb[keep + 1]),
                    xmax = c(g$ub[seq_len(keep)], g$ub[g$K]),
                    f = c(g$f[seq_len(keep)], sum(g$f[(keep + 1):g$K])))
    d$width <- d$xmax - d$xmin
    d$merged <- c(rep("Equal-width classes", keep), "Merged wide class")
    dens <- input$uq_h == "density"
    d$h <- if (dens) d$f / d$width else d$f
    d$lab <- if (dens) sprintf("f = %d\nwidth %s", d$f, format(d$width)) else sprintf("f = %d", d$f)
    ggplot(d) +
      geom_rect(aes(xmin = xmin, xmax = xmax, ymin = 0, ymax = h, fill = merged), colour = "white", linewidth = 0.7) +
      geom_text(aes(x = (xmin + xmax) / 2, y = h, label = lab), vjust = -0.3, size = fs() / 3.5, lineheight = 0.9) +
      scale_fill_manual(values = c("Equal-width classes" = pal[1], "Merged wide class" = pal[2]), name = NULL) +
      scale_y_continuous(expand = expansion(mult = c(0, 0.2))) +
      scale_x_continuous(breaks = c(d$xmin, d$xmax[nrow(d)]), labels = fmt(c(d$xmin, d$xmax[nrow(d)]), g$d + 1)) +
      labs(x = vlab(input$uq_var), y = if (dens) "Frequency density (f \u00f7 width)" else "Frequency (f)",
           title = if (dens) "Area now represents frequency" else "Height = count: the wide class is exaggerated") +
      theme_class(fs())
  }, res = 96)

  # =====================================================================================
  # Visualisation
  # =====================================================================================
  dat_plus <- reactive(add_derived(dat()))
  catf <- function(df, v) { req(v %in% names(df)); as_cat(df, v) }

  # ---- Building a plot ------------------------------------------------------------------
  layer_order <- c("point", "smooth", "labs", "facet")
  observeEvent(input$gg_next, {
    cur <- input$gg_layers; nxt <- setdiff(layer_order, cur)
    if (length(nxt)) updateCheckboxGroupInput(session, "gg_layers", selected = c(cur, nxt[1]))
  })
  observeEvent(input$gg_reset, updateCheckboxGroupInput(session, "gg_layers", selected = character(0)))

  gg_state <- reactive({
    df <- dat_plus(); x <- req(input$gg_x); y <- req(input$gg_y)
    req(x %in% names(df), y %in% names(df))
    L <- input$gg_layers; if (is.null(L)) L <- character(0)
    cg <- input$gg_col
    mapped <- !is.null(cg) && cg != "" && cg %in% names(df) && input$gg_colmode == "mapped"
    fc <- input$gg_facet
    facet <- "facet" %in% L && !is.null(fc) && fc %in% names(df)
    list(df = df, x = x, y = y, L = L, cg = cg, mapped = mapped, fc = fc, facet = facet)
  })
  gg_obj <- reactive({
    s <- gg_state(); d <- s$df
    d$.x <- d[[s$x]]; d$.y <- d[[s$y]]
    if (s$mapped) d$.col <- as_cat(d, s$cg)
    p <- ggplot(d, if (s$mapped) aes(.x, .y, colour = .col) else aes(.x, .y)) + geom_blank() + theme_class(fs())
    if ("point" %in% s$L)
      p <- p + if (s$mapped) geom_point(size = 3.2, alpha = 0.8) else geom_point(size = 3.2, alpha = 0.8, colour = pal[1])
    if ("smooth" %in% s$L)
      p <- p + if (s$mapped) geom_smooth(method = "lm", formula = y ~ x, linewidth = 1.1, fill = "#E3E8EE")
               else geom_smooth(method = "lm", formula = y ~ x, linewidth = 1.1, colour = pal[2], fill = "#F6DCD3")
    if (s$mapped) p <- p + scale_colour_manual(values = rep_len(pal, nlevels(d$.col)))
    p <- if ("labs" %in% s$L) p + labs(x = vlab(s$x), y = vlab(s$y), colour = if (s$mapped) vlab(s$cg))
         else p + labs(x = s$x, y = s$y, colour = if (s$mapped) s$cg)
    if (s$facet) {
      d$.f <- with_n(as_cat(d, s$fc)); p <- p %+% d + facet_wrap(~.f)
    }
    if (!length(s$L))
      p <- p + annotate("label", x = mean(range(d$.x, na.rm = TRUE)), y = mean(range(d$.y, na.rm = TRUE)),
                        label = "Data + aes(): the axes exist,\nbut nothing is drawn yet", size = fs() / 3,
                        label.size = 0, fill = "#FFF8E6", colour = "#8E6210")
    p
  })
  plot_out("gg_plot", gg_obj(), res = 96)
  output$gg_code <- renderText({
    s <- gg_state(); df <- s$df
    lines <- c()
    if (s$facet && s$fc == "AgeGroup")
      lines <- c(lines, "hyp$AgeGroup <- cut(hyp$Age, c(-Inf, 49, 59, Inf), labels = c(\"<50\", \"50-59\", \"60+\"))\n")
    if (s$mapped) lines <- c(lines, lev_code(df, s$cg))
    aesl <- if (s$mapped) sprintf("aes(x = %s, y = %s, colour = %s)", s$x, s$y, s$cg) else sprintf("aes(x = %s, y = %s)", s$x, s$y)
    body <- sprintf("ggplot(hyp, %s)", aesl)
    if ("point" %in% s$L) body <- c(body, if (s$mapped) "geom_point(size = 3, alpha = 0.8)"
                                        else sprintf("geom_point(size = 3, alpha = 0.8, colour = %s)", qq(pal[1])))
    if ("smooth" %in% s$L) body <- c(body, "geom_smooth(method = \"lm\")")
    if ("labs" %in% s$L) body <- c(body, sprintf("labs(x = %s, y = %s%s)", qq(vlab(s$x)), qq(vlab(s$y)),
                                                 if (s$mapped) paste0(", colour = ", qq(vlab(s$cg))) else ""))
    if (s$facet) body <- c(body, sprintf("facet_wrap(~ %s)", s$fc))
    paste0(code_head, paste(lines, collapse = ""), "\n", paste(body, collapse = " +\n  "))
  })
  output$gg_explain <- renderUI({
    s <- gg_state()
    row <- function(on, part, txt) tags$tr(class = if (on) "on" else "off",
                                           tags$td(if (on) "\u2713" else "\u2013"), tags$td(tags$b(part)), tags$td(txt))
    tags$table(class = "layers",
      row(TRUE, "Data", sprintf("hyp: %d patients", nrow(s$df))),
      row(TRUE, "Aesthetics", paste0("x = ", s$x, ", y = ", s$y, if (s$mapped) paste0(", colour = ", s$cg) else "")),
      row(any(c("point", "smooth") %in% s$L), "Geometry",
          paste(c(if ("point" %in% s$L) "points", if ("smooth" %in% s$L) "fitted line with 95% band"), collapse = " + ")),
      row("labs" %in% s$L, "Labels", if ("labs" %in% s$L) "axis titles with units" else "raw column names, no units"),
      row(s$facet, "Facets", if (s$facet) paste("one panel per", tolower(vlab(s$fc))) else "optional"))
  })

  # ---- Comparing groups ------------------------------------------------------------------
  cg_d <- reactive({
    df <- dat(); v <- req(input$cg_y); g <- req(input$cg_g)
    req(v %in% names(df), g %in% names(df))
    validate(need(is.numeric(df[[v]]), "Choose a numeric outcome."))
    gf <- as_cat(df, g); keep <- !is.na(df[[v]]) & !is.na(gf)
    d <- data.frame(val = df[[v]][keep], g = droplevels(gf[keep]))
    d$gn <- with_n(d$g, "\n")
    d
  })
  cg_sum <- reactive({
    d <- cg_d()
    do.call(rbind, lapply(levels(d$gn), function(l) {
      x <- d$val[d$gn == l]
      data.frame(gn = l, n = length(x), mean = mean(x), sd = sd(x), se = sd(x) / sqrt(length(x)),
                 min = min(x), max = max(x), med = median(x))
    }))
  })
  dynamite <- function(s, err, ylab, fs, top = NULL) {
    s$e <- if (err == "se") s$se else s$sd
    s$gn <- factor(s$gn, levels = s$gn)
    p <- ggplot(s, aes(gn, mean, fill = gn)) +
      geom_col(width = 0.55) +
      geom_errorbar(aes(ymin = mean, ymax = mean + e), width = 0.18, linewidth = 0.9, colour = "#33414D") +
      scale_fill_manual(values = rep_len(pal, nrow(s)), guide = "none") +
      scale_y_continuous(expand = expansion(mult = c(0, 0.06))) +
      scale_x_discrete(labels = wrap_first) +
      labs(x = NULL, y = ylab) + theme_class(fs)
    if (!is.null(top)) p <- p + coord_cartesian(ylim = c(0, top))
    p
  }
  cg_bar_obj <- reactive({
    s <- cg_sum()
    dynamite(s, input$cg_err, paste("Mean", vlab(input$cg_y)), fs() * 0.85) +
      labs(subtitle = sprintf("Mean + 1 %s", toupper(input$cg_err)))
  })
  cg_box_obj <- reactive({
    d <- cg_d()
    p <- ggplot(d, aes(gn, val, fill = gn)) + scale_fill_manual(values = rep_len(pal, nlevels(d$gn)), guide = "none") +
      scale_x_discrete(labels = wrap_first) +
      labs(x = NULL, y = vlab(input$cg_y)) + theme_class(fs() * 0.85)
    if (isTRUE(input$cg_boxpts)) {
      p <- p + if (isTRUE(input$cg_dup)) geom_boxplot(width = 0.5, alpha = 0.3, outlier.colour = col_mean, outlier.size = 3.5)
               else geom_boxplot(width = 0.5, alpha = 0.3, outlier.shape = NA)
      set.seed(11)
      p <- p + geom_jitter(width = 0.15, height = 0, alpha = 0.6, size = 2.4, colour = "#33414D")
    } else {
      p <- p + geom_boxplot(width = 0.5, alpha = 0.45, outlier.colour = col_mean, outlier.size = 3.5)
    }
    p
  })
  cg_pts_obj <- reactive({
    d <- cg_d(); s <- cg_sum(); s$gn <- factor(s$gn, levels = levels(d$gn))
    set.seed(7)
    ggplot(d, aes(gn, val, colour = gn)) +
      geom_jitter(width = 0.15, height = 0, alpha = 0.75, size = 2.8) +
      geom_errorbar(data = s, aes(x = gn, ymin = med, ymax = med), inherit.aes = FALSE,
                    width = 0.45, linewidth = 1.3, colour = "#33414D") +
      scale_colour_manual(values = rep_len(pal, nlevels(d$gn)), guide = "none") +
      scale_x_discrete(labels = wrap_first) +
      labs(x = NULL, y = vlab(input$cg_y)) +
      theme_class(fs() * 0.85)
  })
  plot_out("cg_bar", cg_bar_obj(), res = 96)
  plot_out("cg_box", cg_box_obj(), res = 96)
  plot_out("cg_pts", cg_pts_obj(), res = 96)
  cg_code_pre <- reactive(paste0(code_head, lev_code(dat(), input$cg_g), "\n"))
  output$cg_bar_code <- renderText({
    v <- input$cg_y; g <- input$cg_g; e <- input$cg_err
    paste0(cg_code_pre(),
           sprintf("summ <- aggregate(%s ~ %s, data = hyp, FUN = function(x) c(m = mean(x), %s))\n", v, g,
                   if (e == "se") "e = sd(x) / sqrt(length(x))" else "e = sd(x)"),
           "summ <- do.call(data.frame, summ)\n\n",
           sprintf("ggplot(summ, aes(x = %s, y = %s.m)) +\n  geom_col(width = 0.55) +\n  geom_errorbar(aes(ymin = %s.m, ymax = %s.m + %s.e), width = 0.18) +\n  labs(x = NULL, y = %s)\n",
                   g, v, v, v, v, qq(paste("Mean", vlab(v)))),
           "\n# Shown here to be criticised: prefer the boxplot or the points.")
  })
  output$cg_box_code <- renderText({
    v <- input$cg_y; g <- input$cg_g
    body <- if (isTRUE(input$cg_boxpts))
      c(sprintf("geom_boxplot(width = 0.5%s)", if (isTRUE(input$cg_dup)) "" else ", outlier.shape = NA"),
        "geom_jitter(width = 0.15, height = 0, alpha = 0.6)")
    else "geom_boxplot(width = 0.5)"
    paste0(cg_code_pre(), sprintf("ggplot(hyp, aes(x = %s, y = %s)) +\n  ", g, v),
           paste(c(body, sprintf("labs(x = NULL, y = %s)", qq(vlab(v)))), collapse = " +\n  "),
           if (isTRUE(input$cg_boxpts) && isTRUE(input$cg_dup))
             "\n\n# Without outlier.shape = NA the boxplot draws its outliers and geom_jitter draws\n# them again: the same patient appears twice in one figure." else "")
  })
  output$cg_pts_code <- renderText({
    v <- input$cg_y; g <- input$cg_g
    paste0(cg_code_pre(), sprintf("ggplot(hyp, aes(x = %s, y = %s)) +\n", g, v),
           "  geom_jitter(width = 0.15, height = 0, alpha = 0.75) +\n",
           "  stat_summary(fun = median, fun.min = median, fun.max = median,\n               geom = \"crossbar\", width = 0.45) +\n",
           sprintf("  labs(x = NULL, y = %s)", qq(vlab(v))))
  })
  output$cg_read <- renderUI({
    s <- cg_sum(); req(nrow(s) >= 2)
    d <- cg_d(); dp <- dec_places(d$val)
    rng <- paste(sprintf("%s: %s to %s", sub("\\s*\\(n = .*", "", s$gn), fmt(s$min, dp), fmt(s$max, dp)), collapse = "; ")
    ov_lo <- max(s$min); ov_hi <- min(s$max)
    ov <- if (ov_lo <= ov_hi) sum(d$val >= ov_lo & d$val <= ov_hi) else 0
    dup <- if (isTRUE(input$cg_boxpts) && isTRUE(input$cg_dup)) {
      fl <- sum(vapply(split(d$val, d$gn), function(x) { q <- quantile(x, c(.25, .75)); i <- diff(q)
                                                         sum(x < q[1] - 1.5 * i | x > q[2] + 1.5 * i) }, numeric(1)))
      if (fl > 0) tags$p(style = "margin:.3rem 0 0", sprintf("Middle panel: %d flagged patient%s drawn twice, once by the boxplot and once by the jitter.", fl, if (fl > 1) "s are" else " is"))
    }
    div(class = "warn-note",
        sprintf("The bars show %d numbers, a mean and a whisker per group. The patients range %s. %d of %d patients sit inside the range the groups share%s.",
                2 * nrow(s), rng, ov, sum(s$n),
                if (nrow(s) == 2) sprintf(", while the means differ by %s", fmt(abs(diff(s$mean)), 1)) else ""),
        dup)
  })

  # Same n, mean and SD, two clusters
  cg_alt <- reactive({
    input$cg_reshuffle
    d <- cg_d(); l <- levels(d$gn)[1]; x <- d$val[d$gn == l]; n <- length(x)
    validate(need(n >= 4, "The first group needs at least four patients."))
    z <- c(rep(-1, n %/% 2), rep(1, n - n %/% 2)) + rnorm(n, 0, 0.18)
    z <- (z - mean(z)) / sd(z)
    y <- round(mean(x) + sd(x) * z, 1)
    list(orig = x, alt = y, lab = sub("\\s*\\(n = .*", "", l), n = n)
  })
  cg_same <- reactive({
    a <- cg_alt()
    data.frame(val = c(a$orig, a$alt),
               gn = factor(rep(c(paste0(a$lab, "\n(actual)"), paste0(a$lab, "\n(look-alike)")), each = a$n),
                           levels = c(paste0(a$lab, "\n(actual)"), paste0(a$lab, "\n(look-alike)"))))
  })
  plot_out("cg_same_bar", {
    d <- cg_same()
    s <- do.call(rbind, lapply(levels(d$gn), function(l) { x <- d$val[d$gn == l]
      data.frame(gn = l, n = length(x), mean = mean(x), sd = sd(x), se = sd(x) / sqrt(length(x))) }))
    dynamite(s, input$cg_err, paste("Mean", vlab(input$cg_y)), fs()) + labs(subtitle = "The same bars")
  }, res = 96)
  plot_out("cg_same_pts", {
    d <- cg_same(); set.seed(3)
    ggplot(d, aes(gn, val, colour = gn)) + geom_jitter(width = 0.12, height = 0, size = 2.8, alpha = 0.8) +
      stat_summary(fun = mean, geom = "point", shape = 18, size = 6, colour = "#33414D") +
      scale_colour_manual(values = c(pal[1], pal[5]), guide = "none") +
      labs(x = NULL, y = vlab(input$cg_y), subtitle = "Different patients (diamond = mean)") + theme_class(fs())
  }, res = 96)
  output$cg_same_tbl <- renderTable({
    d <- cg_same()
    do.call(rbind, lapply(levels(d$gn), function(l) { x <- d$val[d$gn == l]
      data.frame(Group = gsub("\n", " ", l), n = length(x), Mean = fmt(mean(x), 2), SD = fmt(sd(x), 2),
                 SE = fmt(sd(x) / sqrt(length(x)), 2), Range = paste(fmtn(min(x)), "to", fmtn(max(x))), check.names = FALSE) }))
  }, striped = TRUE, align = "l")

  # ---- Scatter plot -----------------------------------------------------------------------
  sc_d <- reactive({
    df <- dat(); x <- req(input$sc_x); y <- req(input$sc_y)
    req(x %in% names(df), y %in% names(df))
    validate(need(x != y, "Choose two different variables."))
    keep <- !is.na(df[[x]]) & !is.na(df[[y]])
    d <- data.frame(x = df[[x]][keep], y = df[[y]][keep])
    cg <- input$sc_col
    if (!is.null(cg) && cg != "" && cg %in% names(df)) d$col <- with_n(as_cat(df, cg)[keep])
    d
  })
  sc_obj <- reactive({
    d <- sc_d(); lg <- isTRUE(input$sc_log)
    if (lg) validate(need(all(d$x > 0) && all(d$y > 0), "A log scale needs every value above zero."))
    hascol <- "col" %in% names(d)
    p <- ggplot(d, aes(x, y)) + theme_class(fs()) +
      labs(x = vlab(input$sc_x), y = vlab(input$sc_y), colour = if (hascol) vlab(input$sc_col))
    p <- p + if (hascol) geom_point(aes(colour = col), size = 3.2, alpha = 0.85) else geom_point(size = 3.2, alpha = 0.85, colour = pal[3])
    if (hascol) p <- p + scale_colour_manual(values = rep_len(pal, nlevels(d$col)))
    if (isTRUE(input$sc_lm)) p <- p + geom_smooth(method = "lm", formula = y ~ x, colour = pal[2], fill = "#F6DCD3", linewidth = 1.1)
    if (lg) p <- p + scale_x_log10() + scale_y_log10()
    p + labs(subtitle = sprintf("n = %d; r = %.2f%s", nrow(d), cor(d$x, d$y),
                                if (lg) sprintf(" (raw scale), %.2f (log scale)", cor(log(d$x), log(d$y))) else ""))
  })
  plot_out("sc_plot", sc_obj(), res = 96)
  output$sc_plot_code <- renderText({
    cg <- input$sc_col; hascol <- !is.null(cg) && cg != ""
    body <- c(sprintf("ggplot(hyp, aes(x = %s, y = %s%s))", input$sc_x, input$sc_y, if (hascol) paste0(", colour = ", cg) else ""),
              "geom_point(size = 3, alpha = 0.85)",
              if (isTRUE(input$sc_lm)) "geom_smooth(method = \"lm\")",
              if (isTRUE(input$sc_log)) c("scale_x_log10()", "scale_y_log10()"),
              sprintf("labs(x = %s, y = %s)", qq(vlab(input$sc_x)), qq(vlab(input$sc_y))))
    paste0(code_head, if (hascol) lev_code(dat(), cg) else "", "\n", paste(body, collapse = " +\n  "),
           sprintf("\n\ncor(hyp$%s, hyp$%s)%s", input$sc_x, input$sc_y,
                   if (isTRUE(input$sc_log)) sprintf("\ncor(log(hyp$%s), log(hyp$%s))", input$sc_x, input$sc_y) else ""))
  })
  output$sc_read <- renderUI({
    d <- sc_d(); r <- cor(d$x, d$y)
    extra <- if (isTRUE(input$sc_log) && all(d$x > 0) && all(d$y > 0))
      sprintf(" On the log scale r = %.2f: the number changes because r measures straight-line association on whatever scale it is given.", cor(log(d$x), log(d$y)))
    div(class = "ok-note", sprintf("Pearson r = %.2f for %d patients. %s", r, nrow(d),
                                   if (abs(r) >= 0.5) "A clear straight-line association; it does not by itself say why." else "A weak straight-line association."),
        extra)
  })

  # ---- Overplotting -------------------------------------------------------------------------
  op_d <- reactive({
    df <- dat(); x <- req(input$op_x); y <- req(input$op_y)
    req(x %in% names(df), y %in% names(df))
    validate(need(x != y, "Choose two different variables."),
             need(is.numeric(df[[x]]) && is.numeric(df[[y]]), "Choose two numeric variables."))
    keep <- !is.na(df[[x]]) & !is.na(df[[y]])
    data.frame(x = df[[x]][keep], y = df[[y]][keep])
  })
  plot_out("op_plot", {
    d <- op_d(); m <- input$op_mode
    p <- ggplot(d, aes(x, y)) + theme_class(fs()) + labs(x = vlab(input$op_x), y = vlab(input$op_y))
    if (m == "plain") p <- p + geom_point(size = 4, colour = pal[3])
    else if (m == "jitter") { set.seed(5); p <- p + geom_jitter(width = input$op_jit, height = input$op_jit * 0.6, size = 3.6, alpha = 0.5, colour = pal[3]) }
    else p <- p + geom_count(colour = pal[3], alpha = 0.8) + scale_size_area(max_size = 12, name = "Patients")
    if (all(d$x == round(d$x))) p <- p + scale_x_continuous(breaks = integer_breaks)
    if (all(d$y == round(d$y))) p <- p + scale_y_continuous(breaks = integer_breaks)
    p
  }, res = 96)
  output$op_plot_code <- renderText({
    geom <- switch(input$op_mode, plain = "geom_point()",
                   jitter = sprintf("geom_jitter(width = %s, height = %s, alpha = 0.5)", input$op_jit, input$op_jit * 0.6),
                   count = "geom_count()")
    paste0(code_head, "\n", sprintf("ggplot(hyp, aes(x = %s, y = %s)) +\n  %s +\n  labs(x = %s, y = %s)",
                                     input$op_x, input$op_y, geom, qq(vlab(input$op_x)), qq(vlab(input$op_y))))
  })
  output$op_read <- renderUI({
    d <- op_d(); k <- table(paste(d$x, d$y)); hid <- nrow(d) - length(k)
    top <- names(k)[which.max(k)]
    div(class = if (input$op_mode == "plain" && hid > 0) "warn-note" else "ok-note",
        sprintf("%d patients but only %d distinct positions: %d point%s lie exactly on top of another. The busiest spot (%s = %s, %s = %s) holds %d patients.",
                nrow(d), length(k), hid, if (hid == 1) "" else "s",
                vlab(input$op_x), strsplit(top, " ")[[1]][1], vlab(input$op_y), strsplit(top, " ")[[1]][2], max(k)))
  })

  # ---- When may an axis not start at zero? -----------------------------------------------
  ax_d <- reactive({
    df <- dat(); g <- req(input$ax_g); req(g %in% names(df)); gf <- as_cat(df, g)
    if (input$ax_type == "line") {
      validate(need(all(c("SBP_Base", "SBP_6mo") %in% names(df)), "The line chart needs SBP_Base and SBP_6mo."))
      s <- do.call(rbind, lapply(levels(gf), function(l) {
        k <- gf == l & !is.na(gf)
        data.frame(g = factor(l, levels = levels(gf)), time = factor(c("Baseline", "6 months"), levels = c("Baseline", "6 months")),
                   m = c(mean(df$SBP_Base[k], na.rm = TRUE), mean(df$SBP_6mo[k], na.rm = TRUE)), n = sum(k))
      }))
      return(list(type = "line", s = s, lo = min(s$m), ylab = "Mean systolic BP (mmHg)", g = g))
    }
    v <- req(input$ax_var); req(v %in% names(df)); validate(need(is.numeric(df[[v]]), "Choose a numeric outcome."))
    keep <- !is.na(df[[v]]) & !is.na(gf)
    d <- data.frame(val = df[[v]][keep], g = droplevels(gf[keep]))
    s <- do.call(rbind, lapply(levels(d$g), function(l) { x <- d$val[d$g == l]; n <- length(x)
      h <- if (n > 1) qt(0.975, n - 1) * sd(x) / sqrt(n) else 0
      data.frame(g = l, m = mean(x), lcl = mean(x) - h, ucl = mean(x) + h, n = n) }))
    list(type = input$ax_type, d = d, s = s, lo = if (input$ax_type == "box") min(d$val) else min(s$m),
         ylab = paste(if (input$ax_type == "box") "" else "Mean", vlab(v)), g = g, v = v)
  })
  observeEvent(list(input$ax_type, input$ax_var, input$ax_g, dat()), {
    a <- ax_d()
    mx <- if (a$type == "box") floor(quantile(a$d$val, 0.5)) else floor(a$lo * 0.98)
    val <- if (a$type == "bar") 0 else floor(a$lo * 0.9)
    updateSliderInput(session, "ax_start", min = 0, max = max(1, mx), value = min(val, mx))
  })
  ax_obj <- reactive({
    a <- ax_d(); st <- input$ax_start; req(!is.null(st))
    gl <- if (a$type == "line") levels(droplevels(a$s$g)) else levels(a$d$g)
    cols <- setNames(rep_len(pal, length(gl)), gl)
    p <- switch(a$type,
      bar  = ggplot(a$s, aes(factor(g, levels = gl), m, fill = g)) + geom_col(width = 0.55) +
               geom_text(aes(label = fmt(m, 1)), vjust = -0.4, size = fs() / 3) + scale_fill_manual(values = cols, guide = "none"),
      dot  = ggplot(a$s, aes(factor(g, levels = gl), m, colour = g)) +
               geom_errorbar(aes(ymin = lcl, ymax = ucl), width = 0.15, linewidth = 1.1) + geom_point(size = 4.5) +
               scale_colour_manual(values = cols, guide = "none"),
      line = ggplot(a$s, aes(time, m, colour = g, group = g)) + geom_line(linewidth = 1.3) + geom_point(size = 4) +
               scale_colour_manual(values = cols, name = vlab(a$g)),
      box  = ggplot(a$d, aes(g, val, fill = g)) + geom_boxplot(width = 0.5, alpha = 0.45, outlier.colour = col_mean) +
               scale_fill_manual(values = cols, guide = "none"))
    top <- if (a$type == "box") max(a$d$val) else if (a$type == "dot") max(a$s$ucl) else max(a$s$m)
    if (a$type == "bar") top <- top + 0.1 * (top - st)
    p <- p + scale_y_continuous(expand = expansion(mult = c(0, 0.05))) +
      coord_cartesian(ylim = c(st, top)) +
      labs(x = if (a$type == "line") NULL else vlab(a$g), y = a$ylab) + theme_class(fs())
    if (st > 0) {
      if (isTRUE(input$ax_declare)) {
        p <- p + labs(caption = sprintf("Axis starts at %s, not zero", format(st))) +
          annotate("text", x = -Inf, y = st, label = "\u2248", size = fs() / 1.5, colour = col_mean, hjust = -0.1, vjust = -0.1) +
          theme(plot.caption = element_text(colour = col_mean, face = "bold", hjust = 0, size = rel(0.85)))
      }
      if (a$type == "bar") p <- p + theme(axis.line.y = element_line(colour = col_mean, linewidth = 1.2))
    }
    p
  })
  plot_out("ax_plot", ax_obj(), res = 96)
  output$ax_plot_code <- renderText({
    a <- ax_d(); st <- input$ax_start
    lim <- sprintf("coord_cartesian(ylim = c(%s, NA))", st)
    body <- switch(a$type,
      bar = c(sprintf("ggplot(summ, aes(x = %s, y = mean))", a$g), "geom_col(width = 0.55)"),
      dot = c(sprintf("ggplot(summ, aes(x = %s, y = mean))", a$g), "geom_errorbar(aes(ymin = lcl, ymax = ucl), width = 0.15)", "geom_point(size = 4)"),
      line = c(sprintf("ggplot(summ, aes(x = time, y = mean, colour = %s, group = %s))", a$g, a$g), "geom_line()", "geom_point(size = 4)"),
      box = c(sprintf("ggplot(hyp, aes(x = %s, y = %s))", a$g, a$v), "geom_boxplot(width = 0.5)"))
    if (st > 0) body <- c(body, lim, if (isTRUE(input$ax_declare)) sprintf("labs(caption = \"Axis starts at %s, not zero\")", st))
    pre <- switch(a$type,
      bar = sprintf("summ <- aggregate(%s ~ %s, data = hyp, FUN = mean)\nnames(summ)[2] <- \"mean\"\n", a$v, a$g),
      dot = sprintf("summ <- do.call(rbind, lapply(split(hyp$%s, hyp$%s), function(x) {\n  h <- qt(0.975, length(x) - 1) * sd(x) / sqrt(length(x))\n  data.frame(mean = mean(x), lcl = mean(x) - h, ucl = mean(x) + h)\n}))\nsumm$%s <- rownames(summ)\n", a$v, a$g, a$g),
      line = sprintf("long <- data.frame(%s = rep(hyp$%s, 2),\n                   time = factor(rep(c(\"Baseline\", \"6 months\"), each = nrow(hyp)), levels = c(\"Baseline\", \"6 months\")),\n                   sbp = c(hyp$SBP_Base, hyp$SBP_6mo))\nsumm <- aggregate(sbp ~ %s + time, data = long, FUN = mean)\nnames(summ)[3] <- \"mean\"\n", a$g, a$g, a$g),
      box = "")
    paste0(code_head, "\n", pre, "\n", paste(body, collapse = " +\n  "))
  })
  output$ax_verdict <- renderUI({
    a <- ax_d(); st <- input$ax_start; req(!is.null(st))
    if (st == 0) return(div(class = "ok-note", "The axis starts at zero: lengths and positions are both read honestly."))
    if (a$type == "bar") {
      hi <- max(a$s$m); lo <- min(a$s$m)
      if (lo > st && lo > 0) {
        real <- hi / lo; look <- (hi - st) / (lo - st)
        return(div(class = "warn-note", strong("Never for bars. "),
                   sprintf("The bar is read as a length. In the data the taller bar is %.2f times the shorter; on this axis it looks %.2f times as tall, a %.1f-fold exaggeration%s.",
                           real, look, look / real, if (isTRUE(input$ax_declare)) ", and declaring it does not repair the encoding" else "")))
      }
      return(div(class = "warn-note", strong("Never for bars. "), "The axis now cuts into the bars themselves."))
    }
    if (a$type == "box") {
      cut <- sum(a$d$val < st)
      if (cut > 0) return(div(class = "warn-note", strong("Judgement, and this fails it. "),
                              sprintf("%d patient%s below %s %s off the plot.", cut, if (cut > 1) "s" else "", format(st), if (cut > 1) "are" else "is")))
      return(div(class = if (isTRUE(input$ax_declare)) "ok-note" else "warn-note", strong("Acceptable with judgement. "),
                 "The range shown covers every patient", if (isTRUE(input$ax_declare)) " and the axis is declared." else ", but the truncation should be declared."))
    }
    div(class = if (isTRUE(input$ax_declare)) "ok-note" else "warn-note",
        strong(if (a$type == "line") "Usually acceptable for a line over time. " else "Usually acceptable for points. "),
        "Nobody reads a point or a line as a length, and zero mmHg is nowhere near the data. ",
        if (isTRUE(input$ax_declare)) "The truncation is declared on the figure." else "Declare the truncation: concealing it is what turns carelessness into deception.")
  })
  output$ax_rules <- renderUI({
    t <- input$ax_type
    r <- function(k, head, txt) div(class = paste("rule", if (k == t) "rule-on" else ""), tags$b(head), tags$br(), txt)
    tagList(
      r("bar", "Bar charts: never", "The bar encodes a quantity by its length. Cutting the axis breaks the encoding, and the reader cannot see it."),
      r("line", "Line charts over time: usually acceptable", "A time series of blood pressure need not start at zero; nobody reads the line as a length."),
      r("dot", "Points and scatter plots: usually acceptable", "Points encode position, not magnitude. Zero is often nowhere near the data."),
      r("box", "Boxplots: judgement", "Acceptable if the range shown covers all the data and the axis is clearly labelled."),
      div(class = "rule", tags$b("Whenever you truncate, say so"), tags$br(), "A visible break in the axis, or a note on the figure."))
  })

  # ---- Colour and accessibility ------------------------------------------------------------
  ac_d <- reactive({
    df <- dat(); x <- req(input$ac_x); g <- req(input$ac_g)
    req(x %in% names(df), g %in% names(df))
    validate(need(x != g, "Choose two different variables."))
    tb <- as.data.frame(table(x = as_cat(df, x), g = as_cat(df, g)))
    tb$pct <- 100 * tb$Freq / ave(tb$Freq, tb$g, FUN = sum)
    tb
  })
  ac_views <- reactive({ v <- input$ac_views; validate(need(length(v) > 0, "Tick at least one way of seeing.")); v })
  plot_out("ac_plot", {
    tb <- ac_d(); views <- ac_views(); gl <- levels(tb$g)
    base <- setNames(acc_pal(input$ac_pal, length(gl)), gl)
    dd <- do.call(rbind, lapply(views, function(v) { t2 <- tb; t2$view <- v
      t2$fill <- simulate_cvd(base[as.character(t2$g)], v); t2 }))
    dd$view <- factor(dd$view, levels = views)
    dodge <- position_dodge(width = 0.82)
    p <- ggplot(dd, aes(x, pct, group = g)) +
      geom_col(aes(fill = fill), position = dodge, width = 0.78, colour = "white", linewidth = 0.4) +
      scale_fill_identity() +
      scale_y_continuous(expand = expansion(mult = c(0, 0.18))) +
      facet_wrap(~view, ncol = 1) +
      labs(x = vlab(input$ac_x), y = paste("% within each", tolower(vlab(input$ac_g)))) +
      theme_class(fs()) + theme(strip.text = element_text(face = "bold", size = rel(0.95), hjust = 0))
    if (isTRUE(input$ac_lab)) {
      ab <- function(s) ifelse(nchar(s) > 6, paste0(substr(s, 1, 5), "."), s)
      p <- p + geom_text(aes(label = ab(as.character(g)), y = pct), position = dodge, vjust = -0.35,
                         size = fs() / 4.2, colour = "#33414D", fontface = "bold")
    } else {
      key <- do.call(rbind, lapply(views, function(v) data.frame(view = v, g = gl, fill = simulate_cvd(base, v), k = seq_along(gl))))
      key$view <- factor(key$view, levels = views)
      ymax <- max(tb$pct) * 1.15; nx <- nlevels(tb$x)
      key$xpos <- nx + 0.75; key$ypos <- ymax * (1 - 0.16 * (key$k - 1))
      p <- p + geom_point(data = key, aes(x = xpos - 0.12, y = ypos, colour = fill), inherit.aes = FALSE, shape = 15, size = 5) +
        geom_text(data = key, aes(x = xpos, y = ypos, label = g), inherit.aes = FALSE, hjust = 0, size = fs() / 4.2, colour = "#33414D") +
        scale_colour_identity() + coord_cartesian(xlim = c(0.5, nx + 2.1))
    }
    p
  }, res = 96, height = function() 140 + 230 * max(1, length(input$ac_views)))
  output$ac_plot_code <- renderText({
    paste0(code_head, lev_code(dat(), input$ac_x), lev_code(dat(), input$ac_g), "\n",
           sprintf("tab <- as.data.frame(prop.table(table(%s = hyp$%s, %s = hyp$%s), 2) * 100)\n\n",
                   input$ac_x, input$ac_x, input$ac_g, input$ac_g),
           sprintf("ggplot(tab, aes(x = %s, y = Freq, fill = %s)) +\n  geom_col(position = position_dodge(0.82), width = 0.78) +\n", input$ac_x, input$ac_g),
           sprintf("  scale_fill_manual(values = %s) +\n", acc_pal_code[[input$ac_pal]]),
           if (isTRUE(input$ac_lab)) sprintf("  geom_text(aes(label = %s), position = position_dodge(0.82), vjust = -0.35) +\n  guides(fill = \"none\") +\n", input$ac_g) else "",
           sprintf("  labs(x = %s, y = %s)\n\n# Check the figure in greyscale before it leaves your desk:\n# colorspace::desaturate() or the colorblindr package simulate it.",
                   qq(vlab(input$ac_x)), qq(paste("% within each", tolower(vlab(input$ac_g))))))
  })
  output$ac_read <- renderUI({
    gl <- levels(ac_d()$g); cols <- acc_pal(input$ac_pal, length(gl))
    L <- lightness(cols); dl <- if (length(L) > 1) min(diff(sort(L))) else 100
    # a pair stays distinguishable if it differs clearly in lightness or strongly in colour (CIE Lab)
    pair_ok <- function(h) {
      if (length(h) < 2) return(TRUE)
      lab <- grDevices::convertColor(t(grDevices::col2rgb(h)) / 255, from = "sRGB", to = "Lab")
      ij <- utils::combn(length(h), 2)
      all(apply(ij, 2, function(k) abs(lab[k[1], 1] - lab[k[2], 1]) >= 15 || sqrt(sum((lab[k[1], ] - lab[k[2], ])^2)) >= 40))
    }
    ok_grey <- dl >= 20
    ok_cvd <- all(vapply(c("Deuteranopia", "Protanopia"), function(v) pair_ok(simulate_cvd(cols, v)), logical(1)))
    lab_ok <- isTRUE(input$ac_lab)
    div(class = if ((ok_grey && ok_cvd) || lab_ok) "ok-note" else "warn-note",
        tags$b("Greyscale: "), sprintf("the closest two colours differ by %.0f lightness units (0 = black, 100 = white); ", dl),
        if (ok_grey) "they stay apart." else "they merge into one grey.",
        tags$br(), tags$b("Red\u2013green deficiency: "),
        if (ok_cvd) "the colours stay distinguishable." else "two or more colours become hard to tell apart.",
        tags$br(), tags$b("Labels: "),
        if (lab_ok) "direct labels carry the message whatever the colours look like." else "only the colour matching to the key tells the groups apart.")
  })

  # ---- Small multiples ------------------------------------------------------------------------
  fm_d <- reactive({
    df <- dat_plus(); y <- req(input$fm_y); g <- req(input$fm_g); f <- req(input$fm_f)
    req(all(c(y, g, f) %in% names(df)))
    validate(need(is.numeric(df[[y]]), "Choose a numeric outcome."),
             need(g != f, "Choose a panel variable different from the groups."))
    gf <- as_cat(df, g); ff <- as_cat(df, f)
    keep <- !is.na(df[[y]]) & !is.na(gf) & !is.na(ff)
    d <- data.frame(val = df[[y]][keep], g = droplevels(gf[keep]), f = droplevels(ff[keep]))
    nf <- table(d$f)
    d$panel <- if (isTRUE(input$fm_n)) {
      ng <- tapply(d$g, d$f, function(z) paste(as.integer(table(z)), collapse = " + "))
      factor(as.character(d$f), levels = names(nf),
             labels = sprintf("%s\nn = %d (%s)", names(nf), as.integer(nf), ng[names(nf)]))
    } else d$f
    d
  })
  fm_obj <- reactive({
    d <- fm_d(); set.seed(9)
    p <- ggplot(d, aes(g, val, colour = g)) +
      geom_jitter(width = 0.15, height = 0, size = 2.6, alpha = 0.75) +
      scale_colour_manual(values = rep_len(pal, nlevels(d$g)), guide = "none") +
      facet_wrap(~panel, nrow = 1, scales = input$fm_scales) +
      scale_x_discrete(labels = function(z) gsub(" ", "\n", z)) +
      labs(x = NULL, y = vlab(input$fm_y),
           subtitle = if (isTRUE(input$fm_n)) NULL else "No n given: every panel looks equally weighty") +
      theme_class(fs()) +
      theme(strip.text = element_text(face = "bold", size = rel(0.8)), panel.border = element_rect(colour = "#E1E7EC", fill = NA),
            axis.text.x = element_text(size = rel(0.85), lineheight = 0.9))
    if (isTRUE(input$fm_med)) {
      s <- aggregate(val ~ g + panel, data = d, FUN = median)
      p <- p + geom_errorbar(data = s, aes(x = g, ymin = val, ymax = val), inherit.aes = FALSE, width = 0.5, linewidth = 1.2, colour = "#33414D")
    }
    p
  })
  plot_out("fm_plot", fm_obj(), res = 96)
  output$fm_plot_code <- renderText({
    f <- input$fm_f
    pre <- if (f == "AgeGroup") "hyp$AgeGroup <- cut(hyp$Age, c(-Inf, 49, 59, Inf), labels = c(\"<50\", \"50-59\", \"60+\"))\n" else lev_code(dat(), f)
    paste0(code_head, pre, lev_code(dat(), input$fm_g), "\n",
           sprintf("ggplot(hyp, aes(x = %s, y = %s)) +\n  geom_jitter(width = 0.15, height = 0, alpha = 0.75) +\n", input$fm_g, input$fm_y),
           if (isTRUE(input$fm_med)) "  stat_summary(fun = median, fun.min = median, fun.max = median, geom = \"crossbar\", width = 0.5) +\n" else "",
           sprintf("  facet_wrap(~ %s, nrow = 1%s) +\n", f, if (input$fm_scales == "fixed") "" else ", scales = \"free_y\""),
           sprintf("  labs(x = NULL, y = %s)\n\n# n in every panel:\ntable(hyp$%s, hyp$%s)", qq(vlab(input$fm_y)), f, input$fm_g))
  })
  output$fm_tbl <- renderTable({
    d <- fm_d(); gl <- levels(d$g)
    do.call(rbind, lapply(levels(d$f), function(l) {
      s <- d[d$f == l, ]
      ns <- vapply(gl, function(k) sum(s$g == k), integer(1))
      ms <- vapply(gl, function(k) if (any(s$g == k)) fmt(mean(s$val[s$g == k]), 1) else "\u2014", "")
      out <- data.frame(Panel = l, "n (total)" = nrow(s), check.names = FALSE)
      for (k in seq_along(gl)) { out[[paste("n,", gl[k])]] <- ns[k]; out[[paste("Mean,", gl[k])]] <- ms[k] }
      out
    }))
  }, striped = TRUE, align = "l")
  output$fm_read <- renderUI({
    d <- fm_d(); nf <- table(d$f); small <- names(nf)[nf < input$fm_min]
    out <- list()
    if (length(small))
      out <- c(out, list(div(class = "warn-note",
        sprintf("%s hold%s fewer than %d patients (%s). Read the trend across panels, not %s.",
                paste(sprintf("\u201c%s\u201d", small), collapse = " and "), if (length(small) == 1) "s" else "",
                input$fm_min, paste(sprintf("%s: %d", small, as.integer(nf[small])), collapse = ", "),
                if (length(small) == 1) "this panel" else "these panels"))))
    if (input$fm_scales != "fixed")
      out <- c(out, list(div(class = "warn-note", "Free axes make every panel look important and make them impossible to compare. Shared axes are the right default.")))
    if (!length(out)) out <- list(div(class = "ok-note", sprintf("Every panel holds at least %d patients and the axes are shared.", input$fm_min)))
    tagList(out)
  })

  # ---- Table, figure or neither ------------------------------------------------------------------
  output$tf_guide <- renderUI({
    tags$table(class = "table guide",
      tags$thead(tags$tr(tags$th("What you are reporting"), tags$th("Vehicle"))),
      tags$tbody(
        tags$tr(tags$td("One or two numbers, e.g. the difference in BP control between arms"), tags$td(tags$b("A sentence"))),
        tags$tr(tags$td("Exact values readers may re-use, e.g. baseline characteristics by arm"), tags$td(tags$b("A table"))),
        tags$tr(tags$td("A shape, a relationship, or a comparison across many groups"), tags$td(tags$b("A figure"))),
        tags$tr(tags$td("Numbers already given in a table"), tags$td(tags$b("Neither: do not duplicate")))))
  })
  output$t1_tbl <- renderTable({
    df <- dat(); g <- req(input$t1_g); req(g %in% names(df)); gf <- as_cat(df, g)
    vars <- setdiff(vars_of(df, c("nominal", "ordinal", "ratio", "interval")), c(g, followup_vars))
    validate(need(length(vars) > 0, "No baseline variables to show."))
    lv <- levels(gf); nl <- as.integer(table(gf))
    rows <- list()
    for (v in vars) {
      if (scale_of(df, v) %in% c("ratio", "interval") && is.numeric(df[[v]])) {
        dp <- min(2, dec_places(df[[v]]) + 1)
        cells <- vapply(lv, function(l) { x <- df[[v]][gf == l & !is.na(gf)]; x <- x[!is.na(x)]
          sprintf("%s (%s)", fmt(mean(x), dp), fmt(sd(x), dp)) }, "")
        rows[[length(rows) + 1]] <- c(paste0(vlab(v), ", mean (SD)"), cells)
      } else {
        f <- as_cat(df, v)
        rows[[length(rows) + 1]] <- c(paste0(vlab(v), ", n (%)"), rep("", length(lv)))
        for (k in levels(f)) {
          cells <- vapply(seq_along(lv), function(j) { s <- gf == lv[j] & !is.na(gf) & !is.na(f)
            sprintf("%d (%.0f%%)", sum(f[s] == k), 100 * mean(f[s] == k)) }, "")
          rows[[length(rows) + 1]] <- c(paste0("\u2003", k), cells)
        }
      }
    }
    out <- as.data.frame(do.call(rbind, rows), stringsAsFactors = FALSE)
    names(out) <- c("Characteristic", sprintf("%s (n = %d)", lv, nl))
    out
  }, striped = TRUE, align = "l")
  output$t1_note <- renderUI({
    df <- dat(); vars <- setdiff(vars_of(df, c("nominal", "ordinal", "ratio", "interval")), c(input$t1_g, followup_vars))
    div(class = "ok-note", sprintf("%d baseline characteristics. The exact values matter and may be re-used, so this belongs in a table, not a figure. Follow-up measures (%s) are left out.",
                                   length(vars), paste(intersect(followup_vars, names(df)), collapse = ", ")))
  })

  # ---- Spot the faults ----------------------------------------------------------------------------
  ck <- reactiveVal(NULL)
  new_ck <- function() {
    df <- dat()
    can_cat <- all(c("Adherence", "Arm") %in% names(df))
    can_con <- all(c("SBP_6mo", "Arm") %in% names(df))
    validate(need(can_cat || can_con, "This exercise uses the class dataset (Adherence, SBP_6mo, Arm)."))
    type <- if (can_cat && can_con) sample(c("cat", "con"), 1) else if (can_cat) "cat" else "con"
    pool <- if (type == "cat") c("scale", "labels", "zero", "n", "grey", "caption", "one")
            else c("data", "labels", "n", "grey", "caption", "one")
    faults <- sample(pool, sample(2:3, 1))
    if (type == "con" && "data" %in% faults && runif(1) < 0.5) faults <- unique(c(faults, "zero"))
    list(type = type, faults = faults, seed = sample.int(1e5, 1))
  }
  observeEvent(input$ck_new, { ck(new_ck()); updateCheckboxGroupInput(session, "ck_marks", selected = character(0)) })
  observe({ if (is.null(ck())) ck(new_ck()) })
  ck_revealed <- reactiveVal(FALSE)
  observeEvent(input$ck_new, ck_revealed(FALSE))
  observeEvent(input$ck_reveal, ck_revealed(TRUE))
  output$ck_plot <- renderPlot({
    validate(need(all(c("Adherence", "SBP_6mo", "Arm") %in% names(dat())),
                  "This exercise uses the class dataset: press \u201cBack to the class dataset\u201d in the sidebar."))
    k <- req(ck()); df <- dat(); F <- k$faults; has <- function(x) x %in% F
    set.seed(k$seed)
    cols <- if (has("grey")) c("#D7301F", "#1A9850") else c("#E69F00", "#0072B2")
    arm <- as_cat(df, "Arm"); armn <- if (has("n")) arm else with_n(arm)
    if (k$type == "cat") {
      ad <- as_cat(df, "Adherence")
      tb <- as.data.frame(table(x = ad, g = armn)); tb$pct <- 100 * tb$Freq / ave(tb$Freq, tb$g, FUN = sum)
      if (has("scale")) { tot <- tapply(tb$Freq, tb$x, sum); tb$x <- factor(tb$x, levels = names(sort(tot, decreasing = TRUE))) }
      p <- ggplot(tb, aes(x, pct, fill = g)) + geom_col(position = position_dodge(0.8), width = 0.75) +
        scale_fill_manual(values = cols, name = NULL) + theme_class(fs())
      p <- p + if (has("labels")) labs(x = "Adherence", y = "pct") else labs(x = "Medication adherence (1 = never, 5 = always)", y = "% of patients within each arm")
      if (has("zero")) p <- p + coord_cartesian(ylim = c(8, max(tb$pct) * 1.05), expand = FALSE)
      else p <- p + scale_y_continuous(expand = expansion(mult = c(0, 0.08)))
      if (!has("grey")) p <- p + geom_text(aes(label = sub(" \\(n.*", "", g)), position = position_dodge(0.8), vjust = -0.3, size = fs() / 4.5) +
        guides(fill = "none")
    } else {
      d <- data.frame(val = df$SBP_6mo, g = armn); d <- d[!is.na(d$val) & !is.na(d$g), ]
      if (has("data")) {
        s <- aggregate(val ~ g, data = d, FUN = function(x) c(m = mean(x), se = sd(x) / sqrt(length(x))))
        s <- data.frame(g = s$g, m = s$val[, "m"], se = s$val[, "se"])
        p <- ggplot(s, aes(g, m, fill = g)) + geom_col(width = 0.55) +
          geom_errorbar(aes(ymin = m, ymax = m + se), width = 0.15, linewidth = 0.9) + scale_fill_manual(values = cols, name = NULL)
        p <- p + if (has("zero")) coord_cartesian(ylim = c(130, 160), expand = FALSE) else scale_y_continuous(expand = expansion(mult = c(0, 0.06)))
      } else {
        p <- ggplot(d, aes(g, val, colour = g)) + geom_jitter(width = 0.15, height = 0, size = 2.8, alpha = 0.8) +
          stat_summary(fun = median, fun.min = median, fun.max = median, geom = "crossbar", width = 0.45, colour = "#33414D") +
          scale_colour_manual(values = cols, name = NULL)
      }
      p <- p + theme_class(fs()) + (if (has("labels")) labs(x = "Arm", y = "SBP_6mo") else labs(x = NULL, y = "Systolic BP at 6 months (mmHg)"))
      if (has("grey")) p <- p + theme(axis.text.x = element_blank(), legend.position = "right")
      else p <- p + guides(fill = "none", colour = "none")
    }
    p
  }, res = 96)
  output$ck_caption <- renderUI({
    k <- req(ck()); F <- k$faults; nmsg <- if ("n" %in% F) "" else " Usual care n = 20; intervention n = 20."
    good <- if (k$type == "cat")
      paste0("Figure 2. Medication adherence score (1 = never to 5 = always) by treatment arm, as the percentage of patients within each arm.", nmsg)
    else if ("data" %in% F)
      paste0("Figure 2. Systolic blood pressure at 6 months (mmHg) by treatment arm: bar = mean, whisker = one standard error.", nmsg)
    else paste0("Figure 2. Systolic blood pressure at 6 months (mmHg) by treatment arm. Each point is one patient; the horizontal line is the group median.", nmsg)
    cap <- if ("caption" %in% F) (if (k$type == "cat") "Figure 2. Adherence by arm." else "Figure 2. SBP by arm.") else good
    if ("one" %in% F) cap <- paste(cap, "The same values are given in Table 2.")
    div(class = "caption-box", cap)
  })
  output$ck_result <- renderUI({
    req(ck_revealed()); k <- ck(); F <- k$faults; M <- input$ck_marks; if (is.null(M)) M <- character(0)
    hit <- intersect(F, M); miss <- setdiff(F, M); fa <- setdiff(M, F)
    li <- function(id, cls, tag) tags$li(class = cls, tags$b(tag), " ", ck_items[[id]])
    div(class = if (!length(miss) && !length(fa)) "ok-note" else "warn-note",
        tags$b(sprintf("%d of %d faults found%s.", length(hit), length(F),
                       if (length(fa)) sprintf("; %d check%s ticked that the figure passes", length(fa), if (length(fa) > 1) "s" else "") else "")),
        tags$ul(class = "ck-list",
                lapply(hit, li, cls = "ck-hit", tag = "\u2713 Found:"),
                lapply(miss, li, cls = "ck-miss", tag = "\u2717 Missed:"),
                lapply(fa, li, cls = "ck-fa", tag = "\u2013 Passes:")))
  })

  # ---- Captions ----------------------------------------------------------------------------------
  cap_plot_obj <- reactive(switch(input$cap_fig, cg = cg_pts_obj(), sc = sc_obj(), fm = fm_obj()))
  output$cap_plot <- renderPlot(cap_plot_obj(), res = 96)
  output$cap_text <- renderUI({
    f <- input$cap_fig
    if (f == "cg") {
      s <- cg_sum(); grp <- vlab(input$cg_g)
      bad <- sprintf("Figure 3. %s by %s.", input$cg_y, tolower(input$cg_g))
      good <- sprintf("Figure 3. %s by %s. Each point is one patient; the horizontal line is the group median. %s.",
                      vlab(input$cg_y), tolower(grp), paste(sub("\\s*\\(n = (\\d+)\\)", " n = \\1", s$gn), collapse = "; "))
    } else if (f == "sc") {
      d <- sc_d()
      bad <- sprintf("Figure 3. %s vs %s.", input$sc_y, input$sc_x)
      good <- sprintf("Figure 3. %s against %s. Each point is one patient%s%s. n = %d.",
                      vlab(input$sc_y), tolower(vlab(input$sc_x)),
                      if (isTRUE(input$sc_lm)) "; the line is the least-squares fit with its 95% confidence band" else "",
                      if (isTRUE(input$sc_log)) "; both axes on a log scale" else "", nrow(d))
    } else {
      d <- fm_d(); nf <- table(d$f)
      bad <- sprintf("Figure 3. %s by %s and %s.", input$fm_y, tolower(input$fm_g), tolower(input$fm_f))
      good <- sprintf("Figure 3. %s by %s, one panel per %s group. Each point is one patient%s. Panel sizes: %s.",
                      vlab(input$fm_y), tolower(vlab(input$fm_g)), sub(" group$", "", tolower(sub(" \\(.*\\)$", "", vlab(input$fm_f)))),
                      if (isTRUE(input$fm_med)) "; the horizontal line is the group median" else "",
                      paste(sprintf("%s n = %d", names(nf), as.integer(nf)), collapse = ", "))
    }
    tagList(
      div(class = "warn-note", tags$b("\u2717  "), bad),
      div(class = "ok-note", tags$b("\u2713  "), good),
      tags$ol(class = "cap-rules",
              tags$li(tags$b("Say what is plotted:"), " which variable, in which units, for which patients."),
              tags$li(tags$b("Say what each mark means:"), " a point is a patient; a line is a median; a band is an interval."),
              tags$li(tags$b("Give the group sizes,"), " per group, every time."),
              tags$li(tags$b("Do not interpret in the caption:"), " the caption describes; the text argues.")))
  })

  # ---- R code behind the Session 2 displays ---------------------------------------------
  output$gt_hist_code <- renderText({
    g <- gt(); v <- input$gt_var
    paste0(code_head, "\n",
           sprintf("# classes of width %s starting at %s; boundaries sit half a unit outside the stated limits\n", format(g$w), format(g$start)),
           sprintf("brks <- seq(%s, %s, by = %s)\n\n", format(g$brks[1]), format(g$brks[length(g$brks)]), format(g$w)),
           sprintf("table(cut(hyp$%s, breaks = brks))\n\n", v),
           sprintf("ggplot(hyp, aes(x = %s)) +\n  geom_histogram(breaks = brks, colour = \"white\") +\n  labs(x = %s, y = \"Frequency (f)\")",
                   v, qq(vlab(v))))
  })
  output$bar_plot_code <- renderText({
    v <- input$bar_var; pct <- input$bar_y == "pct"
    ord <- if (input$bar_order == "freq")
      sprintf("hyp$%s <- factor(hyp$%s, levels = names(sort(table(hyp$%s), decreasing = TRUE)))\n", v, v, v) else lev_code(dat(), v)
    paste0(code_head, ord, "\n",
           sprintf("ggplot(hyp, aes(x = %s)) +\n  geom_bar(%swidth = 0.62) +\n  labs(x = %s, y = %s)", v,
                   if (pct) "aes(y = after_stat(100 * count / sum(count))), " else "", qq(vlab(v)),
                   qq(if (pct) "Percentage of patients" else "Frequency (f)")))
  })
  output$pie_bar_code <- renderText({
    v <- input$pie_var
    paste0(code_head, lev_code(dat(), v), "\n",
           sprintf("# the bar chart\nggplot(hyp, aes(x = %s)) +\n  geom_bar(width = 0.62) +\n  labs(x = %s, y = \"Frequency (f)\")\n\n", v, qq(vlab(v))),
           sprintf("# the pie, for comparison: a stacked bar in polar coordinates\nggplot(hyp, aes(x = 1, fill = %s)) +\n  geom_bar(width = 1) +\n  coord_polar(theta = \"y\") +\n  theme_void()", v))
  })
  output$cl_plot_code <- renderText({
    v <- input$cl_var; g <- input$cl_grp; pct <- input$cl_y == "pct"
    paste0(code_head, lev_code(dat(), v), lev_code(dat(), g), "\n",
           if (pct) sprintf("tab <- as.data.frame(prop.table(table(%s = hyp$%s, %s = hyp$%s), 2) * 100)\n\nggplot(tab, aes(x = %s, y = Freq, fill = %s)) +\n  geom_col(position = position_dodge(0.8), width = 0.75) +\n  labs(x = %s, y = %s, fill = %s)",
                            v, v, g, g, v, g, qq(vlab(v)), qq(paste("% within each", tolower(vlab(g)))), qq(vlab(g)))
           else sprintf("ggplot(hyp, aes(x = %s, fill = %s)) +\n  geom_bar(position = position_dodge(0.8), width = 0.75) +\n  labs(x = %s, y = \"Frequency (f)\", fill = %s)",
                        v, g, qq(vlab(v)), qq(vlab(g))))
  })
  output$hw_plot_code <- renderText({
    v <- input$hw_var; w <- input$hw_w
    paste0(code_head, "\n",
           sprintf("ggplot(hyp, aes(x = %s)) +\n  geom_histogram(binwidth = %s, colour = \"white\") +\n  labs(x = %s, y = \"Frequency (f)\")\n\n", v, format(w), qq(vlab(v))),
           "# Always try several bin widths before choosing one, and report the width you used.")
  })
  output$bx_plot_code <- renderText({
    v <- input$bx_var; g <- input$bx_grp; hasg <- !is.null(g) && g != ""
    paste0(code_head, if (hasg) lev_code(dat(), g) else "", "\n",
           sprintf("ggplot(hyp, aes(x = %s, y = %s)) +\n  geom_boxplot(width = 0.45, coef = 1.5%s) +\n",
                   v, if (hasg) g else "\"All patients\"", if (isTRUE(input$bx_pts)) ", outlier.shape = NA" else ""),
           if (isTRUE(input$bx_pts)) "  geom_jitter(width = 0, height = 0.1, alpha = 0.7) +\n" else "",
           sprintf("  labs(x = %s, y = NULL)\n\n", qq(vlab(v))),
           sprintf("# fences\nq <- quantile(hyp$%s, c(0.25, 0.75))\nq[1] - 1.5 * diff(q); q[2] + 1.5 * diff(q)", v))
  })
  output$sh_plot_code <- renderText({
    v <- input$sh_var; g <- input$sh_grp; hasg <- !is.null(g) && g != ""
    paste0(code_head, if (hasg) lev_code(dat(), g) else "", "\n",
           sprintf("ggplot(hyp, aes(x = %s)) +\n  geom_histogram(binwidth = %s, colour = \"white\") +\n", v, format(input$sh_w)),
           sprintf("  geom_vline(xintercept = mean(hyp$%s), colour = \"#E76F51\") +\n  geom_vline(xintercept = median(hyp$%s), linetype = \"longdash\", colour = \"#4C8BF5\") +\n", v, v),
           if (hasg) sprintf("  facet_wrap(~ %s, ncol = 1) +\n", g) else "",
           sprintf("  labs(x = %s, y = \"Frequency (f)\")\n\n# adjusted skewness G1\nx <- hyp$%s; n <- length(x); m <- mean(x)\ng1 <- mean((x - m)^3) / mean((x - m)^2)^1.5\ng1 * sqrt(n * (n - 1)) / (n - 2)", qq(vlab(v)), v))
  })
  output$mx_trunc_code <- renderText({
    m <- mx_vals()
    paste0(code_head, lev_code(dat(), m$grp), "\n",
           if (m$type == "mean") sprintf("summ <- aggregate(%s ~ %s, data = hyp, FUN = mean)\n\n", m$v, m$grp)
           else sprintf("summ <- aggregate(%s ~ %s, data = hyp, FUN = function(x) 100 * mean(x == \"Yes\"))\n\n", m$v, m$grp),
           sprintf("ggplot(summ, aes(x = %s, y = %s)) +\n  geom_col(width = 0.55) +\n  coord_cartesian(ylim = c(%s, NA))   # the truncation: do not do this with bars\n",
                   m$grp, m$v, format(input$mx_start)))
  })

  # =====================================================================================
  # Build your slides
  # =====================================================================================
  b_vars <- reactive(names(dat()))
  output$b_data_info <- renderUI({
    df <- dat()
    div(class = if (src() == "builtin") "teach-note" else "ok-note",
        if (src() == "builtin") "Working on the class dataset. Load your own Excel or CSV file from the sidebar (step 2 there) and every step below uses it."
        else sprintf("Working on your file %s: %d rows \u00d7 %d columns.", src(), nrow(df), ncol(df)))
  })
  output$b_scales <- renderUI({
    df <- dat()
    rows <- lapply(names(df), function(v) {
      x <- df[[v]]; ux <- unique(x[!is.na(x)]); sc <- b_default_scale(df, v)
      desc <- sprintf("%s; %d distinct; %d missing", if (is.numeric(x)) "numbers" else "text", length(ux), sum(is.na(x)))
      div(class = "scale-row",
          div(class = "scale-name", tags$b(v), tags$br(), span(class = "muted", desc)),
          selectInput(paste0("bsc_", v), NULL, width = "100%",
                      choices = c("Quantitative" = "quantitative", "Ordinal" = "ordinal", "Nominal" = "nominal", "Identifier: leave out" = "identifier"),
                      selected = sc),
          conditionalPanel(sprintf("input['bsc_%s'] == 'ordinal'", v),
                           textInput(paste0("bord_", v), NULL, width = "100%",
                                     value = paste(get_levels(df, v), collapse = ", "),
                                     placeholder = "Levels in order, separated by commas")))
    })
    tagList(rows)
  })
  bscale <- function(v) {
    s <- input[[paste0("bsc_", v)]]
    if (is.null(s)) b_default_scale(dat(), v) else s
  }
  bcat <- function(df, v) {
    x <- as.character(df[[v]])
    if (bscale(v) == "ordinal") {
      ord <- input[[paste0("bord_", v)]]
      lv <- if (!is.null(ord) && nzchar(ord)) trimws(strsplit(ord, ",")[[1]]) else get_levels(df, v)
      lv <- c(lv, setdiff(sort(unique(x[!is.na(x)])), lv))
      f <- factor(x, levels = lv)
    } else f <- factor(x, levels = get_levels(df, v))
    lab <- level_labels[[v]]
    if (!is.null(lab) && all(levels(f) %in% names(lab))) levels(f) <- unname(lab[levels(f)])
    droplevels(f)
  }
  b_usable <- reactive({ v <- b_vars(); v[vapply(v, function(z) bscale(z) != "identifier", logical(1))] })
  observe({
    u <- b_usable(); catv <- u[vapply(u, function(z) bscale(z) %in% c("nominal", "ordinal"), logical(1))]
    lc <- function(v) setNames(v, vapply(v, b_label, ""))
    s1 <- isolate(input$b_v1); s2 <- isolate(input$b_v2); sf <- isolate(input$b_f)
    updateSelectInput(session, "b_v1", choices = lc(u), selected = if (!is.null(s1) && s1 %in% u) s1 else pick(u, "SBP_6mo"))
    updateSelectInput(session, "b_v2", choices = c("Nothing" = "", lc(u)), selected = if (!is.null(s2) && s2 %in% c("", u)) s2 else pick(u, "Arm"))
    updateSelectInput(session, "b_f", choices = c("No panels" = "", lc(catv)), selected = if (!is.null(sf) && sf %in% c("", catv)) sf else "")
  })

  # What kind of question is this, and which displays fit it?
  b_base <- reactive({
    df <- dat(); v1 <- req(input$b_v1); v2 <- input$b_v2; f <- input$b_f
    req(v1 %in% names(df))
    if (is.null(v2) || !v2 %in% names(df) || v2 == v1) v2 <- ""
    if (is.null(f) || !f %in% names(df) || f %in% c(v1, v2)) f <- ""
    s1 <- bscale(v1); q1 <- s1 == "quantitative"
    validate(need(s1 != "identifier", "Choose a variable that is not an identifier."))
    if (q1) validate(need(is.numeric(df[[v1]]), sprintf("%s is text, so it cannot be quantitative: set its scale to nominal or ordinal.", v1)))
    # with no second variable, a panel variable acts as the group
    if (v2 == "" && f != "") { v2 <- f; f <- "" }
    s2 <- if (v2 != "") bscale(v2) else ""; q2 <- s2 == "quantitative"
    if (q2) validate(need(is.numeric(df[[v2]]), sprintf("%s is text, so it cannot be quantitative.", v2)))
    kind <- if (v2 == "") (if (q1) "quant1" else if (s1 == "ordinal") "ordinal1" else "nominal1")
            else if (q1 && q2) "qq" else if (q1 || q2) "qc" else "cc"
    yv <- v1; gv <- v2
    if (kind == "qc" && !q1) { yv <- v2; gv <- v1 }
    keep <- !is.na(df[[v1]]) & (if (v2 != "") !is.na(df[[v2]]) else TRUE) & (if (f != "") !is.na(df[[f]]) else TRUE)
    dd <- df[keep, , drop = FALSE]
    validate(need(nrow(dd) >= 3, "Fewer than three complete rows for this choice."))
    d <- switch(kind,
      quant1 = data.frame(val = dd[[v1]]),
      ordinal1 = , nominal1 = data.frame(cat = bcat(dd, v1)),
      qc = data.frame(val = dd[[yv]], g = bcat(dd, gv)),
      cc = data.frame(cat = bcat(dd, v1), g = bcat(dd, v2)),
      qq = data.frame(val = dd[[v1]], x = dd[[v2]]))
    if (f != "") d$f <- bcat(dd, f)
    info <- list(n = nrow(d), nlev = if (!is.null(d$cat)) nlevels(d$cat) else 0,
                 maxg = if (!is.null(d$g)) max(table(d$g)) else nrow(d),
                 equal = if (!is.null(d$g)) length(unique(as.integer(table(d$g)))) == 1 else TRUE,
                 ties = if (kind == "qq") 1 - nrow(unique(d[, c("val", "x")])) / nrow(d) else 0)
    recs <- b_recommend(kind, info)
    if (f != "") recs <- recs[recs$type != "hist_facet", ]
    list(kind = kind, d = d, v1 = v1, v2 = v2, yv = yv, gv = gv, f = f, s1 = s1, recs = recs, info = info)
  })
  output$b_kind <- renderUI({
    b <- b_base()
    txt <- switch(b$kind,
      quant1 = sprintf("One quantitative variable (%s), n = %d.", b$v1, b$info$n),
      ordinal1 = sprintf("One ordinal variable (%s) with %d levels, n = %d.", b$v1, b$info$nlev, b$info$n),
      nominal1 = sprintf("One nominal variable (%s) with %d categories, n = %d.", b$v1, b$info$nlev, b$info$n),
      qc = sprintf("A quantitative outcome (%s) compared across the groups of %s, n = %d.", b$yv, b$gv, b$info$n),
      cc = sprintf("Two categorical variables: %s within the groups of %s, n = %d.", b$v1, b$v2, b$info$n),
      qq = sprintf("Two quantitative variables: %s against %s, n = %d.", b$v1, b$v2, b$info$n))
    div(class = "kind-chip", txt, if (b$f != "") sprintf(" One panel per %s.", b$f))
  })
  output$b_options <- renderUI({
    r <- b_base()$recs; ok <- r[r$verdict != "avoid", ]; bad <- r[r$verdict == "avoid", ]
    cur <- isolate(input$b_type)
    sel <- if (!is.null(cur) && cur %in% ok$type) cur else ok$type[1]
    tagList(
      radioButtons("b_type", "Recommended displays", selected = sel,
                   choiceValues = ok$type,
                   choiceNames = lapply(seq_len(nrow(ok)), function(i)
                     span(span(class = paste("badge-v", ok$verdict[i]), if (ok$verdict[i] == "best") "Best" else "Good"),
                          tags$b(ok$label[i]), span(class = "muted", paste0(" \u2014 ", ok$why[i]))))),
      if (nrow(bad)) div(class = "avoid-box", tags$b("Avoid"),
                         tags$ul(lapply(seq_len(nrow(bad)), function(i) tags$li(tags$b(bad$label[i]), ": ", bad$why[i])))))
  })
  output$b_extra <- renderUI({
    t <- input$b_type; req(t)
    if (t %in% c("bar_freq", "bar_order")) radioButtons("b_pct", "Bar height", c("Count" = "n", "Percentage" = "pct"), inline = TRUE)
    else if (t %in% c("scatter", "jitter", "count")) checkboxInput("b_lm", "Add a straight-line fit", FALSE)
    else if (t %in% c("hist", "hist_facet")) numericInput("b_bw", "Bin width", value = default_w(b_base()$d$val), min = 0.01)
  })

  # Specification of the figure being edited
  b_defaults <- reactive({
    b <- b_base(); t <- req(input$b_type); req(t %in% b$recs$type)
    u <- if (nzchar(trimws(input$b_unit))) trimws(input$b_unit) else "observation"
    lab1 <- b_label(b$v1); labg <- if (b$gv != "") b_label(b$gv) else ""
    cnt <- paste0("Number of ", u, "s")
    xl <- switch(b$kind, quant1 = if (t %in% c("strip", "box")) "" else lab1, ordinal1 = , nominal1 = lab1,
                 qc = if (t == "hist_facet") b_label(b$yv) else "", cc = lab1, qq = b_label(b$v2))
    yl <- switch(b$kind,
                 quant1 = switch(t, hist = cnt, ecdf = paste0("Cumulative % of ", u, "s"), lab1),
                 ordinal1 = , nominal1 = if (t == "dot_pct") paste0("% of ", u, "s") else cnt,
                 qc = if (t == "hist_facet") cnt else b_label(b$yv),
                 cc = if (t == "cluster_n") cnt else paste0("% of ", u, "s within each group"),
                 qq = lab1)
    if (b$kind == "cc" && t == "stack_pct") xl <- labg
    title <- switch(b$kind, quant1 = paste("Distribution of", low1(lab1)), ordinal1 = , nominal1 = lab1,
                    qc = paste(b_label(b$yv), "by", low1(labg)), cc = paste(lab1, "by", low1(labg)),
                    qq = paste(lab1, "and", low1(b_label(b$v2))))
    if (b$f != "") title <- paste0(title, ", by ", low1(b_label(b$f)))
    list(title = title, xlab = xl, ylab = yl)
  })
  b_fill <- function() {
    dflt <- b_defaults()
    updateTextInput(session, "b_title", value = dflt$title)
    updateTextInput(session, "b_xlab", value = dflt$xlab)
    updateTextInput(session, "b_ylab", value = dflt$ylab)
    sp <- b_spec(use_defaults = TRUE)
    updateTextAreaInput(session, "b_caption", value = b_caption_default(sp))
  }
  observeEvent(list(input$b_type, input$b_v1, input$b_v2, input$b_f, input$b_unit, input$b_lm, input$b_bw, dat()), b_fill(), ignoreInit = FALSE)
  observeEvent(input$b_refill, b_fill())

  b_spec <- function(use_defaults = FALSE) {
    b <- b_base(); t <- req(input$b_type); req(t %in% b$recs$type)
    dflt <- b_defaults()
    pickt <- function(id, dv) { z <- input[[id]]; if (use_defaults || is.null(z)) dv else z }
    u <- if (nzchar(trimws(input$b_unit))) trimws(input$b_unit) else "observation"
    bw <- input$b_bw; if (is.null(bw) || is.na(bw) || bw <= 0) bw <- if (!is.null(b$d$val)) default_w(b$d$val) else 1
    rng <- function(z) diff(range(z, na.rm = TRUE))
    sp <- list(type = t, verdict = b$recs$verdict[b$recs$type == t], d = b$d, scale1 = b$s1, unit = u,
               lab1 = b_label(b$v1), laby = b_label(b$yv), labg = if (b$gv != "") b_label(b$gv) else "",
               labx = if (b$kind == "qq") b_label(b$v2) else "",
               xlab = pickt("b_xlab", dflt$xlab), ylab = pickt("b_ylab", dflt$ylab), title = pickt("b_title", dflt$title),
               pct = identical(input$b_pct, "pct"), lm = isTRUE(input$b_lm), bw = bw,
               jw = if (b$kind == "qq") 0.02 * rng(b$d$x) + if (all(b$d$x == round(b$d$x))) 0.2 else 0 else 0,
               jh = if (b$kind == "qq") 0.02 * rng(b$d$val) + if (all(b$d$val == round(b$d$val))) 0.2 else 0 else 0)
    sp$caption <- if (use_defaults) "" else input$b_caption
    if (use_defaults || is.null(sp$caption) || !nzchar(sp$caption)) sp$caption <- b_caption_default(sp)
    sp
  }
  b_current <- reactive(b_spec())
  output$b_plot <- renderPlot({ sp <- b_current(); b_draw(sp, fs()) }, res = 96)
  output$b_caption_show <- renderUI(div(class = "caption-box", b_current()$caption))
  output$b_checks <- renderUI({
    ck <- b_checks(b_current())
    tags$ul(class = "checks auto-checks",
            lapply(ck, function(k) tags$li(HTML(paste0(if (k$ok) "<span class='ck-ok'>\u2713</span> " else "<span class='ck-no'>\u2717</span> ",
                                                      htmltools::htmlEscape(k$txt))))))
  })

  # Slides
  slides <- reactiveVal(list())
  observeEvent(input$b_add, {
    sp <- b_current()
    s <- list(id = paste0("s", as.integer(Sys.time()), sample.int(1e6, 1)), title = sp$title, caption = sp$caption,
              notes = if (is.null(input$b_notes)) "" else input$b_notes, kind = sp$type,
              label = b_base()$recs$label[b_base()$recs$type == sp$type],
              plot = b_draw(sp, 18))
    slides(c(slides(), list(s)))
    updateTextAreaInput(session, "b_notes", value = "")
    showNotification(sprintf("Slide %d added: %s", length(slides()), sp$title), type = "message", duration = 4)
  })
  observeEvent(input$b_clear, slides(list()))
  observeEvent(input$b_act, {
    a <- input$b_act; s <- slides(); i <- which(vapply(s, `[[`, "", "id") == a$id); req(length(i) == 1)
    if (a$act == "del") s <- s[-i]
    else if (a$act == "up" && i > 1) s[c(i - 1, i)] <- s[c(i, i - 1)]
    else if (a$act == "down" && i < length(s)) s[c(i, i + 1)] <- s[c(i + 1, i)]
    slides(s)
  })
  output$b_slides <- renderUI({
    s <- slides()
    if (!length(s)) return(div(class = "muted", style = "padding:.6rem 0",
                               "No slides yet. Build a figure above and press \u201cAdd this figure as a slide\u201d. The title slide is added automatically."))
    act <- function(id, a, lab, title) tags$a(href = "#", class = "slide-act", title = title,
      onclick = sprintf("Shiny.setInputValue('b_act', {id: '%s', act: '%s', n: Math.random()}); return false;", id, a), lab)
    tags$ol(class = "slide-list",
            tags$li(class = "slide-title", tags$b("Title slide: "), input$b_deck_title),
            lapply(s, function(x) tags$li(
              div(tags$b(x$title), tags$br(), span(class = "muted", x$label)),
              div(class = "slide-acts", act(x$id, "up", "\u25b2", "Move up"), act(x$id, "down", "\u25bc", "Move down"),
                  act(x$id, "del", "\u2715", "Remove")))))
  })
  b_deck <- reactive(list(title = if (nzchar(input$b_deck_title)) input$b_deck_title else "Displaying the data",
                          sub = input$b_deck_sub))
  output$b_pptx <- downloadHandler(
    filename = function() "presentation.pptx",
    content = function(file) {
      s <- isolate(slides())
      if (!length(s)) { showNotification("Add at least one slide first.", type = "warning"); req(FALSE) }
      b_write_pptx(isolate(b_deck()), s, file)
    })
  output$b_pdf <- downloadHandler(
    filename = function() "presentation.pdf",
    content = function(file) {
      s <- isolate(slides())
      if (!length(s)) { showNotification("Add at least one slide first.", type = "warning"); req(FALSE) }
      b_write_pdf(isolate(b_deck()), s, file)
    })

  # ---- Checkpoints -----------------------------------------------------------------------
  cpv <- reactiveValues(shown = rep(FALSE, length(cps)), sel = rep(NA_character_, length(cps)),
                        res = rep(NA, length(cps)), all = FALSE)
  lapply(seq_along(cps), function(i) {
    observeEvent(input[[paste0("cp_go_", i)]], {
      sel <- input[[paste0("cp_", i)]]
      sh <- cpv$shown; sh[i] <- TRUE; cpv$shown <- sh
      s <- cpv$sel; r <- cpv$res
      if (is.null(sel) || !length(sel)) { s[i] <- NA; r[i] <- NA } else { s[i] <- sel; r[i] <- sel == cps[[i]]$ans }
      cpv$sel <- s; cpv$res <- r
    })
    output[[paste0("cp_fb_", i)]] <- renderUI({
      cp <- cps[[i]]
      if (cpv$all && !cpv$shown[i]) return(div(class = "ok-note", strong(paste0("Answer: ", cp$ans, ". ")), cp$why))
      req(cpv$shown[i])
      sel <- cpv$sel[i]
      if (is.na(sel)) return(div(class = "teach-note", "Choose an option first."))
      if (sel == cp$ans) div(class = "ok-note", strong(paste0("Correct: ", cp$ans, ". ")), cp$why)
      else div(class = "warn-note", strong(paste0("Not ", sel, ". The answer is ", cp$ans, ". ")), cp$why)
    })
  })
  observeEvent(input$cp_reveal, cpv$all <- TRUE)
  observeEvent(input$cp_reset, {
    cpv$shown <- rep(FALSE, length(cps)); cpv$sel <- rep(NA_character_, length(cps))
    cpv$res <- rep(NA, length(cps)); cpv$all <- FALSE
    session$sendCustomMessage("cp_clear", paste0("cp_", seq_along(cps)))
  })
  output$cp_score <- renderUI({
    r <- cpv$res
    tile_set <- lapply(seq_along(cp_sets), function(j) {
      ids <- cp_sets[[j]]
      tile(sprintf("%d / %d", sum(r[ids] %in% TRUE), length(ids)), paste(names(cp_sets)[j], "correct"), c(1, 5)[j])
    })
    div(class = "tiles",
        tile(sprintf("%d / %d", sum(!is.na(r)), length(cps)), "checkpoints answered", 2),
        tile_set[[1]], tile_set[[2]],
        tile(sprintf("%d", sum(r %in% FALSE)), "to revisit", 3))
  })
}

shinyApp(ui, server)
