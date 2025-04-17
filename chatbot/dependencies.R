# dependencies.R

# Function to install missing R packages
install_if_missing <- function(pkg, github = NULL) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    if (!is.null(github)) {
      message(paste("Installing", pkg, "from GitHub (", github, ") ..."))
      remotes::install_github(github)
    } else {
      message(paste("Installing", pkg, "..."))
      install.packages(pkg)
    }
  } else {
    message(paste(pkg, "already installed."))
  }
}

# Install remotes first (needed for GitHub packages)
install_if_missing("remotes")

# List of standard R packages (from CRAN)
required_packages <- c(
  "DT",
  "shiny",
  "shinyjs",
  "httr",
  "pdftools",
  "later",
  "officer",
  "promises",
  "future",
  "mime",
  "reticulate"
)

# Install missing R packages
invisible(lapply(required_packages, install_if_missing))

# Check system dependency for pdftools / Poppler
check_poppler <- function() {
  sys_name <- Sys.info()["sysname"]
  
  if (sys_name == "Linux") {
    # Check if libpoppler-cpp is installed
    result <- suppressWarnings(system("ldconfig -p | grep libpoppler-cpp", intern = TRUE))
    if (length(result) == 0) {
      message("⚠️  Poppler library NOT detected (required by pdftools).")
      message("👉 On Ubuntu/Debian, run this in the terminal:")
      message("   sudo apt-get update && sudo apt-get install -y libpoppler-cpp-dev")
    } else {
      message("✅ Poppler library detected.")
    }
  } else if (sys_name == "Darwin") {
    message("⚠️  On macOS, ensure Poppler is installed.")
    message("👉 Run: brew install poppler")
  } else {
    message("ℹ️  Poppler system library check is not configured for this OS.")
  }
}

reticulate::py_install("extract-msg")



#install.packages("remotes")
#remotes::install_github("hrbrmstr/msgxtractr")

message("✅ dependencies.R loaded and Python + R environments are ready.")
