# Script for analyzing platelet counts and survival time in cancer patients
# This script evaluates platelet counts based on survival time, using Mann-Whitney tests 
# to compare groups with different survival times. Adjusted for the given dataset structure.

# Load required libraries
library(readxl)  # For reading Excel files
library(openxlsx)  # For writing Excel files

# Define the path to the data file
file_path <- "path/to/your/file.xlsx"  # Replace with your actual file path

# Read the data sheet
data <- read_excel(file_path, sheet = "SheetName")  # Replace "SheetName" with the sheet containing your data


# Define relevant variables
parameter <- "Platelet count (/ul)"  # Parameter of interest
survival_time <- "Survival time (month)"  # Time variable
filters <- list(
  Diagnosis = "Diagnosis",                # Diagnosis filter
  Histogenesis = "Histogenesis",          # Histogenesis filter
  Malignancy = "Malignancy",              # Malignancy filter
  Histogenesis_Malignancy = c("Histogenesis", "Malignancy")  # Combined filter for histogenesis type and malignancy
)

# Define the vector of cutoff points (from 4 to 24 months)
cutoff_months <- 4:24

# Initialize an empty list to store sheets for the Excel file
sheets <- list()

# Loop through each filter
for (filter_name in names(filters)) {
  filter_results <- list()  # Store results for the current filter
  
  # Loop through each cutoff point
  for (month in cutoff_months) {
    # Create two groups based on survival time
    group_less_equal <- data[data[[survival_time]] <= month & !is.na(data[[survival_time]]), ]
    group_greater <- data[data[[survival_time]] > month & !is.na(data[[survival_time]]), ]
    
    # Get categories based on the current filter
    if (length(filters[[filter_name]]) == 1) {
      categories <- unique(data[[filters[[filter_name]]]])
    } else {
      categories <- unique(data[, filters[[filter_name]]])
      categories <- unique(apply(categories, 1, paste, collapse = " + "))
    }
    categories <- categories[!is.na(categories)]
    
    # Test the difference in counts between the two groups for each category
    for (category in categories) {
      if (length(filters[[filter_name]]) == 1) {
        group_less_equal_category <- group_less_equal[group_less_equal[[filters[[filter_name]]]] == category, ]
        group_greater_category <- group_greater[group_greater[[filters[[filter_name]]]] == category, ]
      } else {
        category_split <- strsplit(category, " \\+ ")[[1]]
        group_less_equal_category <- group_less_equal[
          group_less_equal[[filters[[filter_name]][1]]] == category_split[1] &
            group_less_equal[[filters[[filter_name]][2]]] == category_split[2], ]
        group_greater_category <- group_greater[
          group_greater[[filters[[filter_name]][1]]] == category_split[1] &
            group_greater[[filters[[filter_name]][2]]] == category_split[2], ]
      }
      
      # Correct the sample count for each group
      n_less_equal <- sum(!is.na(group_less_equal_category[[parameter]]))
      n_greater <- sum(!is.na(group_greater_category[[parameter]]))
      
      if (n_less_equal > 1 & n_greater > 1) {
        # Perform the Mann-Whitney test for the current parameter
        mann_whitney_test <- wilcox.test(group_less_equal_category[[parameter]], group_greater_category[[parameter]])
        
        # Calculate descriptive statistics for each group
        stats_less_equal <- data.frame(
          mean = mean(group_less_equal_category[[parameter]], na.rm = TRUE),
          sd = sd(group_less_equal_category[[parameter]], na.rm = TRUE),
          median = median(group_less_equal_category[[parameter]], na.rm = TRUE),
          Q1 = quantile(group_less_equal_category[[parameter]], 0.25, na.rm = TRUE),
          Q3 = quantile(group_less_equal_category[[parameter]], 0.75, na.rm = TRUE)
        )
        stats_greater <- data.frame(
          mean = mean(group_greater_category[[parameter]], na.rm = TRUE),
          sd = sd(group_greater_category[[parameter]], na.rm = TRUE),
          median = median(group_greater_category[[parameter]], na.rm = TRUE),
          Q1 = quantile(group_greater_category[[parameter]], 0.25, na.rm = TRUE),
          Q3 = quantile(group_greater_category[[parameter]], 0.75, na.rm = TRUE)
        )
        
        # Store the results
        filter_results[[paste0("Cutoff_", month, "_Category_", category)]] <- data.frame(
          Cutoff = month,
          Category = category,
          Statistic = mann_whitney_test$statistic,
          P_Value = mann_whitney_test$p.value,
          N_Less_Equal = n_less_equal,
          N_Greater = n_greater,
          Mean_Less_Equal = stats_less_equal$mean,
          SD_Less_Equal = stats_less_equal$sd,
          Median_Less_Equal = stats_less_equal$median,
          Q1_Less_Equal = stats_less_equal$Q1,
          Q3_Less_Equal = stats_less_equal$Q3,
          Mean_Greater = stats_greater$mean,
          SD_Greater = stats_greater$sd,
          Median_Greater = stats_greater$median,
          Q1_Greater = stats_greater$Q1,
          Q3_Greater = stats_greater$Q3
        )
      }
    }
  }
  
  # Combine the results for all categories into a single data frame
  filter_results_df <- do.call(rbind, filter_results)
  
  # Add the result as a sheet to the list
  sheets[[filter_name]] <- filter_results_df
}

# Write the results to an Excel file with one sheet per filter
output_file <- "Platelet_Analysis_Results.xlsx"
write.xlsx(sheets, file = output_file, rowNames = FALSE)

# End of script
