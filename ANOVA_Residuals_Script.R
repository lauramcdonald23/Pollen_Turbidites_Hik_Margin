install.packages("openxlsx")
library(openxlsx)

# Read the csv, pollenpercent will be the name of the data sheet
pollenpercent <- read.csv("TAN1613_58.csv")

# Create an empty list to store the ANOVA results
pollenpercent_anova_results <- list()

# Create an empty list to store the residuals
residuals_list <- list()

# Iterate over each species column
species_columns <- names(pollenpercent)[-1]  # Exclude the first column (Sample.depth)
for (species in species_columns) {
  # Perform ANOVA for the current species
  formula <- as.formula(paste(species, "~ Sample.depth"))
  result <- aov(formula, data = pollenpercent)
  
  # Store the ANOVA result in the list
  pollenpercent_anova_results[[species]] <- summary(result)
  
  
  # Get the residuals and store them in the list with species name as key
  residuals <- residuals(result)
  residuals_list[[species]] <- residuals
  
  # Print ANOVA results for the current species
  cat("Species:", species, "\n")
  print(summary(result))
  cat("\n")

}

# Convert residuals list to a data frame
residuals_table <- as.data.frame(do.call(cbind, residuals_list))

# Create an empty data frame to store ANOVA results
anova_results_df <- data.frame(Species = species_columns, 
                               Df = numeric(length(species_columns)),
                               Sum_Sq = numeric(length(species_columns)),
                               Mean_Sq = numeric(length(species_columns)),
                               F_value = numeric(length(species_columns)),
                               Pr_gt_F = numeric(length(species_columns)))

# Iterate over each species column
for (i in seq_along(species_columns)) {
  species <- species_columns[i]
  # Extract relevant information from the ANOVA summary
  summary_result <- pollenpercent_anova_results[[species]]
  df <- summary_result[[1]]$"Df"[1]
  sum_sq <- summary_result[[1]]$"Sum Sq"[1]
  mean_sq <- summary_result[[1]]$"Mean Sq"[1]
  F_value <- summary_result[[1]]$"F value"[1]
  pr_gt_f <- summary_result[[1]]$"Pr(>F)"[1]
  
  # Store the information in the data frame
  anova_results_df[i, "Df"] <- df
  anova_results_df[i, "Sum_Sq"] <- sum_sq
  anova_results_df[i, "Mean_Sq"] <- mean_sq
  anova_results_df[i, "F_value"] <- F_value
  anova_results_df[i, "Pr_gt_F"] <- pr_gt_f
}

# Print the ANOVA results table
print(anova_results_df)

write.xlsx(anova_results_df, "ANOVA_Results_58.xlsx")

write.xlsx(residuals_table, "Residuals_Results_58.xlsx")

