#This is a script to run a MANOVA on pollen data from sediment samples
#Arrange data so your independent variable (sample depth) is in the first column and the dependent variables (species) are in the 2nd, 3rd, 4th columns etc.
#Note, your no. of samples must be more than your no. of dependent variables. MANOVA will not work otherwise. To avoid this, you can group species i.e. Podocarps, Beech, Small trees and shrubs.

# Load the necessary libraries
library(MASS)
library(openxlsx)

# Read your dataset (replace 'your_data.csv' with the actual file path)
data <- read.csv("TAN1613_47_GG.csv", header = TRUE)

# Define the dependent variables (species columns) as a matrix
species_data <- as.matrix(data[, -1])  # Exclude the 'Sample depth' column

# Define the independent variable (sample depth)
sample_depth <- data$Sample.depth

# Perform MANOVA
manova_result <- manova(species_data ~ sample_depth)

# Print the summary of MANOVA results
summary(manova_result)

# Extract MANOVA results into a table
manova_table <- summary(manova_result)$stats

# Print the table
print(manova_table)

# Convert MANOVA table to data frame
manova_df <- as.data.frame(manova_table)

# Write MANOVA results to an Excel file
write.xlsx(manova_df, "MANOVA_Results.xlsx")

# Calculate residuals
residuals_matrix <- residuals(manova_result)

# Convert residuals matrix to a data frame
residuals_df <- as.data.frame(residuals_matrix)

# Write residuals to an Excel file
write.xlsx(residuals_df, "MANOVA_Residuals.xlsx")

