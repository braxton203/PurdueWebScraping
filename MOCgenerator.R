# Load necessary libraries
library(readxl)
library(dplyr)

# Function to search for a chemical in all sheets of an Excel file and return the associated row(s) in matrix form
search_chemical_matrix <- function(file_path, chemical_column, chemical_name) {
  # Get all sheet names from the Excel file
  sheet_names <- excel_sheets(file_path)
  
  # Initialize an empty list to store data from each sheet
  all_data <- list()
  
  # Loop through each sheet and extract the relevant data
  for (sheet in sheet_names) {
    # Read each sheet into a data frame
    data <- read_excel(file_path, sheet = sheet)
    
    # Convert to a data frame if needed
    data <- as.data.frame(data)
    
    # Check if the specified chemical column exists
    if (chemical_column %in% colnames(data)) {
      # Filter rows that match the given chemical name
      result <- data %>%
        filter(!!sym(chemical_column) == chemical_name)
      
      # If rows are found, store them in the list
      if (nrow(result) > 0) {
        all_data[[sheet]] <- result
      }
    }
  }
  
  # Combine all the results into a single data frame, if any data was found
  if (length(all_data) > 0) {
    # Bind all results by row (combine all sheets together)
    combined_data <- bind_rows(all_data)
    
    # Extract the unique chemicals and materials
    chemicals <- combined_data[[chemical_column]]
    materials <- colnames(combined_data)[colnames(combined_data) != chemical_column] # Exclude the chemical column
    
    # Create a matrix where rows are chemicals and columns are materials
    chemical_matrix <- matrix(NA, nrow = length(unique(chemicals)), ncol = length(materials),
                              dimnames = list(unique(chemicals), materials))
    
    # Populate the matrix with corresponding values for each chemical
    for (i in 1:nrow(combined_data)) {
      chemical <- combined_data[i, chemical_column]
      for (material in materials) {
        chemical_matrix[chemical, material] <- combined_data[i, material]
      }
    }
    
    # Return the matrix
    return(chemical_matrix)
    
  } else {
    return(paste("Chemical", chemical_name, "not found in any sheet."))
  }
}

# Example usage:
# Define the path to Excel file and other parameters
file_path <- "~/Desktop/Purdue Projects/PurdueDataMine/PurdueTest.xlsx"  # Path to your Excel file
chemical_column <- "Chemical"  # The column name that contains the chemical list

# Ask the user for the chemical they want to search
chemical_name <- readline(prompt = "Enter the chemical name to search: ")

# Call the function and store the result
result_matrix <- search_chemical_matrix(file_path, chemical_column, chemical_name)

# Print the result as a matrix
print(result_matrix)

