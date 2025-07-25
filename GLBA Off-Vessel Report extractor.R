##Daniel Solorzano-Jones
## 7/24/2025

library(tidyr)
library(dplyr)
library(readxl)


#### Function to process a single Excel file ####
process_excel_file <- function(file_path) {
  # Read the Excel file
  data <- read_excel(file_path, sheet = "GLBA Off-Vessel Report")
  
  # Extract the Vessel Name from a specific cell (adjust the cell reference as needed)
  vessel_name <- data[2, 4]  # Assuming Vessel Name is in the first row
  
  # Process the data
  tidy_data <- data %>%
    select("Type of Activity Select from drop down list", "Date             (mm/dd/yy)",
           "Start Time of Activity           (hh:mm)", "End Time of Activity           (hh:mm)", 
           "No. of Passengers in Group", "No. of Crew in group", "Location ",
           "Location Visited Detail                                                                Describe exact location or submit map",
           "Comments and Observations") %>%
    mutate(
      Vessel_Name = vessel_name,
      Total_People = "No. of Passengers in Group" + "No. of Crew in group"
    ) %>%
    rename(
      Activity = "Type of Activity Select from drop down list",
      Passengers = "No. of Passengers in Group",
      Crew = "No. of Crew in group",
      Start_Time = "Start Time of Activity           (hh:mm)",
      End_Time = "End Time of Activity           (hh:mm)",
      Date = "Date             (mm/dd/yy)",
      Location_of_Activity = "Location ",
      Location_Detail = "Location Visited Detail                                                                Describe exact location or submit map",
      Comments = "Comments and Observations"
    ) %>%
    select(Vessel_Name, Activity, Passengers, Crew, Total_People, Start_Time, End_Time, Date, Location_of_Activity, Location_Detail, Comments)
  
  return(tidy_data)
}

# Main function to process multiple Excel files and save to CSV
process_multiple_files <- function(file_paths, output_file) {
  all_data <- bind_rows(lapply(file_paths, process_excel_file))
  
  # Write the combined data to a CSV file
  write.csv(all_data, output_file, row.names = FALSE)
}

# Example usage
# Specify the paths to your Excel files
excel_files <- list.files(path = "C:/Users/dsolorzano-jones/Documents/Test", pattern = "*.xlsx", full.names = TRUE)

# Specify the output CSV file path
output_csv_file <- "C:/Users/dsolorzano-jones/Documents/Test/Output CSV/2025combined_data.csv"

# Process the files and create the CSV
process_multiple_files(excel_files, output_csv_file)

cat("Data processing complete. Output saved to:", output_csv_file)