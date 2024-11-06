#Last updated: Oct 23, 2024 by KCheng

#READ ME
#This code is used to interpret fiber data based on user-defined arguments (thresholding etc.)
#This version is compatible with excel data inputs, as well as and excel outputs
#Please review the code  on lines 31- 71 for "#NEED INPUT", these are user-defined arguments that require your input, however there are placeholder values in place, and some are optional.
#On lines 25-34, are code for installing certain packages, these are only done *once* and you *MUST* install before using, please comment them out using "#"" before the line once you're done


###PART 1- The Once in the (computer's) lifetime Downloads ###

#WHEN YOU FINISH INSTALLING IT ONCE (THE FIRST TIME YOU RUN IT),ADD A "#" INFRONT OF THE LINE OF CODE THAT SAYS "1st TIME ONLY"
#Install packages (only once ever)
#install.packages(c('readxl','openxlsx','ggplot2','dplyr', 'tidyr', 'grid','multcomp', 'multcompView', 'rlang', 'broom'); #1st TIME ONLY

#If you have Windows (only once ever)
#install.packages('installr'); #1st TIME ONLY

#install.Rtools(); ##must have PATH to Rtools()
#set PATH to Rtools(, (only once ever)
#Sys.setenv(R_ZIPCMD= "C:/Program Files/Rtools/bin/zip"); #NEEDINPUT, choose file path + 1st TIME ONLY


###PART 2- The Set Up (this part needs your participation, or not...) ###
# Loading libraries in the R workspace
library(readxl)
library(openxlsx)
library(ggplot2)
library(dplyr)
library(tidyr)
library(grid)
library(multcomp)
library(multcompView)
library(rlang)
library(broom)

# User-input parameters
xlstart_row <- 6 # NEEDINPUT, choose starting row

DNA_thres_a <- 100 # NEEDINPUT, choose DNA threshold
tract_thres_b <- 200 # NEEDINPUT, choose tract threshold
prot1_thres_c <- 200 # NEED INPUT, choose protein 1 threshold
prot2_thres_d <- 200 # NEED INPUT, choose protein 2 threshold

Smooth_it_1 <- 3 # NEED INPUT, choose smoothing value for DNA 
Smooth_it_2 <- 3 # NEED INPUT, choose smoothing value for tracts
Smooth_it_3 <- 6 # NEED INPUT, choose smoothing value for protein 1
Smooth_it_4 <- 6 # NEED INPUT, choose smoothing value for protein 2

unrep_fork <- 2 # NEED INPUT, choose a the amount of unreplicated pixels that would be included in the fork
rep_fork <- 2 # NEED INPUT, choose a the amount of replicated pixels that would be included in the fork

#NOTE: These are your bins, if you want to group different sheets (eg. replicates), you can do so here! If you don't want to set bins or you have less than 6 bins, use "NA" to fill the rest of the inputs bins.
bin_1 <- 'wt' #NEED INPUT, if you want to bin your experiments, input a keyword that is common to all the experiments within quotations (eg. "wt")
bin_2 <- 'mrc' #NEED INPUT, if you want to bin your experiments, input a keyword that is common to all the experiments within quotations (eg. "wt")
bin_3 <- 'cds' #NEED INPUT, if you want to bin your experiments, input a keyword that is common to all the experiments within quotations (eg. "wt")
bin_4 <- 'na' #NEED INPUT, if you want to bin your experiments, input a keyword that is common to all the experiments within quotations (eg. "wt")
bin_5 <- 'na' #NEED INPUT, if you want to bin your experiments, input a keyword that is common to all the experiments within quotations (eg. "wt")
#bin_6 <- "more bins?" #NEED INPUT, if you want to bin your experiments, input a keyword that is common to all the experiments within quotations (eg. "wt")

outfilename <- "Output file.xlsx" # NEED INPUT, choose output file name
output_folder <- "C:/Users/Keren/Desktop/Output_Folder" # NEED INPUT, choose the folder path to store all outputs +add an "/Output_Folder" at the end
dir.create(output_folder, showWarnings = FALSE) # Create the output folder if it doesn't exist
outfile <- file.path(output_folder, outfilename)

#NOTE: This is your upper limits for plots. Adding user inputs for the below would work the best after you run the code with your desired threshold and smoothing. Use "NA" if you don't want to set a limit.
scatter_plot_ymax <- NA  #NEED INPUT, add an upper limit for the individual sheet scatter plot of initial pixel intensity in PART 3
ind_coloc_p_ymax <-NA  #NEED INPUT, add an upper limit for the individual sheet bar graph of protein co-localization (column 3 and 4 co-localization to tract) in PART 5
bin_coloc_p_ymax <- NA  #NEED INPUT, add an upper limit for the binned bar graph of protein co-localization in REF 10.5
ind_tract_ymax <- NA  #NEED INPUT, add an upper limit for the individual violin plot of tract lengths in REF 7.5
bin_tract_ymax <-NA  #NEED INPUT, add an upper limit for the binned bar graph of tract lengths in REF 7.5

#Code start here!
parameters <- data.frame(DNA_thres_a, tract_thres_b, prot1_thres_c, prot2_thres_d, Smooth_it_1, Smooth_it_2, Smooth_it_3, Smooth_it_4, unrep_fork, rep_fork)
plotyaxis <-data.frame(scatter_plot_ymax, ind_coloc_p_ymax, bin_coloc_p_ymax, ind_tract_ymax, bin_tract_ymax)

# Select input file and obtaining information on possible sheets
path <- file.choose()
file_name <- excel_sheets(path)

# Recording parameters, time, and original file and sheet in excel output sheet
wb <- createWorkbook()
addWorksheet(wb, "Threshold Parameters")
noteparameters <- "These were your following parameters: "
writeData(wb, "Threshold Parameters", x = noteparameters, startCol = 1, startRow = 5)
writeData(wb, "Threshold Parameters", parameters, startCol = 1, startRow = 6)

timestamp <- format(Sys.time(), "%Y-%m-%d_%H-%M")
notedate_time <- "Date and time: "
writeData(wb, "Threshold Parameters", x = notedate_time, startCol = 1, startRow = 10)
writeData(wb, "Threshold Parameters", timestamp, startCol = 1, startRow = 11)

ogfilename <- sub("_.*", ".xlsx", basename(path))
notetitle <- "RODDBLOBS have been applied to "
output_title <- paste(notetitle, ogfilename)
writeData(wb, "Threshold Parameters", x = output_title, startCol = 1, startRow = 1)

noteyaxis <- "These are your input for y-axis upper limit (if any)"
writeData(wb, "Threshold Parameters", x = noteyaxis, startCol = 1, startRow = 13)
writeData(wb, "Threshold Parameters", x = plotyaxis, startCol = 1, startRow = 14)

#Initialize an empty list to store longest_lengths_df for each sheet
all_longest_lengths_col1 <- list()
all_longest_lengths_col2 <- list()
all_longest_lengths_col3 <- list()
all_longest_lengths_col4 <- list()

#Intitalize an empty list to store coloc_df for each sheet
bin_col <- list()

# Loop over each sheet in the Excel file
for (sheet in file_name) {
  
  ### PART 3 - Plotting the initial data ###
  data <- read_xlsx(path, sheet = sheet, na = " ", skip = (xlstart_row - 1))
  
  # Just in case debugging step: Print the first few rows of the data- if empty or not the same as the og, then it might not be calling the correct data (Can comment out)
  #print(head(data))
  
  # Reshape the data to long format for plotting
  data_long <- pivot_longer(data, cols = everything(), names_to = "Variable", values_to = "Value")
  
  # Just in case debugging step: Print the first few rows of the reshaped data- if this print out to console is empty then the program might not be calling the correct data (can comment out)
  #print(head(data_long))
  
  plot_idata <- file.path(output_folder, paste0("scatter_plot_", sheet, ".png"))
  
  # Create the initial plot and store it in a variable
  p <- ggplot(data_long, aes(x = seq_along(Value), y = Value, color = Variable)) +
    geom_point(aes(size = 1.5), show.legend = TRUE) +  # Change show.legend to TRUE
    scale_color_manual(values = rainbow(length(unique(data_long$Variable)))) +
    labs(title = paste("Scatter Plot of Initial Data - Sheet", sheet), 
         x = "Index", 
         y = "Intensity") +
    scale_y_continuous(trans = scales::log10_trans()) +  # Log scale for y-axis
    theme_linedraw() +  # Use a clean theme
    theme(plot.title = element_text(size = 30),
          plot.subtitle = element_text(size = 25),
          axis.title.x = element_text(size = 25),
          axis.title.y = element_text(size = 25),
          axis.text = element_text(size = 20),      
          legend.key.size = unit(1.5, 'cm'),
          legend.text = element_text(size = 18), 
          legend.title = element_text(size = 18),
          panel.grid.major = element_line(color = "grey65", size = 1),
          panel.grid.minor = element_line(color = "grey80", size = 0.5))
  
  # Conditionally set ylim if scatter_plot_ymax is not NA
  if (!is.na(scatter_plot_ymax)) {
    p <- p + ylim(c(NA, scatter_plot_ymax))
  }
  
  # Save the plot using ggsave
  ggsave(filename = plot_idata, plot = p, width = 11, height = 12)
  
  # Close the current graphics device if it is open
  while (!is.null(dev.list())) dev.off()
  
  ### PART 4 - Applying Thresholds to Data and defining logical vectors ###
  thresholds <- c(DNA_thres_a, tract_thres_b, prot1_thres_c, prot2_thres_d)
  
  threshold_results <- data.frame(Map(function(column, thres) {
    ifelse(is.na(column), FALSE, column > thres)
  }, data, thresholds))
  
  colnames(threshold_results) <- paste0(colnames(data), "_thresh (>", thresholds, ")")
  
  
  ### Part 5- Smooth the gaps! This turns gaps of 'falses' to 'trues' based off user-defined values ###
  
  smooth_threshold <- function(column, smooth) {
    rle_data <- rle(column)
    lengths <- rle_data$lengths
    values <- rle_data$values
    
    # Identify where the lengths of consecutive FALSEs are less than or equal to the smoothing parameter
    false_gaps <- which(values == FALSE & lengths <= smooth)
    
    for (i in false_gaps) {
      # Check if the FALSEs are sandwiched between TRUEs
      if (i > 1 && i < length(values) && values[i - 1] == TRUE && values[i + 1] == TRUE) {
        values[i] <- TRUE
      }
    }
    
    # Reconstruct the smoothed column
    inverse.rle(list(lengths = lengths, values = values))
  }
  
  # Assign smoothing values to each column
  smoothed <- c(Smooth_it_1, Smooth_it_2, Smooth_it_3, Smooth_it_4)
  
  # Apply the smoothing function to each column with its corresponding smoothing value
  smoothed_results <- data.frame(Map(function(column, smooth) {
    smooth_threshold(column, smooth)
  }, threshold_results, smoothed))
  
  
  # Add names to the smoothed results
  colnames(smoothed_results) <- paste0(colnames(threshold_results), "_smoothed")
  
  
  ### PART 6- Count the number of "True" vectors, this gives us the length of tracts ###
  count_consecutive_trues <- function(vec) {
    count <- 0
    result <- vector("numeric", length(vec))
    for (i in seq_along(vec)) {
      if (isTRUE(vec[i])) {
        count <- count + 1
        result[i] <- count
      } else {
        count <- 0
        result[i] <- NA
      }
    }
    result
  }
  
  threshold_counts <- as.data.frame(lapply(smoothed_results, count_consecutive_trues))
  
  
  ### PART 7: Calculate average of longest consecutive trues and get all lengths ###
  calc_avg_longest_consec <- function(vec) {
    rle_vec <- rle(vec)
    longest_lengths <- rle_vec$lengths[rle_vec$values == TRUE]
    if (length(longest_lengths) == 0) {
      return(list(avg = NA, lengths = NA))
    } else {
      return(list(avg = mean(longest_lengths), lengths = longest_lengths))
    }
  }
  
  # Calculate the longest lengths for each thresholded column
  result_list <- lapply(smoothed_results, calc_avg_longest_consec)
  avg_longest_consec <- sapply(result_list, `[[`, "avg")
  longest_lengths <- lapply(result_list, `[[`, "lengths")
  
  avg_longest_consec_df <- data.frame(Channel = names(avg_longest_consec), AvgLongestConsec = avg_longest_consec)
  
  # Get the maximum length of all lengths
  max_length <- max(sapply(longest_lengths, length))
  longest_lengths_padded <- lapply(longest_lengths, function(x) {
    length(x) <- max_length
    return(x)
  })
  
  longest_lengths_df <- as.data.frame(do.call(cbind, longest_lengths_padded))
  colnames(longest_lengths_df) <- paste0(names(threshold_results), "_all_lengths")
  
  # Store the longest_lengths_df in the list
  all_longest_lengths_col1[[sheet]] <- longest_lengths_df[,1, drop = FALSE]
  all_longest_lengths_col2[[sheet]] <- longest_lengths_df[,2, drop = FALSE]
  all_longest_lengths_col3[[sheet]] <- longest_lengths_df[,3, drop = FALSE]
  all_longest_lengths_col4[[sheet]] <- longest_lengths_df[,4, drop = FALSE]
  
  
  ### PART 8 - Determining fork ends and middle, as well as unreplicated regions ###
  label_fork_ends <- function(column, unrep_fork, rep_fork) {
    column <- as.logical(column)
    result <- rep(NA, length(column))  # Initialize with NAs
    
    #print(paste("Initial column:", toString(column)))  # Debugging: print the input column
    #Check the class of the input column
    #print(class(column))
    
    
    i <- 1
    while (i <= length(column)) {
      if (isTRUE(column[i])) {
        # Look for the beginning of consecutive TRUEs
        start_true <- i
        while (i <= length(column) && isTRUE(column[i])) {
          i <- i + 1
        }
        end_true <- i - 1  # Last TRUE in the sequence
        
        
        ##print(paste("Start TRUE at:", start_true))  # Debugging
        ##print(paste("End TRUE at:", end_true))  # Debugging
        
        
        # Label the first TRUE as a FORK
        result[start_true] <- "FORK"
        
        # Label the last TRUE as a FORK
        if (start_true != end_true) {
          result[end_true] <- "FORK"
        }
        
        # Apply 'rep_fork' labeling
        if (rep_fork > 0 && (end_true - start_true + 1) > 1) {
          for (j in (start_true + 1):(start_true + rep_fork)) {
            if (j < end_true && isTRUE(column[j]) && is.na(result[j])) {
              result[j] <- "R FORK"
            }
          }
          for (j in (end_true - rep_fork):(end_true - 1)) {
            if (j > start_true && isTRUE(column[j]) && is.na(result[j])) {
              result[j] <- "R FORK"
            }
          }
        }
        
        # Apply 'unrep_fork' labeling
        if (unrep_fork > 0) {
          for (j in (start_true - unrep_fork):(start_true - 1)) {
            if (j > 0 && !isTRUE(column[j]) && is.na(result[j])) {
              result[j] <- "UNR FORK"
            }
          }
          for (j in (end_true + 1):(end_true + unrep_fork)) {
            if (j <= length(column) && !isTRUE(column[j]) && is.na(result[j])) {
              result[j] <- "UNR FORK"
            }
          }
        }
        
        # Label the remaining TRUEs in the middle as 'TRACT' (only if there are more than 2 TRUEs)
        if (end_true - start_true > (2* rep_fork)) {
          for (j in (start_true + 1):(end_true - 1)) {
            if (is.na(result[j])) {
              result[j] <- "TRACT"
            }
          }
        }
      } else {
        i <- i + 1
      }
    }
    return(result)
    
    #print(length(result)) #debugging
  }
  
  # Function to update column 2 to "." where column 1 is FALSE
  update_col2_based_on_col1 <- function(smoothed_results) {
    col2_updated <- smoothed_results[, 2]  # Copy column 2
    col2_updated[!smoothed_results[, 1]] <- "."  # Set to "." where column 1 is FALSE
    return(col2_updated)
  }
  
  # Function to replace TRUE and FALSE with NA (add NA in blank spaces)
  remove_TF <- function(column) {
    # Replace TRUE and FALSE with NA
    column[column == TRUE | column == FALSE | column == " "] <- "NA"
    return(column)
  }
  
  # Applying the multiple Part 8 functions
  
  # Applying the update_col2_based_on_col1 function to label col 2 ".", when col 1 is FALSE
  updated_col2 <- update_col2_based_on_col1(smoothed_results)
  
  
  # Create the result data frame with updated col2
  fork_temp <- smoothed_results
  
  
  # Apply the labeling to the updated column 2 and update fork_temp accordingly
  labeled_col2 <- label_fork_ends(updated_col2, unrep_fork, rep_fork)
  
  fork_temp[, 2] <- labeled_col2
  
  #Remove T/F from updated_col2 and create a new (mostly) empty df
  updated_col2_no_TF <-remove_TF(updated_col2)
  
  # Create a new dataframe with four columns where only the second column is updated_col2_no_TF
  col2.no_TF <- data.frame(
    V1 = NA,
    V2 = updated_col2_no_TF,
    V3 = NA,
    V4 = NA
  )
  
  # Add temp. names to match column names for merging
  colnames(col2.no_TF) <- paste0(colnames(threshold_results), "_smoothed")
  
  # Create the final merged result
  fork_result <- fork_temp  
  
  # Rename the columns to match the original intention
  colnames(fork_temp) <- paste0(colnames(threshold_results), "_smoothed")
  
  #Merge fork_temp and updated_col2.no_TF
  
  # Function to ensure col2.no_TF has the same number of rows as fork_temp
  ensure_row_count <- function(df1, df2) {
    if (nrow(df1) != nrow(df2)) {
      stop("The number of rows in the data frames do not match.")
    }
  }
  
  # Check if col2.no_TF is empty
  if (nrow(col2.no_TF) == 0) {
    warning("col2.no_TF is empty. The fork_temp data frame will not be modified for the second column.")
    
    # If col2.no_TF is empty, just use fork_temp as fork_result
    fork_result <- fork_temp
    
  } else {
    # Ensure col2.no_TF has the same number of rows as fork_temp
    ensure_row_count(col2.no_TF, fork_temp)
    
    # Convert the second column of col2.no_TF to character
    col2.no_TF <- as.data.frame(col2.no_TF)
    col2.no_TF[, 2] <- as.character(col2.no_TF[, 2])
    
    # Check if the conversion worked
    if (!is.character(col2.no_TF[, 2])) {
      stop("Column 2 in col2.no_TF should be a character vector.")
    }
  }
  
  # Ensure column names are set appropriately
  colnames(fork_result) <- colnames(fork_temp)
  
  # Update the second column in fork_result with the values from col2.no_TF wherever it is "."
  fork_result[, 2] <- ifelse(col2.no_TF[, 2] == ".", col2.no_TF[, 2], fork_temp[, 2])
  
  #Change blank <NA> into "|"
  col2_name <- colnames(fork_result)[2]
  fork_result[[col2_name]][is.na(fork_result[[col2_name]])] <- "|"
  
  
  ### PART 9- Determining protein co-localization with forks and tracts! ###
  
  coloc_protein <- function(protein_channel, col2, col1) {
    if (is.na(protein_channel) || is.na(col2)) {
      return(NA)  # Handle NA values explicitly
    } else if (protein_channel > 0 & col1 == TRUE & col2 == "UNR FORK") { 
      return("COLO UNR FORK") 
    } else if (protein_channel > 0 & col1 == TRUE & col2 == "R FORK") {
      return("COLO R FORK")
    } else if (protein_channel > 0 & col1 == TRUE & col2 == "FORK") {
      return("COLO FORK")
    } else if (protein_channel > 0 & col1 == TRUE & col2 == "TRACT") {
      return("COLO TRACT")
    } else if (protein_channel > 0 & col1 == TRUE & col2 == "|") {
      return("COLO UNR REGION")
    } else {
      return(NA)  # Or some other default value
    }
  }
  
  # Applying coloc_protein function with the updated conditions
  coloc_p1 <- mapply(function(protein_channel, col2, col1) coloc_protein(protein_channel, col2, col1), 
                     fork_result[, 3], fork_result[, 2], fork_result[, 1])
  coloc_p2 <- mapply(function(protein_channel, col2, col1) coloc_protein(protein_channel, col2, col1), 
                     fork_result[, 4], fork_result[, 2], fork_result[, 1])
  
  # Adding the colocalization results to the fork_result data frame
  fork_result[, 3] <- coloc_p1
  fork_result[, 4] <- coloc_p2
  
  
  ###Part 10- Counting the amount of co-localization of protein 
  
  #Function to count colocalization 
  count_coloc_occurrences <- function(column) {
    unrep_region_count <- sum(column == "COLO UNR REGION", na.rm = TRUE)
    tract_count <- sum(column == "COLO TRACT", na.rm = TRUE)
    fork_count <- sum(column %in% c("COLO UNR FORK", "COLO R FORK", "FORK"), na.rm = TRUE)
    c(COLO_UNR_REGION = unrep_region_count, COLO_TRACT = tract_count, FORKS = fork_count)
  }
  
  coloc_col3 <- count_coloc_occurrences(fork_result[[3]])
  coloc_col4 <- count_coloc_occurrences(fork_result[[4]])
  
  coloc_df <- data.frame(
    Region = c("COLO UNR REGION", "COLO TRACT", "FORKS"),
    Column3 = as.numeric(coloc_col3),
    Column4 = as.numeric(coloc_col4), 
    totalprotein3 = sum(as.numeric(coloc_col3), na.rm = TRUE),  # Sum values in coloc_col3
    totalprotein4 = sum(as.numeric(coloc_col4), na.rm = TRUE)  # Sum values in coloc_col4
  )
  
  # Function plotting bar graphs for colocalization counts
  coloc_plot <- function(coloc, column_number) {
    p <- ggplot(coloc, aes(x = Region, y = get(paste0("Column", column_number)), fill = Region)) +
      geom_bar(stat = "identity") +
      labs(title = paste("Colocalization of Proteins in Column", column_number, "for", sheet), x = "Region", y = "# of Protein") +
      theme_linedraw() +
      theme(plot.title = element_text(size = 30),
            plot.subtitle = element_text(size = 25),
            axis.title.x = element_text(size = 25),
            axis.title.y = element_text(size = 25),
            axis.text = element_text(size = 20),      
            legend.key.size = unit(1, 'cm'),
            legend.text = element_text(size = 18), 
            legend.title = element_text(size = 18),
            panel.grid.major = element_line(color = "grey65", size = 1),
            panel.grid.minor = element_line(color = "grey80", size = 0.5))
    
    # Conditionally set ylim if ind_coloc_p_ymax is not NA
    if (!is.na(ind_coloc_p_ymax)) {
      p <- p + ylim(c(NA, ind_coloc_p_ymax))
    }
    
    return(p)
  }
  
  # Applying coloc counts to individual sheets
  plot_col3 <- coloc_plot(coloc_df, 3)
  plot_col4 <- coloc_plot(coloc_df, 4)
  
  plot_coloc3 <- file.path(output_folder, paste0("bar_plot_coloc_col3_", sheet, ".png"))
  plot_coloc4 <- file.path(output_folder, paste0("bar_plot_coloc_col4_", sheet, ".png"))
  
  ggsave(filename = plot_coloc3, plot = plot_col3, width = 12, height = 6)
  ggsave(filename = plot_coloc4, plot = plot_col4, width = 12, height = 6)
  
  # Store the coloc_df in a list
  bin_col[[sheet]] <- coloc_df[drop = FALSE]
  
  
  ###PART 11 - Output and saving the work ###
  #For individual sheet
  sheet_name <- paste("Thres.Results -", sheet)
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet_name, threshold_results, startCol = 1, startRow = 1)
  writeData(wb, sheet_name, smoothed_results, startCol= 6, startRow= 1)
  writeData(wb, sheet_name, threshold_counts, startCol = 12, startRow = 1)
  writeData(wb, sheet_name, coloc_df, startCol= 18, startRow= 1)
  writeData(wb, sheet_name, fork_result, startCol = 18, startRow = 7) 
  writeData(wb, sheet_name, avg_longest_consec_df, startCol = 24, startRow = 1)
  writeData(wb, sheet_name, longest_lengths_df, startCol = 24, startRow = 7)
  insertImage(wb, sheet_name, plot_idata, startCol = 30, startRow = 2)
  insertImage(wb, sheet_name, plot_coloc3, startCol = 30, startRow = 16)
  insertImage(wb, sheet_name, plot_coloc4, startCol = 30, startRow = 28)
}  

# Output and saving the work
saveWorkbook(wb, outfile, overwrite = TRUE)


### PART 12- Adding summary plots ###
# Final summary workbook sheet
addWorksheet(wb, "Summary")

note_tukey <- "The results of the Tukey is printed below. This is also shown in the plots, where experiments with the same letter are not significant"
writeData(wb, "Summary", x = note_tukey, startCol = 1, startRow = 1)
note_tukey_eg <- "For example, for experiment/sheet 1 and 2, if both have the letter 'a' over it, this means that these are NOT signifcantly different"
writeData(wb, "Summary", x = note_tukey_eg, startCol = 1, startRow = 2)


###REF 10.5- Continuing binning coloc proteins for bar graph ###
#Binning function for coloc protein
binning_by_keywords <- function(data, bin_keywords, re_col_names) {
  # Remove any 'na' entries from the bin keywords list
  bin_keywords <- bin_keywords[bin_keywords != "na"]
  
  # Initialize an empty list to store binned data
  binned_data_list <- list()
  
  # Loop through each bin keyword
  for (bin_keyword in bin_keywords) {
    # Use grepl to find column names that match the bin keyword, allowing partial matches
    matching_cols <- names(data)[grepl(paste0(".*", bin_keyword, ".*"), names(data), ignore.case = TRUE)]
    
    # If there are matching columns, copy those to the binned data list
    if (length(matching_cols) > 0) {
      for (col in matching_cols) {
        # Create a temporary dataframe with the current column and a new column named after the bin keyword
        temp_data <- data.frame(Bin = bin_keyword, Value = data[[col]])
        
        # Add the temporary data frame to the binned data list
        binned_data_list[[length(binned_data_list) + 1]] <- temp_data
      }
    } else {
      print(paste("There were no matching keywords for:", bin_keyword))
    }
  }
  
  # Combine all binned data into a single dataframe
  if (length(binned_data_list) > 0) {
    binned_data <- bind_rows(binned_data_list)
    
    # Rename columns to desired names
    colnames(binned_data) <- re_col_names #this is a concatenated 'list' of your column names, this 'list' has to be the same number as columns in the data
    
    return(binned_data)
  } else {
    return(NULL)
  }
}

# Using the binning function for coloc
bin_keywords <- c(bin_1, bin_2, bin_3, bin_4, bin_5)
coloc_colname <- c("Bins", "Region", "Column3", "Column4", "totalprotein3", "totalprotein4") 

binned_data_coloc <- binning_by_keywords(bin_col, bin_keywords, coloc_colname)

# Calculate the proportion of each value for bin
cal_proportions <- function(binned_data, column_name, totalp_name, group, subgroup = NA) {
  propo_result <- binned_data %>%
    group_by(!!sym(group)) %>%
    mutate(Proportion = !!sym(column_name) / !!sym(totalp_name) * 100) %>%
    ungroup() %>%
    mutate(Proportion = replace_na(Proportion, 0)) # Replace NaN with 0
  
  # Debugging: Print result
  # print(propo_result)
  
  return(propo_result)
}
# Calculate mean and proportion for binning for bin
cal_mean <- function(binned_data, column_name, group, subgroup = NA) {
  result <- if (!is.na(subgroup)) {
    binned_data %>%
      group_by(!!sym(group), !!sym(subgroup)) %>%
      summarize(
        Mean = mean(.data[[column_name]], na.rm = TRUE),
        .groups = 'drop'  # This drops the grouping after summarizing
      )
  } else if (!is.na(group)) {
    binned_data %>%
      group_by(!!sym(group)) %>%
      summarize(
        Mean = mean(.data[[column_name]], na.rm = TRUE),
        .groups = 'drop'  # This drops the grouping after summarizing
      )
  } else {
    binned_data %>%
      summarize(
        Mean = mean(.data[[column_name]], na.rm = TRUE)
      )
  }
  
  # Calculate the proportion of each value to 100% per bin
  result <- result %>%
    group_by(!!sym(group)) %>%
    mutate(Proportion = Mean / sum(Mean) * 100) %>%
    ungroup() %>%
    

  # Debugging: Print result
  # print(result)
  
  return(result)
}


# Define the group and subgroup column names as strings for coloc proteins
group <- "Bins"
subgroup <- "Region"

# Apply the proportion to the coloc binned data
propo_data_col3 <- cal_proportions(binned_data_coloc, "Column3", "totalprotein3", group, subgroup)
propo_data_col4 <- cal_proportions(binned_data_coloc, "Column4", "totalprotein4", group, subgroup)

#Apply mean and proportion
mean_data_col3 <- cal_mean(binned_data_coloc, "Column3", group, subgroup)
mean_data_col4 <- cal_mean(binned_data_coloc, "Column4", group, subgroup)

# Function to perform F-test on proportion data
f_test <- function(proportion_data, group, subgroup) {
  # Convert strings to symbols
  group_col <- sym(group)
  subgroup_col <- sym(subgroup)
  
  # Get unique subgroups
  unique_subgroups <- proportion_data %>% pull(!!subgroup_col) %>% unique()
  
  # Initialize a list to store results
  f_test_results <- list()
  
  # Perform F-test for each unique subgroup
  for (subgroup_value in unique_subgroups) {
    # Filter data for the current subgroup
    subgroup_data <- proportion_data %>% filter(!!subgroup_col == subgroup_value)
    
    # Check if there are enough data points for F-test
    if (nrow(subgroup_data) < 2) {
      warning(paste("Not enough data points for subgroup:", subgroup_value))
      next
    }
    
    # Get unique levels of the grouping variable
    group_levels <- unique(subgroup_data[[group]])
    
    # Check number of levels in the grouping variable
    if (length(group_levels) < 2) {
      warning(paste("Skipping F-test for subgroup:", subgroup_value, 
                    "- needs at least 2 levels in group:", group))
      next
    }
    
    # Generate all combinations of group levels
    group_combinations <- combn(group_levels, 2, simplify = FALSE)
    
    # Perform F-test for each combination
    for (combination in group_combinations) {
      subg1 <- combination[1]
      subg2 <- combination[2]
      
      # Filter data for the two levels
      test_data <- subgroup_data %>% filter(!!group_col %in% c(subg1, subg2))
      
      # Check if there are enough data points for both levels
      if (nrow(test_data %>% filter(!!group_col == subg1)) < 1 || 
          nrow(test_data %>% filter(!!group_col == subg2)) < 1) {
        warning(paste("Not enough data points for levels:", subg1, "and", subg2))
        next
      }
      
      # Perform F-test (var.test)
      f_test_result <- var.test(test_data$Proportion ~ test_data[[group]])
      
      # Store results in a data frame
      f_test_results <- rbind(f_test_results, data.frame(
        Subgroup = subgroup_value,
        subg1 = subg1,
        subg2 = subg2,
        F_value = f_test_result$statistic,
        p_value = f_test_result$p.value
      ))
    }
  }
  
  print(f_test_results) #Debugging
  
  return(f_test_results)
}

# Perform F-test on the sample data
f_test_col3 <- f_test(propo_data_col3, group, subgroup)
f_test_col4 <- f_test(propo_data_col4, group, subgroup)
  
#Creating f-test result label for the plot
generate_ft_labels <- function(ft_results) {
  # Debugging: Check the structure of the input
  print("Structure of f_test_results:")
  str(ft_results)
  
  # Check if the required columns exist in the results
  required_columns <- c("Subgroup", "F_value", "p_value")
  if (!all(required_columns %in% colnames(ft_results))) {
    stop("The input data frame does not contain required columns: Subgroup, F_value, p_value.")
  }
  
  # Create labels for the F-test results based on p-value conditions
  ft_labels <- ft_results %>%
    mutate(Label = case_when(
      p_value > 0.05 ~ "ns",                     # Not significant
      p_value <= 0.05 & p_value > 0.01 ~ "*",    # Significant at p < 0.05
      p_value <= 0.01 & p_value > 0.001 ~ "**",  # Significant at p < 0.01
      p_value <= 0.001 ~ "***"                  # Significant at p < 0.001
    ))
  
  # Debugging: Check the labels before selecting
  print("Labels before selecting:")
  print(ft_labels)
  
  # Selecting relevant columns using dplyr::select
  ft_labels <- ft_labels %>%
    dplyr::select(Subgroup, Label)
  
  # Print the final ft_labels data frame
  print("Final output:")
  print(ft_labels)
  
  return(ft_labels)
}

# Apply the function to the F-test results
ft_labels_col3 <- generate_ft_labels(f_test_col3)
ft_labels_col4 <- generate_ft_labels(f_test_col4)
  
# Function plotting grouped bar graphs with error bars for colocalization counts 
plot_binned_bar_tracts <- function(summary_data, column_number, f_test_label) {
  
  # Rename the column to ensure consistency
  colnames(summary_data)[which(colnames(summary_data) == "Regions")] <- "Region"
  
  # Convert Region column to a factor to ensure it has levels
  summary_data$Region <- as.factor(summary_data$Region)
  
  # Debugging code to check levels and subgroups
  #print("Levels of Region in summary_data:")
  #print(levels(summary_data$Region))
  #print("Subgroup values in f_test_label:")
  #print(f_test_label$Subgroup)
  
  # Ensure f_test_label Subgroup is a factor with the same levels as summary_data Region
  f_test_label$Subgroup <- factor(f_test_label$Subgroup, levels = levels(summary_data$Region))
  
  # Debugging: Check if the Subgroup values are correctly factored
  #print("Factored Subgroup values in f_test_label:")
  #print(f_test_label$Subgroup)
  
  # Determine the maximum y-value for positioning the text annotations
  max_y_value <- max(summary_data$Proportion, na.rm = TRUE)* 1.3
  
  # Start building the ggplot object without ylim
  p <- ggplot(summary_data, aes(x = Region, y = Proportion, fill = Bins)) +
    geom_bar(stat = "identity", position = position_dodge(), width = 0.7) +
    labs(title = paste("Percentage of Protein Colocalization in Column", column_number), 
         x = "Region", y = "% of Protein Colocalization") +
    theme_linedraw() +
    theme(plot.title = element_text(size = 28),
          plot.subtitle = element_text(size = 25),
          axis.title.x = element_text(size = 25),
          axis.title.y = element_text(size = 25),
          axis.text = element_text(size = 20),      
          legend.key.size = unit(1, 'cm'),
          legend.text = element_text(size = 18), 
          legend.title = element_text(size = 18),
          panel.grid.major = element_line(color = "grey65", size = 1),
          panel.grid.minor = element_line(color = "grey80", size = 0.5))
  
  # Conditionally set ylim if bin_coloc_p_ymax is not NA
  if (!is.na(bin_coloc_p_ymax)) {
    p <- p + ylim(c(NA, bin_coloc_p_ymax))
    max_label_y <- bin_coloc_p_ymax * 0.85  # Position label slightly below the max y-axis limit
  } else {
    p <- p + ylim(c(0, max_y_value))
    max_label_y <- max_y_value * 0.85  # Position label slightly below the calculated max y-axis limit
  }
  
  # Add the geom_label for the F-test label
  #p <- p + geom_label(data = f_test_label, aes(x = Subgroup, y = max_label_y, label = Label), 
  #                    vjust = -0.5, size = 8, inherit.aes = FALSE, show.legend = FALSE, fill = "white")
  
  # Return the final plot
  return(p)
}

# Plot the grouped bar graph for binned regions for protein colocalization
plot_col3_binned <- plot_binned_bar_tracts(mean_data_col3, 3, ft_labels_col3 )
plot_col4_binned <- plot_binned_bar_tracts(mean_data_col4, 4, ft_labels_col3 )

plot_coloc3_binned_path <- file.path(output_folder, "Binned_bar_plot_coloc_col3.png")
plot_coloc4_binned_path <- file.path(output_folder, "Binned_bar_plot_coloc_col4.png")

ggsave(filename = plot_coloc3_binned_path, plot = plot_col3_binned, width = 9, height = 6)
ggsave(filename = plot_coloc4_binned_path, plot = plot_col4_binned, width = 9, height = 6)

insertImage(wb, "Summary", plot_coloc3_binned_path, startCol = 16, startRow = 7)
insertImage(wb, "Summary", plot_coloc4_binned_path, startCol = 16, startRow = 17)

#write f-test results in the "Summary" sheet
notefcol3 <- "F-test results for column 3:"
writeData(wb, "Summary", x = notefcol3, startCol = 24, startRow = 8)
writeData(wb, "Summary", f_test_col3, startCol = 24, startRow = 7)

notefcol4 <- "F-test results for column 4:"
writeData(wb, "Summary", x = notefcol4, startCol = 30, startRow = 8)
writeData(wb, "Summary", f_test_col4, startCol = 30, startRow = 7)


### REF 7.5 - Continuing plotting violin and bar plots for tract lengths comparing all experiments ###

# Function to output violin plot with ANOVA & Tukey test and return the ggplot object
output_violin_plot <- function(all_longest_lengths_list, output_folder, column_number) {
  combined_data <- list()
  cat("combined_data printing", combined_data, "\n")
  
  # Combine data from all sheets
  for (sheet_name in names(all_longest_lengths_list)) {
    df <- all_longest_lengths_list[[sheet_name]]
    
    if (nrow(df) > 0 && any(!is.na(df[[1]]))) { # Adjusted for first column
      df$Sheet <- sheet_name
      combined_data[[sheet_name]] <- df
    } else {
      warning(paste("Data for sheet", sheet_name, "is empty or contains only NA values. Skipping."))
    }
  }
  
  combined_df <- bind_rows(combined_data)
  
  if (nrow(combined_df) <= 1) {
    warning("Combined data frame has 1 or fewer valid data points. Skipping plot.")
    return(NULL)
  }
  
  # Print summary for debugging
  print(summary(combined_df))
  print(head(combined_df))
  
  # Convert columns to appropriate types
  combined_df$Channel <- as.numeric(as.character(combined_df$Channel))
  combined_df$Sheet <- as.factor(combined_df$Sheet)
  
  # Filter out non-finite values, 0s, and NAs
  combined_df <- combined_df %>%
    filter(is.finite(Channel) & Channel > 0)
  
  # Perform ANOVA
  anova_result <- aov(Channel ~ Sheet, data = combined_df)
  anova_summary <- summary(anova_result)
  print(anova_summary)
  
  p_value <- round(anova_summary[[1]]$`Pr(>F)`[1], 5)
  
  # Checking TukeyHSD since ANOVA is significant
  gobblegobble <- suppressWarnings(TukeyHSD(anova_result, conf.level = 0.95, adjust.method = "holm"))
  print(gobblegobble)  # Print the entire output to inspect its structure
  
  # Rounding p-values in TukeyHSD results and ensure it's numeric
  tukey_out <- as.data.frame(gobblegobble$Sheet)
  tukey_out$`p adj` <- as.numeric(as.character(tukey_out$`p adj`))
  tukey_out$`p adj` <- round(tukey_out$`p adj`, 5)
  print(tukey_out)
  
  # Extract TukeyHSD results
  if ("Sheet" %in% names(gobblegobble)) {
    tukey_out <- as.data.frame(gobblegobble$Sheet)
  } else {
    warning("TukeyHSD output does not contain 'Sheet'. Please check the structure of the output.")
    tukey_out <- data.frame()  # Create an empty data frame if columns not found
  }
  
  # Extract TukeyHSD results for specific columns
  tukey_out <- as.data.frame(gobblegobble$Sheet)
  print(colnames(tukey_out)) # Print column names for debugging
  
  # Add row names as a new column for Sheet comparisons
  tukey_out$Sheet_Comparison <- rownames(tukey_out)
  
  # Select the correct columns based on actual column names
  if (all(c("diff", "p adj") %in% colnames(tukey_out))) {
    tukey_out <- tukey_out[, c("Sheet_Comparison", "diff", "p adj")]
  } else {
    warning("Expected columns 'diff' and 'p adj' not found in TukeyHSD results.")
    tukey_out <- data.frame() # Create an empty data frame if columns not found
  }
  
  writeData(wb, "Summary", tukey_out, startCol = (column_number - 1) * 5 , startRow = 37)
  notetresults <- "Tukey results for violin plot"
  writeData(wb, "Summary", x = notetresults, startCol = (column_number - 1) * 5 , startRow = 36)
  
  # Printing Tukey results onto the plot
  generate_tukey_labels <- function(tukey_results){
    # Extract label and factor levels 
    tukey_fac_levels <- cld(glht(anova_result, linfct = mcp(Sheet = "Tukey")), level = 0.95)
    tukey_labels <- data.frame(Letters = tukey_fac_levels$mcletters$Letters)
    
    # Label in the same order as the plot 
    tukey_labels$Sheet = rownames(tukey_labels)
    tukey_labels = tukey_labels[order(tukey_labels$Sheet) , ]
    return(tukey_labels)
  }
  
  # Apply function to data
  t_label <- generate_tukey_labels(gobblegobble)
  print(t_label)
  
  # Calculate the averages for each sheet
  averages <- combined_df %>%
    group_by(Sheet) %>%
    summarize(mean_length = mean(Channel, na.rm = TRUE), .groups = 'drop')
  
  # Calculate the max y value for setting the ylim
  max_y <- max(combined_df$Channel, na.rm = TRUE) * 1.3
  
  # Create the violin plot with average, ANOVA, and Tukey results
  p <- ggplot(combined_df, aes(x = Sheet, y = Channel, fill = Sheet)) +
    geom_violin(trim = FALSE) +
    geom_point(data = averages, aes(x = Sheet, y = mean_length), color = "#361d4a", size = 5, shape = 18) + # Adds the average points
    geom_label(data = averages, aes(x = Sheet, y = mean_length, label = round(mean_length, 1)), vjust = -1, size = 8, fill = "white", nudge_x = 0.25, nudge_y = -0.25) + # Adds average values labels
    labs(title = paste("Violin Plot of Lengths by Sheet - Column", column_number), 
         subtitle = paste("ANOVA p-value:", round(p_value, 5)), # p-value is rounded to 5 decimal places
         x = "Sheet names", 
         y = "Length") +
    theme_linedraw() +
    theme(plot.title = element_text(size = 30),
          plot.subtitle = element_text(size = 25),
          axis.title.x = element_text(size = 25),
          axis.title.y = element_text(size = 25),
          axis.text = element_text(size = 20),      
          legend.key.size = unit(1, 'cm'),
          legend.text = element_text(size = 18), 
          legend.title = element_text(size = 18),
          panel.grid.major = element_line(color = "grey65", size = 1),
          panel.grid.minor = element_line(color = "grey80", size = 0.5))
  
  # Conditionally set ylim if ind_tract_ymax is not NA
  if (!is.na(ind_tract_ymax)) {
    p <- p + ylim(c(NA, ind_tract_ymax))
    max_label_y <- ind_tract_ymax * 0.95  # Position label slightly below the max y-axis limit
  } else {
    p <- p + ylim(c(0, max_y))
    max_label_y <- max_y * 0.95  # Position label slightly below the calculated max y-axis limit
  }
  
  # Add the geom_label for t_label after setting ylim
  p <- p + geom_label(data = t_label, aes(x = Sheet, y = max_label_y, label = Letters), 
                      vjust = 0, size = 8, show.legend = FALSE, fill = "white")
  
  print(p)
  
  # Define the dynamic file path for saving the plot
  plot_file_path <- file.path(output_folder, paste0("violin_plot_column_", column_number, ".png"))
  
  # Save the plot
  tryCatch({
    ggsave(filename = plot_file_path, plot = p, width = 21, height = 10)
    
    # Attempt to insert image into the workbook
    cat("Attempting to insert image for column", column_number, "...\n")  # Debugging
    insertImage(wb, "Summary", plot_file_path, startCol = 1, startRow = (column_number + 1) * 4)
    
    cat("Inserted image for column", column_number, "at row", (column_number + 1) * 4 , "\n") # Debugging
  }, error = function(e) {
    warning(paste("Failed to save or insert plot:", e$message))
  })
  
  # Return the plot file path
  return(plot_file_path)
}

# Individual Plot violin plots for each column list
summary_v_plot_channel_1 <- output_violin_plot(all_longest_lengths_col1, output_folder, 1)
summary_v_plot_channel_2 <- output_violin_plot(all_longest_lengths_col2, output_folder, 2)
summary_v_plot_channel_3 <- output_violin_plot(all_longest_lengths_col3, output_folder, 3)
summary_v_plot_channel_4 <- output_violin_plot(all_longest_lengths_col4, output_folder, 4)


### REF 7.55 - More continuing plotting binned violin plots for tract lengths ###
#Using the binning function for tract lengths (function From REF 10.5)
#obtain associated channel name- since it's dynamic, we create a variable with these names and separate them
lengths_colname_list <- "lengths"
lengths_colname_1 <- c("Bins_mean", lengths_colname_list)
lengths_colname_2 <- c("Bins_mean", lengths_colname_list)
lengths_colname_3 <- c("Bins_mean", lengths_colname_list)
lengths_colname_4 <- c("Bins_mean", lengths_colname_list)

#Using the binning function
binned_data_lengths_1 <- binning_by_keywords(all_longest_lengths_col1, bin_keywords, lengths_colname_1)
binned_data_lengths_2 <- binning_by_keywords(all_longest_lengths_col2, bin_keywords, lengths_colname_2)
binned_data_lengths_3 <- binning_by_keywords(all_longest_lengths_col3, bin_keywords, lengths_colname_3)
binned_data_lengths_4 <- binning_by_keywords(all_longest_lengths_col4, bin_keywords, lengths_colname_4)

# Calculate mean, standard error for binned data
calculate_mean_STE <- function(binned_data, column_name, group) {
  # Ensure the group variable is a factor
  binned_data[[group]] <- as.factor(binned_data[[group]])
  
  # Grouping and calculating Mean and SE
  binned_data_summary <- binned_data %>%
    group_by(!!sym(group)) %>%
    summarize(
      Mean = mean(.data[[column_name]], na.rm = TRUE),
      SE = sd(.data[[column_name]], na.rm = TRUE) / sqrt(n()),
      .groups = 'drop'
    )
  
  print(str(binned_data_summary))
  
  return(binned_data_summary)
}

# Re-define the group and subgroup column names as strings for tract lengths (function from REF 10.5)
group <- "Bins_mean"

# Using the Mean and STE function to the tract lengths binned data (function from REF 10.5)
summary_data_lengths_1 <- calculate_mean_STE(binned_data_lengths_1, lengths_colname_list, group)
summary_data_lengths_2 <- calculate_mean_STE(binned_data_lengths_2, lengths_colname_list, group)
summary_data_lengths_3 <- calculate_mean_STE(binned_data_lengths_3, lengths_colname_list, group)
summary_data_lengths_4 <- calculate_mean_STE(binned_data_lengths_4, lengths_colname_list, group)

# Altered One-way ANOVA Test and bar Plot for BINNED lengths
anova_and_plot <- function(bin_data_lengths, sum_data_lengths, output_folder, column_number) {
  
  # Check for NA values in the second column (the DNA lengths)
  if (any(is.na(bin_data_lengths[[2]]))) {
    bin_data_lengths <- bin_data_lengths[!is.na(bin_data_lengths[[2]]), ] 
    print("Removed NA from the second column")
  }
  
  # Get the number of rows in each data frame
  rows_bin <- nrow(bin_data_lengths)
  rows_sum <- nrow(sum_data_lengths)
  
  # Calculate the maximum length of rows
  max_length <- max(rows_bin, rows_sum)
  
  # Adjust the lengths of the data frames to the maximum length so that you can put bin_data_lengths and sum_data_lengths in one df
  
  # Check and adjust bin_data_lengths
  if (rows_bin < max_length) {
    bin_data_lengths <- rbind(bin_data_lengths, data.frame(Bins = rep(NA, max_length - rows_bin), 
                                                           Length = rep(NA, max_length - rows_bin)))
    print("Extended bin_data_lengths with NA rows")
  } else if (rows_bin == max_length) {
    # Do nothing and move to the next part of the script
    print("No extension needed for bin_data_lengths")
  }
  
  # Check and adjust sum_data_lengths
  if (rows_sum < max_length) {
    sum_data_lengths_wNA <- rbind(sum_data_lengths, data.frame(Bins_mean = rep(NA, max_length - rows_sum), 
                                                               Mean = rep(NA, max_length - rows_sum), 
                                                               SE = rep(NA, max_length - rows_sum)))
    print("Extended sum_data_lengths with NA rows")
  } else if (rows_sum == max_length) {
    sum_data_lengths_wNA <- sum_data_lengths  # No need to extend, just assign
    print("No extension needed for sum_data_lengths")
  }
  
  # Combine the two data frames
  bin_data_lengths <- cbind(bin_data_lengths, sum_data_lengths_wNA)
  
  # Rename the first column of bin_data_lengths to "Bins"
  colnames(bin_data_lengths)[1] <- "Bins"
  
  # Replace NA values in the Mean column with 0 (or any other value you prefer)
  bin_data_lengths <- bin_data_lengths %>%
    mutate(Mean = ifelse(is.na(Mean), 0, Mean))  # Replace NA with 0 or any other value you choose
  
  # Debugging: Print column names and first few rows of the data and dimension 
  #print(dim(bin_data_lengths))
  #print(names(bin_data_lengths))
  #print((bin_data_lengths))
  
  # Check if Bins column is a factor
  if (!is.factor(bin_data_lengths$Bins)) {
    bin_data_lengths$Bins <- as.factor(bin_data_lengths$Bins)
  }
  
  # Ensure Bins is a factor with levels based on unique values
  bin_data_lengths$Bins <- factor(bin_data_lengths$Bins, levels = unique(bin_data_lengths$Bins))
  
  # Rename the second column to "lengths" for anova
  colnames(bin_data_lengths)[2] <- "lengths"
  
  # Perform One-way ANOVA using the renamed column
  anova_test_result <- aov(lengths ~ Bins, data = bin_data_lengths)
  anova_summary <- summary(anova_test_result)
  print("One-way ANOVA Test Result:")
  print(anova_summary)
  
  # Extract p-value from ANOVA summary
  p_value <- anova_summary[[1]][["Pr(>F)"]][1]
  print(paste("Extracted p-value:", p_value))
  
  # Check if p_value is numeric and handle NA or non-numeric cases
  if (is.na(p_value) || !is.numeric(p_value)) {
    p_value <- "NA"
    print("No p-value from ANOVA test")
  } else {
    p_value <- round(p_value, 5)
    print("p-value from ANOVA test obtained")
  }
  
  # Checking TukeyHSD since ANOVA is significant
  tuk_test <- suppressWarnings(TukeyHSD(anova_test_result, conf.level = 0.95)) 
  print(tuk_test)
  plot(tuk_test)
  
  # Rounding p-values in TukeyHSD results and ensure it's numeric
  tukey_out <- as.data.frame(tuk_test$Bins)
  tukey_out$`p adj` <- as.numeric(as.character(tukey_out$`p adj`))
  tukey_out$`p adj` <- round(tukey_out$`p adj`, 5)
  print(tukey_out)
  
  # Extract TukeyHSD results
  if ("Bins" %in% names(tuk_test)) {
    tukey_out <- as.data.frame(tuk_test$Bins)
  } else {
    warning("TukeyHSD output does not contain 'Bins'. Please check the structure of the output.")
    tukey_out <- data.frame()  # Create an empty data frame if columns not found
  }
  
  # Extract TukeyHSD results for specific columns
  tukey_out <- as.data.frame(tuk_test$Bins)
  print(colnames(tukey_out)) # Print column names for debugging
  
  # Add row names as a new column for Sheet comparisons
  tukey_out$Sheet_Comparison <- rownames(tukey_out)
  
  # Select the correct columns based on actual column names
  if (all(c("diff", "p adj") %in% colnames(tukey_out))) {
    # Reference the first column (Sheet_Comparison) for mrc-wt, cds-wt, etc.
    tukey_out <- tukey_out[, c("Sheet_Comparison", "diff", "p adj")]
  } else {
    warning("Expected columns 'diff' and 'p adj' not found in TukeyHSD results.")
    tukey_out <- data.frame() # Create an empty data frame if columns not found
  }
  
  # Rounding p-values in TukeyHSD results and ensure it's numeric
  tukey_out$`p adj` <- as.numeric(as.character(tukey_out$`p adj`))
  tukey_out$`p adj` <- round(tukey_out$`p adj`, 5)
  
  writeData(wb, "Summary", tukey_out, startCol = (column_number - 1) * 5 , startRow = 81)
  notetresults <- "Tukey results for binned bar plot"
  writeData(wb, "Summary", x = notetresults, startCol = (column_number - 1) * 5 , startRow = 80)
  
  # Printing tukey results on to the plot (from og. output_violin_plot function)
  generate_tukey_labels2 <- function(tukey_results, variable) {
    # Extract label and factor levels 
    tukey_fac_levels <- cld(glht(anova_test_result, linfct = mcp(Bins = "Tukey")), level = 0.95)
    tukey_labels <- data.frame(Letters = tukey_fac_levels$mcletters$Letters)
    
    # Reference the first column of the Tukey output as labels
    tukey_labels$Bins_mean = rownames(tukey_labels)
    tukey_labels = tukey_labels[order(tukey_labels$Bins_mean), ]
    
    return(tukey_labels)
  }
  
  # Apply function to data
  t_label <- generate_tukey_labels2(tuk_test)
  print(names(t_label))
  print("Tukey test results")
  print(t_label)
  
  # DEBUGGING
  # Debugging: Print column names and first few rows of the data and dimension 
  # print("Dimensions of bin_data_lengths:")
  # print(dim(bin_data_lengths))
  # print("Column names of bin_data_lengths:")
  # print(names(bin_data_lengths))
  # print("First few rows of bin_data_lengths:")
  # print(head(bin_data_lengths))
  
  # Debugging: Print the structure of sum_data_lengths (for plotting)
  # print("Structure of sum_data_lengths:")
  # print(str(sum_data_lengths))
  #  print(colnames(sum_data_lengths))
  #  print(head(sum_data_lengths$Mean))
  #  print("First few rows of sum_data_lengths:")
  #  print(head(sum_data_lengths))
  
  # Calculate the maximum y value for setting the ylim
  max_y <- max(sum_data_lengths$Mean + sum_data_lengths$SE, na.rm = TRUE) * 1.3
  
  # Create the bar graph with mean and SE bars
  p <- ggplot(sum_data_lengths, aes(x = Bins_mean, y = Mean, fill = Bins_mean)) +
    geom_bar(stat = "identity", position = "dodge") +
    geom_errorbar(aes(ymin = Mean - SE, ymax = Mean + SE), 
                  width = 0.5, size = 1.25, position = position_dodge(1.0), color = "#545454") +
    geom_label(aes(label = round(Mean, 1)), vjust = -1, size = 8, fill = "white", nudge_x = 0.25, nudge_y = -0.25) +
    labs(title = paste("Mean Lengths for column", column_number),
         subtitle = paste("One-way ANOVA p-value:", p_value),
         x = "Bins",
         y = "Mean Length (px)") +
    theme_linedraw() +
    theme(plot.title = element_text(size = 30),
          plot.subtitle = element_text(size = 25),
          axis.title.x = element_text(size = 25),
          axis.title.y = element_text(size = 25),
          axis.text = element_text(size = 20),      
          legend.key.size = unit(1, 'cm'),
          legend.text = element_text(size = 18), 
          legend.title = element_text(size = 18),
          panel.grid.major = element_line(color = "grey65", size = 1),
          panel.grid.minor = element_line(color = "grey80", size = 0.5)) 
    
  # Conditionally set ylim if bin_tract_ymax is not NA
  if (!is.na(bin_tract_ymax)) {
    p <- p + ylim(c(0, bin_tract_ymax))
    max_label_y <- bin_tract_ymax * 0.95  # Position label slightly below the max y-axis limit
  } else {
    p <- p + ylim(c(0, max_y))
    max_label_y <- max_y * 0.95  # Position label slightly below the calculated max y-axis limit
  }
  
  # Add the geom_label for t_label after setting ylim
  p <- p + geom_label(data = t_label, aes(x = Bins_mean, y = max_label_y, label = Letters), 
                      vjust = 0, size = 8, fill = "white", show.legend = FALSE)
  
  print(p)
  
  # Save the plot
  plot_file_path <- file.path(output_folder, paste0("Binned_DNA_lengths_for_column_", column_number, ".png"))
  ggsave(filename = plot_file_path, plot = p, width = 9, height = 6)
  
  insertImage(wb, "Summary", plot_file_path, startCol = 8, startRow = (column_number + 1) * 4)
  
  # Return the plot file path
  return(plot_file_path)
}

#Using the function
binned_summary_v_plot_channel_1 <- anova_and_plot(binned_data_lengths_1, summary_data_lengths_1, output_folder, 1)
binned_summary_v_plot_channel_2 <- anova_and_plot(binned_data_lengths_2, summary_data_lengths_2, output_folder, 2)
binned_summary_v_plot_channel_3 <- anova_and_plot(binned_data_lengths_3, summary_data_lengths_3, output_folder, 3)
binned_summary_v_plot_channel_4 <- anova_and_plot(binned_data_lengths_4, summary_data_lengths_4, output_folder, 4)

# Save the workbook
saveWorkbook(wb, outfile, overwrite = TRUE)
cat("Parameters and threshold results saved to", outfile)

# Clear Environment for a fresh start on analysis!
rm(list = ls())