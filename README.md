# Platelet Count and Survival Time Analysis in Canine Cancer Patients

This repository contains an R script for performing a statistical analysis of platelet counts and survival time in cancer patients. The script uses Mann-Whitney tests to compare platelet counts between groups of patients with different survival times.

## Objective

The goal of this code is to evaluate whether there is a significant difference in platelet counts of cancer patients based on their survival time. The code applies statistical tests to assess this relationship over various survival time cutoff points ranging from 4 to 24 months.

## Features

- Reads data from an Excel file containing information about patients.
- Applies Mann-Whitney tests to compare groups of patients based on different survival time cutoff points.
- Calculates descriptive statistics for each group, such as mean, standard deviation, median, and quartiles.
- Generates an Excel file with the results of the analysis, where each filter is stored in a separate sheet.

## How to Use

1. Clone this repository:
    ```bash
    git clone https://github.com/your-username/platelet-survival-analysis.git
    ```

2. Install the required dependencies. The code is written in R and requires the following packages:
    ```R
    install.packages("readxl")
    install.packages("openxlsx")
    ```

3. Modify the file path to point to your Excel data file:
    ```r
    file_path <- "path/to/your/file.xlsx"
    ```

4. Run the script:
    ```R
    source("platelet-survival-analysis.R")
    ```

5. The script will generate an Excel file with the results, saved as `Platelet_Analysis_Results.xlsx`.

## Inputs

The expected input file should be an Excel file with the following fields:

- **Platelet count (/ul)**: Platelet count per microliter.
- **Survival time (month)**: Survival time in months.
- **Diagnosis**: Diagnosis of the patients.
- **Histogenesis**: Cancer histogenesis type.
- **Malignancy**: Type of malignancy.

## Outputs

The script generates an Excel file containing a sheet for each filter, with the following results for each survival time cutoff point:

- **Cutoff**: The survival time cutoff point (in months).
- **Category**: The category of the filter (e.g., diagnosis type, histogenesis, etc.).
- **Mann-Whitney Statistic**: The Mann-Whitney test statistic.
- **P-Value**: The p-value of the Mann-Whitney test.
- **N_Less_Equal**: Number of patients in the group with survival time less than or equal to the cutoff point.
- **N_Greater**: Number of patients in the group with survival time greater than the cutoff point.
- **Descriptive Statistics**: Mean, standard deviation, median, and quartiles (Q1, Q3) for each group.




