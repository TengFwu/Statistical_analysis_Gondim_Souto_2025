
# Packages
if (!require("readxl")) install.packages("readxl")
library(readxl)

if (!require("tidyverse")) install.packages("tidyverse")
library(tidyverse)

if (!require("survival")) install.packages("survival")
library(survival)

if (!require("survminer")) install.packages("survminer")
library(survminer)

if (!require("psych")) install.packages("psych")
library(psych)


# Dataset

file <- "Supplementary_data.xlsx"

data <- read_xlsx(file)

# Classify thrombocytosis intensity

thrombo_class <- data %>%
  mutate(
    Intensity = ifelse(`Platelet count (/ul)` <= 600000, "Mild", "Moderate-severe"),
    Intensity = factor(Intensity, levels = c("Mild", "Moderate-severe")),
    meta_status = case_when( #Not performed will be NA values
      Metastasis == "Detected" ~ "Detected",
      Metastasis == "No evidence" ~ "No evidence" 
    ),
    meta_status = factor(meta_status, levels = c("No evidence", "Detected"))
  )

thrombo_class <- thrombo_class %>% 
  filter(Malignancy == "Malignant")


thrombo_class$Censure <- as.factor(thrombo_class$Censure)
thrombo_class$`Survival time (month)` <- as.numeric(thrombo_class$`Survival time (month)`)
thrombo_class$Histogenesis <- as.factor(thrombo_class$Histogenesis)
thrombo_class$meta_status <- as.factor(thrombo_class$meta_status)
thrombo_class$Malignancy <- as.factor(thrombo_class$Malignancy)


# Kaplan-meier
fit_thromb <- survfit(Surv(time = `Survival time (month)`, event = Censure == 1) ~ Intensity, 
                     data = thrombo_class)

summary(fit_thromb)$table



ggsurvplot(
  fit_thromb,
  data = thrombo_class,
  pval = TRUE,
  conf.int = FALSE,
  risk.table = TRUE,
  legend.title = "Intensity of thrombocytosis",   
  legend.labs = c("Mild", "Moderate-Severe"),
  palette = c("#1b9e77", "#d95f02"),
  xlab = "Survival time (months)",
  ylab = "Survival probability",
  font.x = 16,        
  font.y = 16,        
  font.legend = 16,   
  font.tickslab = 16  
)

# Multivariate Cox model
modelo_cox_multi <- coxph(
Surv(time = `Survival time (month)`, event = Censure == 1) ~ Intensity + Histogenesis + meta_status,
  data = thrombo_class
)

summary(modelo_cox_multi)



##### Same approach only with histopathology exams #####
thrombo_class_histo <- thrombo_class %>% 
  filter(`Cito/Histo` == "Histopathology")



# Kaplan-meier
fit_thromb_histo <- survfit(Surv(time = `Survival time (month)`, event = Censure == 1) ~ Intensity, 
                      data = thrombo_class_histo)

summary(fit_thromb_histo)$table

ggsurvplot(
  fit_thromb_histo,
  data = thrombo_class_histo,
  pval = TRUE,
  conf.int = FALSE,
  risk.table = TRUE,
  legend.title = "Intensity of thrombocytosis",   
  legend.labs = c("Mild", "Moderate-Severe"),
  palette = c("#1b9e77", "#d95f02"),
  xlab = "Survival time (months)",
  ylab = "Survival probability",
  font.x = 16,        
  font.y = 16,        
  font.legend = 16,  
  font.tickslab = 16 
)


# Multivariate Cox model
modelo_cox_multi <- coxph(
  Surv(time = `Survival time (month)`, event = Censure == 1) ~ Intensity +  Histogenesis + meta_status,
  data = thrombo_class_histo
)

summary(modelo_cox_multi)



### Analysis of the most prevalent diagnosis

# 1. Mammary Carcinoma 

thrombo_class_neo <- thrombo_class 


thrombo_class_mama <- thrombo_class_neo %>% 
  filter(Diagnosis == "Mammary carcinoma")

# If you want to analyse only the histopathological diagnosis
# thrombo_class_mama <- thrombo_class_mama %>% filter(`Cito/Histo` == "Histopathology") 

# Kaplan-meier
fit_thromb_mama <- survfit(Surv(time = `Survival time (month)`, event = Censure == 1) ~ Intensity, 
                            data = thrombo_class_mama)


summary(fit_thromb_mama)$table

ggsurvplot(
  fit_thromb_mama,
  data = thrombo_class_mama,
  pval = TRUE,
  conf.int = FALSE,
  risk.table = TRUE,
  legend.title = "Intensity of thrombocytosis",   
  legend.labs = c("Mild", "Moderate-Severe"),
  palette = c("#1b9e77", "#d95f02"),
  xlab = "Survival time (months)",
  ylab = "Survival probability",
  font.x = 16,        
  font.y = 16,        
  font.legend = 16,   
  font.tickslab = 16  
)


# Multivariate Cox model
modelo_cox_mama <- coxph(
  Surv(time = `Survival time (month)`, event = Censure == 1) ~ Intensity + meta_status,
  data = thrombo_class_mama
)

summary(modelo_cox_mama)


# 2. Squamous cell carcinoma
thrombo_class_scc <- thrombo_class_neo %>% 
  filter(Diagnosis == "Squamous cell carcinoma")

# Kaplan-meier
fit_thromb_scc <- survfit(Surv(time = `Survival time (month)`, event = Censure == 1) ~ Intensity, 
                           data = thrombo_class_scc)


ggsurvplot(
  fit_thromb_scc,
  data = thrombo_class_scc,
  pval = TRUE,
  conf.int = FALSE,
  risk.table = TRUE,
  legend.title = "Intensity of thrombocytosis",   
  legend.labs = c("Mild", "Moderate"),
  palette = c("#1b9e77", "#d95f02"),
  xlab = "Survival time (months)",
  ylab = "Survival probability"
)


# Multivariate Cox model
modelo_cox_scc <- coxph(
  Surv(time = `Survival time (month)`, event = Censure == 1) ~ Intensity + meta_status,
  data = thrombo_class_scc
)

summary(modelo_cox_scc)



# 3. Mast cell tumor
thrombo_class_mast <- thrombo_class_neo %>% 
  filter(Diagnosis == "Mast cell tumor")

# Kaplan-meier
fit_thromb_mast <- survfit(Surv(time = `Survival time (month)`, event = Censure == 1) ~ Intensity, 
                          data = thrombo_class_mast)

summary(fit_thromb_mast)$table

ggsurvplot(
  fit_thromb_mast,
  data = thrombo_class_mast,
  pval = TRUE,
  conf.int = FALSE,
  risk.table = TRUE,
  legend.title = "Intensity of thrombocytosis",   
  legend.labs = c("Mild", "Moderate-Severe"),
  palette = c("#1b9e77", "#d95f02"),
  xlab = "Survival time (months)",
  ylab = "Survival probability",
  font.x = 16,        
  font.y = 16,        
  font.legend = 16,   
  font.tickslab = 16  
)

  

# Multivariate Cox model
modelo_cox_mast <- coxph(
  Surv(time = `Survival time (month)`, event = Censure == 1) ~ Intensity + meta_status,
  data = thrombo_class_mast
)

summary(modelo_cox_mast)


# 4. Hemangiosarcoma
thrombo_class_heman <- thrombo_class_neo %>% 
  filter(Diagnosis == "Hemangiosarcoma")

# Kaplan-meier
fit_thromb_heman <- survfit(Surv(time = `Survival time (month)`, event = Censure == 1) ~ Intensity, 
                           data = thrombo_class_heman)


ggsurvplot(
  fit_thromb_heman,
  data = thrombo_class_heman,
  pval = TRUE,
  conf.int = FALSE,
  risk.table = TRUE,
  legend.title = "Intensity of thrombocytosis",   
  legend.labs = c("Mild", "Moderate-Severe"),
  palette = c("#1b9e77", "#d95f02"),
  xlab = "Survival time (months)",
  ylab = "Survival probability"
)



# Multivariate Cox model
modelo_cox_heman <- coxph(
  Surv(time = `Survival time (month)`, event = Censure == 1) ~ Intensity + meta_status,
  data = thrombo_class_heman
)

summary(modelo_cox_heman)


# 5. Sarcoma
thrombo_class_sarcoma <- thrombo_class_neo %>% 
  filter(Diagnosis == "Sarcoma")

# Kaplan-meier
fit_thromb_sarcoma <- survfit(Surv(time = `Survival time (month)`, event = Censure == 1) ~ Intensity, 
                            data = thrombo_class_sarcoma)


ggsurvplot(
  fit_thromb_sarcoma,
  data = thrombo_class_sarcoma,
  pval = TRUE,
  conf.int = FALSE,
  risk.table = TRUE,
  legend.title = "Intensity of thrombocytosis",   
  legend.labs = c("Mild", "Moderate-Severe"),
  palette = c("#1b9e77", "#d95f02"),
  xlab = "Survival time (months)",
  ylab = "Survival probability"
)



# Multivariate Cox model
modelo_cox_sarcoma <- coxph(
  Surv(time = `Survival time (month)`, event = Censure == 1) ~ Intensity + meta_status,
  data = thrombo_class_sarcoma
)

summary(modelo_cox_sarcoma)






