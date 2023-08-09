############################# CHECK MASTER #####################################

library(readxl)
library(writexl)
library(openxlsx)
library(tidyverse)
library(stringr)

# Load the the relevant data set to be used for the project CHECK

setwd("C:/Users/ryanr/Reclaim Finance/Reclaim Cloud - Documents/5. Donnees Financieres/GFANZ")
cred_ff <- read_excel("C:/Users/ryanr/Reclaim Finance/Reclaim Cloud - Documents/5. Donnees Financieres/GFANZ/230515_creditor_ff-exit_join.xlsx")

# Certain variables need to be transformed into the correct class, the below is based upon the variable names in the GCEL after applying the "clean_names" function

numeric_var_ff <- c(
  "hydrocarbons_production_in_2021_mmboe",
  "resources_under_development_and_field_evaluation_as_of_september_2022_mmboe",
  "installed_coal_power_capacity_mw",
  "annual_coal_production_in_million_metric_tons", 
  "coal_share_of_revenue_short",
  "expansion_plans_coal_power_total_in_mw")

cred_ff <- cred_ff %>% 
  mutate_at(vars(all_of(numeric_var_ff)), 
            ~ as.numeric(str_replace_all(
              ., "[<>NI/]",""))
  )

################################################################################
####################### RBC
################################################################################

rbc_coal <- cred_ff %>% filter(
  investor_parent == "Royal Bank of Canada" &
    gcel_company == "Yes" &
    company_in_scope == "Yes"
)

rbc_coal_pivot <- rbc_coal %>% 
  group_by(group) %>% 
  summarise(coal_share_of_revenue_short = first(na.omit(coal_share_of_revenue_short)),
            gcel_expansion_plans = first(na.omit(gcel_expansion_plans)),
            coal_share_of_revenue_thermal_total = first(na.omit(coal_share_of_revenue_thermal_total)))%>% 
  distinct()


# Create a new workbook
wb <- createWorkbook()

# Add the rbc_coal data frame to a new worksheet
addWorksheet(wb, "rbc_coal")
writeData(wb, "rbc_coal", rbc_coal)

# Add the rbc_coal_pivot data frame to a new worksheet
addWorksheet(wb, "rbc_coal_pivot")
writeData(wb, "rbc_coal_pivot", rbc_coal_pivot)

######### Saving the file  
###  Save the workbook to an Excel file
openxlsx::saveWorkbook(wb, "C:/Users/ryanr/Reclaim Finance/Reclaim Cloud - Documents/5. Donnees Financieres/GFANZ/Project_check/rbc_all_coal.xlsx", overwrite = T)


write_xlsx(rbc_coal, "C:/Users/ryanr/Reclaim Finance/Reclaim Cloud - Documents/5. Donnees Financieres/GFANZ/rbc_all_coal.xlsx")


################################################################################
####################### Natixis
################################################################################

natixis_all <- cred_ff %>% 
  filter(
  investor_parent == "Groupe BPCE" &
  gcel_company == "Yes" &
  company_in_scope == "Yes"
  )


# filter companies based on policy of bank
natixis_breach <- natixis_all %>% 
  filter(
    (closing_issue_filing_date >= "2020-10-01" & str_detect(expansion_plans, "mining"))|
    (closing_issue_filing_date >= "2020-10-01" & str_detect(expansion_plans, "power"))|
    (closing_issue_filing_date >= "2020-10-01" & coal_share_of_revenue_short >= 0.25)
  )

###########################pivot not to be used#################################
natixis_coal_pivot <- natixis_coal %>% 
  group_by(group) %>% 
  summarise(coal_share_of_revenue_short = first(na.omit(coal_share_of_revenue_short)),
            gcel_expansion_plans = first(na.omit(gcel_expansion_plans)),
            coal_share_of_revenue_thermal_total = first(na.omit(coal_share_of_revenue_thermal_total)))%>% 
  distinct()
################################################################################

# Create a new workbook
natixis <- createWorkbook()

# Add the natixis_coal data frame to a new worksheet
addWorksheet(natixis, "natixis_all_coal")
writeData(natixis, "natixis_all_coal", natixis_all)

# Add the natixis_coal_pivot data frame to a new worksheet
addWorksheet(natixis, "natixis_breach")
writeData(natixis, "natixis_breach", natixis_breach)

# Save the workbook to an Excel file
openxlsx::saveWorkbook(natixis, "C:/Users/ryanr/Reclaim Finance/Reclaim Cloud - Documents/5. Donnees Financieres/GFANZ/Project_check/natixis_all_coal.xlsx", overwrite = T)


################################################################################
###################### Sogen
################################################################################

# filter based on the correct criteria for Sogen
sogen_coal <- cred_ff %>% filter(
  investor_parent == "Société Générale" &
    gcel_company == "Yes" &
    company_in_scope == "Yes") 


sogen_breach <- sogen_coal %>%
  filter(
    (closing_issue_filing_date >= "2020-07-01" & str_detect(expansion_plans, "mining"))|
    (closing_issue_filing_date >= "2022-01-01" & str_detect(expansion_plans, "power"))|
    (closing_issue_filing_date >= "2022-01-01" & str_detect(expansion_plans, "infrastructure"))|
    (closing_issue_filing_date >= "2020-07-01" & coal_share_of_revenue_short >= 25) |
    (closing_issue_filing_date >= "2020-07-01" & annual_coal_production_thermal_total > 10)
    )

###########################pivot not to be used#################################
sogen_coal_pivot <- sogen_breach %>% 
  group_by(group) %>% 
  summarise(coal_share_of_revenue_short = first(na.omit(coal_share_of_revenue_short)),
            gcel_expansion_plans = first(na.omit(gcel_expansion_plans)),
            coal_share_of_revenue_thermal_total = first(na.omit(coal_share_of_revenue_thermal_total)))%>% 
  distinct()
################################################################################


# Create a new workbook
sogen <- loadWorkbook("C:/Users/ryanr/Reclaim Finance/Reclaim Cloud - Documents/5. Donnees Financieres/GFANZ/Project_check/sogen_all_coal.xlsx")

removeWorksheet(sogen, sheet = "sogen_coal_breach")
removeWorksheet(sogen, sheet = "sogen_all_coal")

# Add the _coal data frame to a new worksheet
addWorksheet(sogen, "sogen_all_coal")
writeData(sogen, "sogen_all_coal", sogen_coal)

# Add the _coal_pivot data frame to a new worksheet
addWorksheet(sogen, "sogen_coal_breach")
writeData(sogen, "sogen_coal_breach", sogen_breach)

# Save the workbook to an Excel file
openxlsx::saveWorkbook(sogen, "C:/Users/ryanr/Reclaim Finance/Reclaim Cloud - Documents/5. Donnees Financieres/GFANZ/Project_check/sogen_all_coal.xlsx", overwrite = T)


################################################################################
###################### CA
################################################################################


CA_coal <- cred_ff %>% filter(
  investor_parent == "Crédit Agricole" &
    gcel_company == "Yes" &
    company_in_scope == "Yes") 


CA_breach <- CA_coal %>% 
  filter(
    (closing_issue_filing_date >= "2021-01-01" & str_detect(expansion_plans, "mining"))|
    (closing_issue_filing_date >= "2021-01-01" & str_detect(expansion_plans, "power"))|
    (closing_issue_filing_date >= "2021-01-01" & str_detect(expansion_plans, "infrastructure"))|
    (closing_issue_filing_date >= "2021-01-01" & coal_share_of_revenue_short >= 25)
  )

CA_coal_pivot <- CA_breach %>% 
  group_by(group) %>% 
  summarise(coal_share_of_revenue_short = first(na.omit(coal_share_of_revenue_short)),
            gcel_expansion_plans = first(na.omit(gcel_expansion_plans)),
            coal_share_of_revenue_thermal_total = first(na.omit(coal_share_of_revenue_thermal_total)))%>% 
  distinct()


# Create a new workbook
CA <- createWorkbook()

# Add the rbc_coal data frame to a new worksheet
addWorksheet(CA, "CA")
writeData(CA, "CA", CA_coal)

# Add the rbc_coal_pivot data frame to a new worksheet
addWorksheet(CA, "CA_coal_breach")
writeData(CA, "CA_coal_pivot", CA_breach)

# Save the workbook to an Excel file
openxlsx::saveWorkbook(CA, "C:/Users/ryanr/Reclaim Finance/Reclaim Cloud - Documents/5. Donnees Financieres/GFANZ/Project_check/CA_all_coal.xlsx", overwrite = T)

################################################################################
###################### BNP
################################################################################


BNP_coal <- cred_ff %>% filter(
  investor_parent == "BNP Paribas" &
    gcel_company == "Yes" &
    company_in_scope == "Yes") 


BNP_breach <- BNP_coal %>% 
  filter(
    (closing_issue_filing_date >= "2020-07-01" & str_detect(expansion_plans, "mining"))|
    (closing_issue_filing_date >= "2020-07-01" & coal_share_of_revenue_short >= 20)| #group loophole
    (closing_issue_filing_date >= "2022-01-01" & str_detect(expansion_plans, "power"))|
    (closing_issue_filing_date >= "2022-01-01" & coal_share_of_revenue_short >= 25)|  
    (closing_issue_filing_date >= "2022-01-01" & str_detect(expansion_plans, "infrastructure"))|
    (closing_issue_filing_date >= "2022-01-01" & coal_share_of_revenue_short >= 20)
  )

# Create a new workbook
BNP <- createWorkbook()

# Add the rbc_coal data frame to a new worksheet
addWorksheet(BNP, "BNP_all_coal")
writeData(BNP, "BNP_all_coal", BNP_coal)

# Add the rbc_coal_pivot data frame to a new worksheet
addWorksheet(BNP, "BNP_coal_breach")
writeData(BNP, "BNP_coal_breach", BNP_breach)

# Save the workbook to an Excel file
openxlsx::saveWorkbook(BNP, "C:/Users/ryanr/Reclaim Finance/Reclaim Cloud - Documents/5. Donnees Financieres/GFANZ/Project_check/BNP_all_coal.xlsx", overwrite = T)


