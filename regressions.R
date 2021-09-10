library(tidyverse)
library(openxlsx)
library(stargazer)

obs <- read.xlsx(file.choose(), 1)

view(obs)

#GET DESCRIPTIVE STATISTICS USING Stargazer
#Select variables
obs_DS <- obs %>%
  select(c(NPL:Inflation_CPI), -Total_Assets)

stargazer(obs_DS, out = "Descriptive Statistics.html", header = FALSE, 
          type = "html", title = "Descriptive Stasitics", digits = 3)

#GET CORRELATION MATRIX
#Select variables
obs_matrix <- obs %>%
  select(c(Tier1_Ratio:Income_Div), -Total_Assets)

obs_matrix

#Get correlation matrix rounded into 3 decimals
mcor <- round(cor(obs_matrix, use = "complete.obs"), 3)

#Get the LOWER triangle of the matrix
upper.tri(mcor, diag = FALSE)

#Hide the upper triangle
tri_mcor <- mcor
tri_mcor[upper.tri(mcor)] <- ""

#Transfer tri_mcor into a dataframe
df_tri_mcor <- as.data.frame(tri_mcor)

#Write the matrix to xlsx
write.xlsx(df_tri_mcor, file = "Correlation Matrix.xlsx", 
           sheetName = "Correlation Matrix", colNames = TRUE, rowNames = TRUE,
           keepNA = TRUE)


#GET YEAR FIXED OLS BASELINE REGRESSION TABLE USING Stargazer
#Coverting formats
obs$Year <- as.factor(obs$Year)
obs$Ownership <- as.factor(obs$Ownership)

str(obs)
#NPL ~ Tier1_Ratio
Linear_NPL_1 <- lm(NPL ~ Tier1_Ratio + LnAssets + NL_to_TA + ROE +
                     Income_Div + Concentration + GDP_Growth + Inflation_CPI +
                     Year, data = obs)
#NPL ~ TC_to_RWA
Linear_NPL_2 <- lm(NPL ~ TC_to_RWA + LnAssets + NL_to_TA + ROE +
                     Income_Div + Concentration + GDP_Growth + Inflation_CPI +
                     Year, data = obs)
#LLR ~ Tier1_Ratio
Linear_LLR_1 <- lm(LLR ~ Tier1_Ratio + LnAssets + NL_to_TA + ROE +
                     Income_Div + Concentration + GDP_Growth + Inflation_CPI +
                     Year, data = obs)
#LLR ~ TC_to_RWA
Linear_LLR_2 <- lm(LLR ~ TC_to_RWA + LnAssets + NL_to_TA + ROE +
                     Income_Div + Concentration + GDP_Growth + Inflation_CPI +
                     Year, data = obs)

#Write the baseline regression table
stargazer(Linear_NPL_1, Linear_NPL_2, Linear_LLR_1, Linear_LLR_2, type = "html", 
          out = "OLS Baseline.html", align = TRUE, 
          title = "Baseline Regression - Year Fixed Effect Model", 
          omit = "Year", omit.labels = "Year Fixed Effects", no.space = TRUE)


#GET FIXED EFFECT MODEL-OWNERSHIP STRUCTURE
#NPL ~ Tier1_Ratio
Linear_NPL_1 <- lm(NPL ~ Tier1_Ratio + Ownership + LnAssets + NL_to_TA + ROE +
                     Income_Div + Concentration + GDP_Growth + Inflation_CPI +
                     Year, data = obs)
#NPL ~ TC_to_RWA
Linear_NPL_2 <- lm(NPL ~ TC_to_RWA + Ownership + LnAssets + NL_to_TA + ROE +
                     Income_Div + Concentration + GDP_Growth + Inflation_CPI +
                     Year, data = obs)
#LLR ~ Tier1_Ratio
Linear_LLR_1 <- lm(LLR ~ Tier1_Ratio + Ownership + LnAssets + NL_to_TA + ROE +
                     Income_Div + Concentration + GDP_Growth + Inflation_CPI +
                     Year, data = obs)
#LLR ~ TC_to_RWA
Linear_LLR_2 <- lm(LLR ~ TC_to_RWA + Ownership + LnAssets + NL_to_TA + ROE +
                     Income_Div + Concentration + GDP_Growth + Inflation_CPI +
                     Year, data = obs)

#Write the regression table-ownership
stargazer(Linear_NPL_1, Linear_NPL_2, Linear_LLR_1, Linear_LLR_2, type = "html", 
          out = "Year-fixed Ownership Model.html", align = TRUE, 
          title = "Fixed Effect Regression - Ownership Structure",
          omit = "Year", omit.labels = "Year Fixed Effects",no.space = TRUE)


#GET INTERACTION REGRESSION MODEL
#NPL ~ Tier1_Ratio
Interaction_NPL_1 <- lm(NPL ~ Tier1_Ratio + Ownership + 
                          Tier1_Ratio:Ownership + LnAssets + NL_to_TA + ROE + 
                          Income_Div + Concentration + GDP_Growth + 
                          Inflation_CPI + Year, data = obs)
#NPL ~ TC_to_RWA
Interaction_NPL_2 <- lm(NPL ~ TC_to_RWA + Ownership + 
                          TC_to_RWA:Ownership + LnAssets + NL_to_TA + ROE + 
                          Income_Div + Concentration + GDP_Growth + 
                          Inflation_CPI + Year, data = obs)
#LLR ~ Tier1_Ratio
Interaction_LLR_1 <- lm(LLR ~ Tier1_Ratio + Ownership + 
                          Tier1_Ratio:Ownership + LnAssets + NL_to_TA + ROE + 
                          Income_Div + Concentration + GDP_Growth + 
                          Inflation_CPI + Year, data = obs)
#LLR ~ TC_to_RWA
Interaction_LLR_2 <- lm(LLR ~ TC_to_RWA + Ownership + 
                          TC_to_RWA:Ownership + LnAssets + NL_to_TA + ROE + 
                          Income_Div + Concentration + GDP_Growth + 
                          Inflation_CPI + Year, data = obs)

#Write the interaction regression table
stargazer(Interaction_NPL_1, Interaction_NPL_2, Interaction_LLR_1, Interaction_LLR_2, 
          type = "html", out = "Regression Table-Interaction.html", align = TRUE,
          title = "Regression with Interaction between Ownership and Regulation",
          omit = "Year", omit.labels = "Year Fixed Effects", no.space = TRUE)