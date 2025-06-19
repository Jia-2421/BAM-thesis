library(readxl)
library(dplyr)
library(stargazer)
library(plm)
library(plm)
library(lmtest)
library(sandwich)
library(zoo)
library(writexl)
library(readxl)
library(dplyr)
library(stringdist)
library(writexl) 
library(lubridate)
library(stringr)
# Define the file path
path <- "C:/Users/jiaju/Desktop"
path <- "D:/666"
# Read Sheet2 from the Excel file
df <- read_excel(file.path(path, "1116(other&kaggle).xlsx"), sheet = "ESG_combined")
cds <- read_excel(file.path(path, "match.xlsx"), sheet = "Sheet15")


##########Matching datasets
# Ensure both columns are character type
df$Company <- as.character(df$Company)
df1$ESG <- as.character(df1$ESG)

# Filter rows in df where Company exists in df1$ESG
matched_df <- df[df$Company %in% df1$ESG, ]

write_xlsx(matched_df, "C:/Users/jiaju/Desktop/ALL_ESG.xlsx")


###########Separate Company with the score name
df_separated <- df %>%
  separate(
    col   = NAME,
    into  = c("company", "score_name"),
    sep   = " - ",    # that is: space, en‐dash (U+2013), space
    extra = "merge",  # if by chance there are more than two pieces, merge them into `score_name`
    fill  = "right"   # if a row has no separator at all, put everything into `company`, leave score_name = NA
  )
write_xlsx(df_separated, "C:/Users/jiaju/Desktop/control_fill.xlsx")



########Extract CDS company name 
df1 <- read_excel(file.path(path, "matched_names.xlsx"), sheet = "Canada ESG")
# Extract part before SEN, SNR, or SUB (case insensitive)
df$Name <- sub("\\s+(SEN|SNR|SUB).*", "", df$Name, ignore.case = TRUE)

# View result
print(df)

# Keep only the first occurrence of each unique company name
df_unique <- df[!duplicated(df$Name), ]

write_xlsx(df_unique, "C:/Users/jiaju/Desktop/matched_names_parEU.xlsx")


###############Matching CDS and ESG names
# Step 3: Get unique values
companies <- unique(df$CDS)
us_firms <- unique(df$ESG)
 
# Step 4: Generate all pair combinations and calculate similarity
all_matches <- expand.grid(Company = companies, US_Firm = us_firms, stringsAsFactors = FALSE) %>%
  mutate(Similarity = 1 - stringdist(tolower(Company), tolower(US_Firm), method = "jw"))

# Step 5: Keep the best match per company
best_matches <- all_matches %>%
  group_by(Company) %>%
  slice_max(Similarity, n = 1, with_ties = FALSE)

# Step 6: Ensure each US firm is used only once
final_matches <- best_matches %>%
  arrange(desc(Similarity)) %>%
  distinct(US_Firm, .keep_all = TRUE)
 
write_xlsx(final_matches, "C:/Users/jiaju/Desktop/matched_names_CDSESG.xlsx")

####################Replace NA with mean
# Read Sheet2 from the Excel file
df <- read_excel(file.path(path, "1ST_ESG.xlsx"), sheet = "Sheet1")
df_wide <- read_excel(file.path(path, "1ST_ESG.xlsx"), sheet = "Sheet3")
## fill in NA and empty cells with mean
# Replace empty strings ("") with NA
df[df == ""] <- NA

# Convert columns to numeric (skip Year column if present)
df[, -1] <- lapply(df[, -1], function(x) as.numeric(as.character(x)))

# Calculate the average per column (excluding NAs)
col_means <- colMeans(df[, -1], na.rm = TRUE)

# Fill NAs with column averages
for (col in names(col_means)) {
  df[[col]][is.na(df[[col]])] <- col_means[col]
}

write_xlsx(df, "C:/Users/jiaju/Desktop/1st_esg_mean.xlsx")

#################Pivot ESG tale into long format with quarters
eu_esg_clean <- read_excel(file.path(path, "1116(other&kaggle).xlsx"), sheet = "ESG_combined")
# Step 1: Split NAME column into Company and Metric
eu_esg_clean <- eu_esg_clean %>%
  separate(NAME, into = c("Company_Name", "Metric"), sep = " - ", remove = FALSE)

# Step 2: Convert year columns to character
year_cols <- as.character(2016:2024)
eu_esg_clean[year_cols] <- lapply(eu_esg_clean[year_cols], as.character)

# Step 3: Pivot to long format
esg_long <- eu_esg_clean %>%
  pivot_longer(cols = all_of(year_cols), names_to = "Year", values_to = "Score")

# Step 4: Expand each year into 4 quarters
esg_quarters <- esg_long %>%
  slice(rep(1:n(), each = 4)) %>%
  mutate(
    Quarter = rep(c("Q1", "Q2", "Q3", "Q4"), times = nrow(esg_long)),
    Date = paste0(Quarter, " ", Year)
  )

# Step 5: Final wide table: 1 row = Company × Date, 1 column per score
esg_final <- esg_quarters %>%
  select(Date, Quarter, Year, Company_Name, Metric, Score,
         `Company`, `COUNTRY`, `ISIN_CODE`, `TICKER SYMBOL`,
         `ICB industry name`, `ICB sector name`) %>%
  pivot_wider(
    names_from = Metric,
    values_from = Score,
    values_fn = ~ .[1]
  ) %>%
  relocate(Date, Quarter, Year, Company_Name, .before = everything())

# Step 6: Export to Excel
write_xlsx(esg_final, "esg_2nd_long_format.xlsx")


#################Pivot CDS table into long format
# Read Sheet2 from the Excel file
cds <- read_excel(file.path(path, "match.xlsx"), sheet = "Sheet15")

# Step 2: Pivot into long format
df_long <- cds %>%
  pivot_longer(
    cols = -Company,                # All columns except 'Company'
    names_to = "Quarter_Year",      # Combine 'QX YYYY' into one column
    values_to = "cds"
  ) %>%
  separate(Quarter_Year, into = c("Quarter", "Year"), sep = " ") %>%
  mutate(
    Year = as.integer(Year),
    cds = as.numeric(str_replace(cds, ",", "."))  # Convert decimal comma to point
  ) %>%
  select(Company, Year, Quarter, cds)  # Arrange the columns


write_xlsx(df_long, "C:/Users/jiaju/Desktop/cds_long.xlsx")

################Assign numbers to control variables for matching
# Read Sheet2 from the Excel file
df <- read_excel(file.path(path, "1116.xlsx"), sheet = "Firm_market")

df <- df %>%
  group_by(Company) %>%
  mutate(score_id = rep(1:11, length.out = n())) %>%
  ungroup()

write_xlsx(df, "C:/Users/jiaju/Desktop/numbered_control.xlsx")

#################Pivot Control variables table into long format
#1 ── Gather the quarter columns into long format ────────────────────────────
#   • turns `Q1 2016`, `Q2 2016`, … into rows
#   • pulls out `quarter` (=1‒4) and `year` (YYYY)
df_long <- df_raw %>% 
  pivot_longer(
    cols = matches("^Q[1-4]\\s+\\d{4}$"),       # every "Q# YYYY" column
    names_to = c("quarter", "year"),
    names_pattern = "Q([1-4])\\s+(\\d{4})",
    values_to = "value"
  )

# 3) ── Widen *metrics* so each score_name becomes its own column ─────────────
tidy_panel <- df_long %>% 
  select(Company, year, quarter, score_name, value) %>% 
  mutate(
    value = str_replace(value, ",", ".") |> as.numeric()   # optional: EU → US decimal
  ) %>% 
  pivot_wider(
    names_from  = score_name,      # the 11 metrics you listed
    values_from = value,
    values_fn   = first            # keep the 1st if (rarely) duplicates exist
  ) %>% 
  arrange(Company, year, quarter)  # nice ordering

write_xlsx(tidy_panel, "CONTROL_LONG3.xlsx")

###################Convert currency
final_df<- read_excel(file.path(path, "FINAL.xlsx"), sheet = "Sheet1")
Currency_df <- read_excel(file.path(path, "Currency.xlsx"), sheet = "Sheet1")
# Step 1: Merge currency data with main dataset based on Country
final_df <- final_df %>%
  left_join(currency_df, by = "Country")

# Step 2: Convert TOTAL ASSET and MARKET VALUE using exchange rate
final_df <- final_df %>%
  mutate(
    TOTAL_ASSET_USD = TOTAL_ASSET * currency_rate,
    MARKET_VALUE_USD = MARKET_VALUE * currency_rate
  )

##################Add risk free rate
# Read Sheet2 from the Excel file
df <- read_excel(file.path(path, "match.xlsx"), sheet = "ESG")
df2 <- read_excel(file.path(path, "match.xlsx"), sheet = "Risk free rate")

# Step 1: Reshape df2 to long format
df2_long <- df2 %>%
  pivot_longer(cols = starts_with("Q"), names_to = "Date", values_to = "Value") %>%
  rename(Country = Name)

# Step 2: Standardize column formats (preserve "Q1 2016" order)
df2_long <- df2_long %>%
  mutate(
    Country = str_trim(toupper(Country)),
    Date = str_trim(Date)
  )

# Step 3: Prepare df to match same format
df <- df %>%
  mutate(
    Country = str_trim(toupper(Country)),
    Date = str_trim(Date)  # should already be in "Q1 2016" format
  )

# Step 4: Join on Country and Date
df_joined <- df %>%
  left_join(df2_long, by = c("Country", "Date"))


write_xlsx(df_joined, "C:/Users/jiaju/Desktop/df_joined.xlsx")

################Add Industry and country dummy variables
# Define the lists
target_countries <- c(
  "AUSTRIA", "BELGIUM", "CZECH REPUBLIC", "DENMARK", "FINLAND", "FRANCE", "GERMANY", "GREECE",
  "ITALY", "NETHERLANDS", "NORWAY", "PORTUGAL", "SPAIN", "SWEDEN", "SWITZERLAND", 
  "UNITED KINGDOM", "UNITED STATES"
)

target_industries <- c(
  "Basic Materials", "Consumer Discretionary", "Consumer Staples", "Energy", "Financials",
  "Health Care", "Industrials", "Real Estate", "Technology", "Telecommunications", "Utilities"
)

# Clean column values
df$Country <- toupper(trimws(df$Country))
df$`ICB industry name` <- trimws(df$`ICB industry name`)

# Dummy columns for countries — using country name as column
for (country in target_countries) {
  col_name <- gsub(" ", "_", country)
  df[[col_name]] <- ifelse(df$Country == country, 1, 0)
}

# Dummy columns for industries — using industry name as column
for (industry in target_industries) {
  col_name <- gsub(" ", "_", industry)
  df[[col_name]] <- ifelse(df$`ICB industry name` == industry, 1, 0)
  
write_xlsx(df, "C:/Users/jiaju/Desktop/FINAL.xlsx")
  