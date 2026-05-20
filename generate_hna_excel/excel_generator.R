library(readr)
library(tidyverse)
library(data.table)
library(openxlsx2)
library(dplyr)
library(janitor)
library(lubridate)


#'==============================================================================
# CONFIGURATION ----
#'==============================================================================

#' Define the list of CSD DGUIDs to process. Add or remove codes as needed.
community_list <- c(
  2448050,
  1001101,
  1102075,
  1101002,
  1103061,
  1209034,
  5926014
)


output_dir <- "output"
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)


setwd(".")



#'==============================================================================
# BROAD LOAD (once, outside loop) ----
#'==============================================================================

#' Load all input datasets once to avoid repeated I/O inside the loop.
#' Large tables are keyed after loading for fast binary-search filtering.

## ...UTF-8 CLEANER ----
#' Ensures all character columns are valid UTF-8 before writing to the
#' XML-based xlsx format. Prevents illegal XML character errors from accented
#' place names (e.g. French Canadian geographies).
clean_utf8 <- function(df) {
  df <- as.data.frame(df)
  char_cols <- vapply(df, is.character, logical(1))
  df[char_cols] <- lapply(df[char_cols], function(x) {
    iconv(x, from = "CP1252", to = "UTF-8", sub = "byte")
  })
  df
}

## ...HART DATA ----
hart_ft_df  <- fread("input/hart_ft.csv")
hart_df     <- fread("input/hart.csv")
hart16_df   <- fread("input/hart2016.csv")

setkey(hart_ft_df, dguid)
setkey(hart_df,    geo)

## ...HART LOOKUPS ----
char <- fread("input/convert_char.csv")
ind  <- fread("input/convert_ind.csv")
shel <- fread("input/convert_sh.csv")

## ...VARID SETS (defined before census load so they can be used to pre-filter) ----
lf_varid         <- c(2224:2230, 2237:2245, 2593:2610)
naics_varid      <- c(2259:2281)
pop16_varid      <- c(8, 10:12, 14:23, 25:29, 39:40)
pop21_varid      <- c(8, 10:12, 14:23, 25:29, 30:40)
imm_vis_varid    <- c(1522:1537, 1665:1675, 1683:1684)
mob1_varid       <- c(1974:1982)
mob5_varid       <- c(1983:1991)
lknow_varid      <- c(383, 387)
educ_varid       <- c(2014:2017)
indi_varid       <- c(1402:1410)
misc_inc_varid   <- c(243, 252, 1494)
main16_varid     <- c(1658:1666)
main21_varid     <- c(1456:1464)
dtype_varid      <- c(4, 41:49)
dage_varid       <- c(1440:1448)
chn_census_varid <- c(1437:1439, 1449:1451, 1465:1499)

## ...CENSUS ----
#' Pre-filter to the union of all VARIDs ever needed before keying.
#' This shrinks the table each per-CSD DGUID filter has to scan on every iteration.
all_varid21 <- unique(c(
  lf_varid, naics_varid, pop21_varid, imm_vis_varid,
  indi_varid, educ_varid, lknow_varid, mob1_varid, mob5_varid,
  misc_inc_varid, main21_varid, dtype_varid, dage_varid, chn_census_varid
))
all_varid16 <- unique(c(pop16_varid, main16_varid))

census16_raw <- fread("input/census2016.csv")
census21_raw <- fread("input/census2021.csv")

census16 <- census16_raw[VARID %in% all_varid16]
census21 <- census21_raw[VARID %in% all_varid21]
rm(census16_raw, census21_raw)

setkey(census16, DGUID, VARID)
setkey(census21, DGUID, VARID)

flag16 <- fread("input/2016_census_flag.csv") %>%
  filter(flag == "y") %>%
  select(varid, var)

flag21 <- fread("input/2021_census_flag.csv") %>%
  filter(flag == "y") %>%
  select(varid, var)

## ...PLACE NAMES ----
names_df <- fread("input/Census2021_names.csv")

## ...CMA REFERENCE ----
ref_cma <- fread("input/CMA_CA Block Reference list by CSD _ RMR_AR Liste de référence des îlots par SDR.csv") %>%
  distinct(CSDUID_SDRIDU, CMAPUID_RMRPIDU) %>%
  rename(dguid = 1, cma = 2) %>%
  mutate(
    cma = as.character(cma),
    cma = as.numeric(substr(cma, nchar(cma) - 2, nchar(cma)))
  )

## ...STIR / CHN ----
stir_df <- fread("input/stir2021.csv")

## ...CONDITION / SUITABILITY ----
cond_df_raw <- fread("input/cond2021.csv")

## ...COMMUTE ----
commute_raw <- fread("input/commute_csd.csv")
setkey(commute_raw, WORK, RES)

## ...PROJECTIONS ----
projections_raw <- fread("input/projections_csd.csv")
setkey(projections_raw, dguid)

## ...NON-MARKET ----
nm_raw_df <- fread("input/nonmarket.csv") %>%
  clean_names() %>%
  select(csd_code, county, housing_type, number_of_beds, number_of_units)

## ...WORK FROM HOME ----
wfh_raw <- fread("input/wfh_csd.csv")

## ...POPULATION ESTIMATES ----
estimates_csd_raw <- fread("input/estimates_csd.csv")
estimates_cd_raw  <- fread("input/estimates_cd_years.csv")

## ...DEMOGRAPHIC COMPONENTS (CD-level; all files combined once) ----
comp_cd_raw <- rbind(
  fread("input/components_cd_births.csv"),
  fread("input/components_cd_deaths.csv"),
  fread("input/components_cd_emi.csv"),
  fread("input/components_cd_imm.csv"),
  fread("input/components_cd_nperm.csv"),
  fread("input/components_cd_interpr.csv"),
  fread("input/components_cd_intrapr.csv"),
  fread("input/components_cd_res.csv") %>%
    mutate('2021/2022' = 0, '2022/2023' = 0, '2023/2024' = 0, '2024/2025' = 0)
)

## ...DISABILITY HEADSHIP ----
headship_raw <- fread("input/headship_disability.csv")

## ...COSTS / RENT ----
cost_index_raw <- fread("input/cost_index_csd.csv")
avg_rent_raw   <- fread("input/average_rent_cma.csv")
cmhc_rms_raw   <- fread("input/cmhc_rms.csv")
sales_raw      <- fread("input/sales_csd.csv")
tuition_raw    <- fread("input/tuition_csd.csv")

## ...VULNERABILITY / CLIMATE INDICES ----
cimd_raw      <- fread("input/cimd_csd.csv")
cisr_raw      <- fread("input/cisr_csd.csv")
cisv_raw      <- fread("input/cisv_csd.csv")
msdi_raw      <- fread("input/msdi_csd.csv")
flood_raw     <- fread("input/flood_csd.csv")
coastal_raw   <- fread("input/flood_pei_coastal.csv")
humid_raw     <- fread("input/humidex_csd.csv")
hot_raw       <- fread("input/hotdays_csd.csv")
cold_raw      <- fread("input/colddays_csd.csv")
rain_raw      <- fread("input/precipitation_csd.csv")
clim_vuln_raw <- fread("input/clim_vuln_csd.csv")

## ...PERMITS ----
permits_raw <- fread("input/permits_csd.csv")

## ...INCOME ----
inc_source_raw <- fread("input/income_source2021.csv")
inc_char_raw   <- fread("input/income_char.csv")
inc_hart_raw   <- fread("input/income_hart.csv")
mbm_raw        <- fread("input/mbm_csd.csv")
wage_raw       <- fread("input/wage_pr.csv")
lwage_raw      <- fread("input/livingwage_csd.csv")
incdist_raw    <- fread("input/incomedist_csd.csv")

## ...HOUSING ----
cmhc_scss_raw    <- fread("input/cmhc_scss.csv")
hhtype_main_raw  <- fread("input/hhtype_main_dw_csd.csv")
hhsize_bed_raw   <- fread("input/hhsize_bed_dw_csd.csv")
echn_raw         <- fread("input/echn2021.csv")
shelters_raw_all <- fread("input/shelters_csd.csv")
dw_coll_raw      <- fread("input/collective_pr.csv")
dw_str_raw       <- fread("input/str_csd.csv")
transit_raw      <- fread("input/transit_csd.csv")

## ...INVESTMENT ----
prop_res_raw <- fread("input/property_resident_own_csd.csv")
prop_own_raw <- fread("input/property_owner_demo_csd.csv")

## ...WORKBOOK TEMPLATE (loaded once; cloned each iteration) ----
#' Clone this object at the start of each iteration rather than re-reading
#' from disk, which avoids repeated I/O for every CSD processed.
wb_template <- wb_load("input/RKI databook.xlsx")

cat("All broad inputs loaded.\n")


#'==============================================================================
# LOOP: ONE OUTPUT PER CSD ----
#'==============================================================================

for (community in community_list) {
  
  cat("\n------------------------------------------------------------\n")
  cat(sprintf("Processing CSD: %s\n", community))
  cat("------------------------------------------------------------\n")
  csd_start <- proc.time()[["elapsed"]]
  
  #'----------------------------------------------------------------------------
  # GEO SETUP ----
  #'----------------------------------------------------------------------------
  
  community_cd <- as.numeric(substr(community, 1, 4))
  community_pr <- as.numeric(substr(community, 1, 2))
  
  #' Combine CSD, CD, and PR codes for multi-level filtering throughout
  community_all <- c(community, community_cd, community_pr)
  
  geo_df <- data.table(
    geo = community_all
  ) %>%
    left_join(names_df, by = c("geo" = "DGUID")) %>%
    rename(name = GEO)
  
  
  #'----------------------------------------------------------------------------
  # COMMUTE ----
  #'----------------------------------------------------------------------------
  
  cat("  Loading commute data...\n")
  
  #' All workers commuting into this CSD
  commute_in_all <- commute_raw %>% filter(WORK %in% community)
  
  #' All workers commuting out of this CSD (values negated for net flow charts)
  commute_out_all <- commute_raw %>%
    filter(RES %in% community) %>%
    mutate_at(c("TOTAL", "LATENT", "x15_24", "x25_34", "x35_44", "x45_54", "x55_64", "x65"), ~ . * -1)
  
  #' Full commute flow (in + out), used for broad commute table
  commute_df <- rbind(commute_in_all, commute_out_all)
  
  #' CSD codes with meaningful inflow (>1% share), used for projection summaries
  commute_in_csds <- commute_raw %>%
    filter(WORK %in% community, P_OUT > 0.01) %>%
    distinct(RES) %>%
    pull(RES)
  
  #' Influential commute flows only (>1% share, primary flow), for influence area table
  commute_in_infl <- commute_raw %>%
    filter(WORK %in% community, P_OUT > 0.01, FLOW == 1)
  
  commute_out_infl <- commute_raw %>%
    filter(RES %in% community, P_OUT > 0.01, FLOW == 1) %>%
    mutate_at(c("TOTAL", "LATENT", "x15_24", "x25_34", "x35_44", "x45_54", "x55_64", "x65"), ~ . * -1)
  
  commute_infl_df <- rbind(commute_in_infl, commute_out_infl)
  
  
  #'----------------------------------------------------------------------------
  # ECONOMY ----
  #'----------------------------------------------------------------------------
  
  cat("  Loading economy data...\n")
  
  wfh_df <- wfh_raw %>% filter(DGUID %in% community)
  
  lf_df <- census21 %>%
    filter(DGUID %in% community_all, VARID %in% lf_varid) %>%
    left_join(flag21, by = c("VARID" = "varid"))
  
  naics_df <- census21 %>%
    filter(DGUID %in% community_all, VARID %in% naics_varid) %>%
    left_join(flag21, by = c("VARID" = "varid"))
  
  
  #'----------------------------------------------------------------------------
  # POPULATION ----
  #'----------------------------------------------------------------------------
  
  cat("  Loading population data...\n")
  
  #' 2016 Census population variables
  pop16_df <- census16 %>%
    filter(DGUID %in% community_all, VARID %in% pop16_varid) %>%
    left_join(flag21, by = c("VARID" = "varid")) %>%
    mutate(YEAR = 2016, SUB = "Census")
  
  #' 2021 Census population variables
  pop21_df <- census21 %>%
    filter(DGUID %in% community_all, VARID %in% pop21_varid) %>%
    left_join(flag21, by = c("VARID" = "varid")) %>%
    mutate(YEAR = 2021, SUB = "Census")
  
  #' Aggregated population of inbound commuter communities (for influence area context)
  pop21comm_df <- census21 %>%
    filter(DGUID %in% commute_in_csds, VARID %in% pop21_varid[pop21_varid %in% c(8, 10:12, 14:23, 25:29)]) %>%
    group_by(VARID) %>%
    summarize(VALUE = sum(VALUE, na.rm = TRUE)) %>%
    mutate(DGUID = 0, .before = VARID) %>%
    left_join(flag21, by = c("VARID" = "varid")) %>%
    mutate(YEAR = 2021, SUB = "Commute")
  
  pop_df <- rbind(pop16_df, pop21_df, pop21comm_df)
  rm(pop16_df, pop21_df, pop21comm_df)
  
  #' Immigration and visible minority
  imm_vis_df <- census21 %>%
    filter(DGUID %in% community_all, VARID %in% imm_vis_varid) %>%
    left_join(flag21, by = c("VARID" = "varid"))
  
  #' Population estimates
  est_csd_df <- estimates_csd_raw %>% filter(dguid %in% community_all)
  est_cd_df  <- estimates_cd_raw  %>% filter(dguid %in% community_cd)
  est_pr_df  <- estimates_cd_raw  %>% filter(dguid %in% community_pr)
  
  #' Demographic components of growth (CD-level)
  comp_cd_df <- comp_cd_raw %>% filter(dguid %in% community_cd)
  
  #' Disability headship rates (provincial)
  disa_hr_df <- headship_raw %>% filter(dguid %in% community_pr)
  
  #' Miscellaneous Census variables (mobility, education, language, immigration, income)
  census_misc_df <- census21 %>%
    filter(
      DGUID %in% community_all,
      VARID %in% c(indi_varid, educ_varid, lknow_varid, mob1_varid, mob5_varid, imm_vis_varid, misc_inc_varid)
    ) %>%
    left_join(flag21, by = c("VARID" = "varid"))
  
  
  #'----------------------------------------------------------------------------
  # PROJECTIONS ----
  #'----------------------------------------------------------------------------
  
  cat("  Loading projection data...\n")
  
  proj_df <- projections_raw %>% filter(dguid %in% community_all)
  
  #' Helper: pivot projections from wide to age-group summaries
  summarize_projections <- function(data) {
    data %>%
      select(-scenario) %>%
      pivot_longer(cols = 4:29, names_to = "year", values_to = "value") %>%
      as.data.table() %>%
      group_by(year) %>%
      summarize(
        TOTAL   = sum(value[row == 1],          na.rm = TRUE),
        '15_24' = sum(value[row %in% c(5:6)],   na.rm = TRUE),
        '25_34' = sum(value[row %in% c(7:8)],   na.rm = TRUE),
        '35_44' = sum(value[row %in% c(9:10)],  na.rm = TRUE),
        '45_54' = sum(value[row %in% c(11:12)], na.rm = TRUE),
        '55_64' = sum(value[row %in% c(13:14)], na.rm = TRUE),
        '65'    = sum(value[row %in% c(15:19)], na.rm = TRUE)
      ) %>%
      as.data.table()
  }
  
  proj_in_df  <- projections_raw %>% filter(dguid %in% commute_in_csds, scenario == "M4") %>% summarize_projections()
  proj_out_df <- projections_raw %>% filter(dguid %in% community,        scenario == "M4") %>% summarize_projections()
  
  #' Latent commuter summaries (incoming and outgoing)
  comm_in <- commute_in_infl %>%
    summarize(
      TOTAL   = sum(TOTAL,  na.rm = TRUE),
      HH      = sum(HH,     na.rm = TRUE),
      LATENT  = sum(LATENT, na.rm = TRUE),
      '15_24' = sum(x15_24, na.rm = TRUE),
      '25_34' = sum(x25_34, na.rm = TRUE),
      '35_44' = sum(x35_44, na.rm = TRUE),
      '45_54' = sum(x45_54, na.rm = TRUE),
      '55_64' = sum(x55_64, na.rm = TRUE),
      '65'    = sum(x65,    na.rm = TRUE)
    ) %>%
    as.data.table() %>%
    mutate(FLOW = "Incoming", .before = TOTAL)
  
  comm_out <- commute_out_infl %>%
    summarize(
      TOTAL   = sum(TOTAL,  na.rm = TRUE),
      HH      = sum(HH,     na.rm = TRUE) / n(),
      LATENT  = sum(LATENT, na.rm = TRUE),
      '15_24' = sum(x15_24, na.rm = TRUE),
      '25_34' = sum(x25_34, na.rm = TRUE),
      '35_44' = sum(x35_44, na.rm = TRUE),
      '45_54' = sum(x45_54, na.rm = TRUE),
      '55_64' = sum(x55_64, na.rm = TRUE),
      '65'    = sum(x65,    na.rm = TRUE)
    ) %>%
    as.data.table() %>%
    mutate(FLOW = "Outgoing", .before = TOTAL)
  
  comm_df <- rbind(comm_in, comm_out)
  rm(comm_in, comm_out, commute_in_infl, commute_out_infl)
  
  
  #'----------------------------------------------------------------------------
  # HOUSEHOLDS ----
  #'----------------------------------------------------------------------------
  
  cat("  Loading household data...\n")
  
  main16_df <- census16 %>%
    filter(DGUID %in% community_all, VARID %in% main16_varid) %>%
    left_join(flag16, by = c("VARID" = "varid")) %>%
    mutate(YEAR = 2016, .after = DGUID)
  
  main21_df <- census21 %>%
    filter(DGUID %in% community_all, VARID %in% main21_varid) %>%
    left_join(flag21, by = c("VARID" = "varid")) %>%
    mutate(YEAR = 2021, .after = DGUID)
  
  main_df <- rbind(main16_df, main21_df)
  rm(main16_df, main21_df)
  
  #' Household cross-tabulations (maintainer, type, size, bedroom) from shared source
  hh_cross_raw <- hhtype_main_raw %>% filter(dguid %in% community_all)
  hh_size_raw  <- hhsize_bed_raw  %>% filter(dguid %in% community_all)
  
  main_ten_df <- hh_cross_raw %>%
    filter(hhtype == "Total", dtype == "Total") %>%
    select(-hhtype, -dtype) %>%
    rename(char = hhmain)
  
  hhtype_df <- hh_cross_raw %>%
    filter(hhmain == "Total", dtype == "Total") %>%
    select(-hhmain, -dtype) %>%
    rename(char = hhtype)
  
  hhsize_df <- hh_size_raw %>%
    filter(bed == "Total", dtype == "Total") %>%
    select(-bed, -dtype) %>%
    rename(char = hhsize)
  
  bed_df <- hh_size_raw %>%
    filter(hhsize == "Total", dtype == "Total") %>%
    select(-hhsize, -dtype) %>%
    rename(char = bed)
  
  hh_ten_df <- rbind(main_ten_df, hhtype_df, hhsize_df, bed_df)
  rm(main_ten_df, hhtype_df, hhsize_df, bed_df, hh_cross_raw, hh_size_raw)
  
  
  #'----------------------------------------------------------------------------
  # COSTS ----
  #'----------------------------------------------------------------------------
  
  cat("  Loading cost / rent data...\n")
  
  cost_index_df <- cost_index_raw %>% filter(dguid %in% community_all)
  
  tuition_index_df <- tuition_raw %>%
    left_join(ref_cma, by = c("dguid" = "cma"), relationship = "many-to-many") %>%
    filter(geo != "Canada", nchar(as.character(dguid)) != 2) %>%
    mutate(dguid = dguid.y) %>%
    select(-dguid.y) %>%
    filter(dguid %in% community) %>%
    pivot_wider(names_from = ref_date, values_from = value)
  
  avg_rent_df <- avg_rent_raw %>%
    left_join(ref_cma, by = c("dguid" = "cma"), relationship = "many-to-many") %>%
    mutate(dguid = dguid.y) %>%
    select(-dguid.y) %>%
    filter(dguid %in% community_all) %>%
    pivot_wider(names_from = ref_date, values_from = value) %>%
    as.data.table()
  
  cmhc_rms_df <- cmhc_rms_raw %>% filter(csd  %in% community_all)
  sales_df    <- sales_raw    %>% filter(dguid %in% community_all)
  
  
  #'----------------------------------------------------------------------------
  # INDICES ----
  #'----------------------------------------------------------------------------
  
  cat("  Loading vulnerability and climate indices...\n")
  
  cimd_df      <- cimd_raw      %>% filter(csd    %in% community)
  cisr_df      <- cisr_raw      %>% filter(csd    %in% community)
  cisv_df      <- cisv_raw      %>% filter(csd    %in% community)
  msdi_df      <- msdi_raw      %>% filter(munic  %in% community)
  flood_df     <- flood_raw     %>% filter(DGUID  %in% community_all)
  coastal_df   <- coastal_raw   %>% filter(DGUID  %in% community_all)
  humid_df     <- humid_raw     %>% filter(CSDUID %in% community_all)
  hot_df       <- hot_raw       %>% filter(CSDUID %in% community_all)
  cold_df      <- cold_raw      %>% filter(CSDUID %in% community_all)
  rain_df      <- rain_raw      %>% filter(CSDUID %in% community_all)
  clim_vuln_df <- clim_vuln_raw %>% filter(DGUID  %in% community_all)
  
  
  #'----------------------------------------------------------------------------
  # ACTIVITY ----
  #'----------------------------------------------------------------------------
  
  cat("  Loading permits data...\n")
  
  permits_df <- permits_raw %>%
    filter(dguid %in% community_all) %>%
    pivot_wider(names_from = ref_date, values_from = total) %>%
    as.data.frame()
  
  
  #'----------------------------------------------------------------------------
  # INCOME ----
  #'----------------------------------------------------------------------------
  
  cat("  Loading income data...\n")
  
  inc_source_df <- inc_source_raw %>% filter(DGUID %in% community_all)
  inc_char_df   <- inc_char_raw   %>% filter(dguid %in% community_all)
  inc_hart_df   <- inc_hart_raw   %>% filter(geo   %in% community_all)
  
  #' Income categories by housing need threshold (HART)
  inc_cat_df <- hart_df %>%
    filter(geo %in% community_all, shelco %in% 1:6, indic %in% 1:3, char == 1) %>%
    arrange(indic) %>%
    select(-char) %>%
    left_join(shel, by = c("shelco" = "shelter")) %>%
    left_join(ind,  by = c("indic"  = "chn")) %>%
    mutate(shelco = desc_sh, indic = desc_chn) %>%
    select(-desc_sh, -desc_chn)
  
  mbm_df     <- mbm_raw     %>% filter(dguid %in% community_all)
  wage_df    <- wage_raw    %>% filter(dguid %in% community_pr)
  lwage_df   <- lwage_raw   %>% filter(dguid %in% community)
  incdist_df <- incdist_raw %>% filter(geo   %in% community_all)
  
  
  #'----------------------------------------------------------------------------
  # HOUSING ----
  #'----------------------------------------------------------------------------
  
  cat("  Loading housing data...\n")
  
  cmhc_scss_df <- cmhc_scss_raw %>% filter(csd %in% community_all)
  
  #' Dwelling type and age (combined from Census 2021)
  dw_df <- rbind(
    census21 %>% filter(DGUID %in% community_all, VARID %in% dtype_varid) %>% left_join(flag21, by = c("VARID" = "varid")),
    census21 %>% filter(DGUID %in% community_all, VARID %in% dage_varid)  %>% left_join(flag21, by = c("VARID" = "varid"))
  )
  
  dw_ten_df <- hhtype_main_raw %>%
    filter(dguid %in% community_all, hhtype == "Total", hhmain == "Total") %>%
    select(-hhtype, -hhmain) %>%
    rename(char = dtype)
  
  #' Affordability (HART federal format, 2016 and 2021)
  dw_aff_df   <- hart_ft_df %>% filter(dguid %in% community_all)
  dw_aff16_df <- hart16_df  %>% filter(geo   %in% community_all)
  
  #' Non-market housing (CSD, CD, provincial rollup)
  nm_csd_df <- nm_raw_df %>%
    filter(csd_code %in% community) %>%
    group_by(housing_type) %>%
    summarize(beds = sum(number_of_beds, na.rm = TRUE), units = sum(number_of_units, na.rm = TRUE)) %>%
    as.data.table() %>% mutate(dguid = community)
  
  nm_cd_df <- nm_raw_df %>%
    mutate(cd = case_when(
      county == "Kings"  ~ 1101,
      county == "Queens" ~ 1102,
      county == "Prince" ~ 1103,
      TRUE ~ NA_real_
    )) %>%
    filter(cd %in% community_cd) %>%
    group_by(housing_type) %>%
    summarize(beds = sum(number_of_beds, na.rm = TRUE), units = sum(number_of_units, na.rm = TRUE)) %>%
    as.data.table() %>% mutate(dguid = community_cd)
  
  nm_pr_df <- nm_raw_df %>%
    group_by(housing_type) %>%
    summarize(beds = sum(number_of_beds, na.rm = TRUE), units = sum(number_of_units, na.rm = TRUE)) %>%
    as.data.table() %>% mutate(dguid = 11)
  
  nm_df <- rbind(nm_csd_df, nm_cd_df, nm_pr_df) %>%
    mutate(
      dguid        = as.numeric(dguid),
      housing_type = ifelse(housing_type == "", "Other", housing_type)
    )
  rm(nm_csd_df, nm_cd_df, nm_pr_df)
  
  #' Unhoused / point-in-time counts (PEI CD-level hardcoded values)
  unhoused_df <- data.table(
    dguid    = c(1101, 1102, 1103),
    unhoused = c(13,   137,  60),
    pr       = c(11,   11,   11)
  )
  
  #' Emergency shelter capacity (beds and shelter counts)
  shelter_raw <- shelters_raw_all %>%
    filter(dguid %in% community) %>%
    mutate(shelter_type = paste(shelter_type, target, sep = " - ")) %>%
    select(-target)
  
  shelter_df <- rbind(
    shelter_raw %>% select(-shelter) %>% pivot_wider(names_from = ref_date, values_from = beds)    %>% mutate(category = "beds",    .after = shelter_type),
    shelter_raw %>% select(-beds)    %>% pivot_wider(names_from = ref_date, values_from = shelter) %>% mutate(category = "shelter", .after = shelter_type) %>% as.data.table()
  )
  rm(shelter_raw)
  
  #' Collective dwellings, STRs, transit (provincial/CSD-level)
  dw_coll_df <- dw_coll_raw %>% filter(dguid %in% community_pr)
  dw_str_df  <- dw_str_raw  %>% filter(dguid %in% community_all)
  transit_df <- transit_raw %>% filter(dguid %in% community_all)
  
  
  #'----------------------------------------------------------------------------
  # INVESTMENT ----
  #'----------------------------------------------------------------------------
  
  cat("  Loading investment / property data...\n")
  
  prop_res_df <- prop_res_raw %>% filter(dguid %in% community)
  prop_own_df <- prop_own_raw %>% filter(dguid %in% community)
  
  
  #'----------------------------------------------------------------------------
  # HOUSING NEED ----
  #'----------------------------------------------------------------------------
  
  cat("  Loading housing need data...\n")
  
  chn_census_df <- census21 %>%
    filter(DGUID %in% community_all, VARID %in% chn_census_varid) %>%
    left_join(flag21, by = c("VARID" = "varid"))
  
  chn_hart_df <- hart_df %>%
    filter(
      geo %in% community_all, shelco == 1,
      indic %in% c(1:7, 9:12),
      char  %in% c(1, 2, 4, 9:11, 17, 19, 20, 22, 24, 26:27, 31, 46:50, 54, 56:61, 63:65)
    ) %>%
    arrange(char, indic) %>%
    select(-shelco) %>%
    left_join(char, by = c("char"  = "characteristics")) %>%
    left_join(ind,  by = c("indic" = "chn")) %>%
    mutate(char = desc_char, indic = desc_chn) %>%
    select(-desc_char, -desc_chn)
  
  #' STIR / Core Housing Need breakdown
  #' Helper: filter STIR, optionally convert a field to a rate, pivot by tenure
  build_stir <- function(data, hhmain_val, drop_cols, rate_col = NULL, rate_denom = NULL, cat_label, value_col) {
    out <- data %>% filter(HHMAIN == hhmain_val, DGUID %in% community_all) %>% select(-all_of(drop_cols))
    if (!is.null(rate_col)) out <- out %>% mutate(!!rate_col := .data[[rate_col]] / .data[[rate_denom]])
    out %>%
      select(-HHMAIN, -any_of(setdiff(c(rate_col, rate_denom), value_col))) %>%
      pivot_wider(names_from = TENURE, values_from = all_of(value_col)) %>%
      as.data.table() %>%
      mutate(category = cat_label, .before = VARID)
  }
  
  stir_ppl_df  <- build_stir(stir_df, "Total", c("CHN_EXAM","CHN"), "STIR_BD", "TOTAL", "Total",        "TOTAL")
  stir_tot_df  <- build_stir(stir_df, "PHM",   c("CHN_EXAM","CHN"), "STIR_BD", "TOTAL", "Total",        "TOTAL")
  stir_aff_df  <- build_stir(stir_df, "PHM",   c("CHN_EXAM","CHN"), "STIR_BD", "TOTAL", "Unaffordable", "STIR_BD")
  stir_exam_df <- build_stir(stir_df, "PHM",   c("TOTAL","STIR_BD"), "CHN",    "CHN_EXAM", "Examined",  "CHN_EXAM")
  stir_chn_df  <- build_stir(stir_df, "PHM",   c("TOTAL","STIR_BD"), "CHN",    "CHN_EXAM", "Core need",  "CHN")
  
  stir_all_df <- rbind(stir_tot_df, stir_aff_df, stir_exam_df, stir_chn_df)
  rm(stir_tot_df, stir_aff_df, stir_exam_df, stir_chn_df)
  
  #' Condition and suitability (inadequate / unsuitable housing rates)
  cond_df <- cond_df_raw %>%
    filter(HHMAIN == "PHM", DGUID %in% community_all) %>%
    mutate(SUIT_BD = SUIT_BD / TOTAL, COND_BD = COND_BD / TOTAL)
  
  cond_all_df <- rbind(
    cond_df %>% select(-HHMAIN, -SUIT_BD, -COND_BD) %>% pivot_wider(names_from = TENURE, values_from = TOTAL)   %>% as.data.table() %>% mutate(category = "Total",     .before = VARID),
    cond_df %>% select(-HHMAIN, -TOTAL,   -COND_BD) %>% pivot_wider(names_from = TENURE, values_from = SUIT_BD) %>% as.data.table() %>% mutate(category = "Unsuitable", .before = VARID),
    cond_df %>% select(-HHMAIN, -TOTAL,   -SUIT_BD) %>% pivot_wider(names_from = TENURE, values_from = COND_BD) %>% as.data.table() %>% mutate(category = "Inadequate", .before = VARID)
  )
  rm(cond_df)
  
  #' Extreme core housing need
  echn_df <- echn_raw %>% filter(dguid %in% community_all)
  
  
  #'----------------------------------------------------------------------------
  # SAVE WORKBOOK ----
  #'----------------------------------------------------------------------------
  
  cat("  Writing Excel workbook...\n")
  
  output_path <- file.path(output_dir, sprintf("HNABlueprint_databook_%s.xlsx", community))
  
  #' Clone the in-memory template rather than re-reading from disk each iteration.
  #' R's copy-on-modify semantics make this a fast in-memory operation.
  wb <- wb_template
  
  sheets <- list(
    "df_geo"          = geo_df,
    "df_wfh"          = wfh_df,
    "df_commute"      = commute_df,
    "df_commute_infl" = commute_infl_df,
    "df_lf"           = lf_df,
    "df_naics"        = naics_df,
    "df_hh_ten"       = hh_ten_df,
    "df_pop"          = pop_df,
    "df_est_csd"      = est_csd_df,
    "df_est_cd"       = est_cd_df,
    "df_est_pr"       = est_pr_df,
    "df_comp_cd"      = comp_cd_df,
    "df_disa"         = disa_hr_df,
    "df_proj"         = proj_df,
    "df_proj_in"      = proj_in_df,
    "df_proj_out"     = proj_out_df,
    "df_comm_latent"  = comm_df,
    "df_census_misc"  = census_misc_df,
    "df_main"         = main_df,
    "df_cost_index"   = cost_index_df,
    "df_cmhc_rms"     = cmhc_rms_df,
    "df_avg_rent"     = avg_rent_df,
    "df_sales"        = sales_df,
    "df_cimd"         = cimd_df,
    "df_cisr"         = cisr_df,
    "df_cisv"         = cisv_df,
    "df_msdi"         = msdi_df,
    "df_permits"      = permits_df,
    "df_wage"         = wage_df,
    "df_inc_char"     = inc_char_df,
    "df_inc_source"   = inc_source_df,
    "df_inc_hart"     = inc_hart_df,
    "df_inc_cat"      = inc_cat_df,
    "df_lwage"        = lwage_df,
    "df_incdist"      = incdist_df,
    "df_mbm"          = mbm_df,
    "df_dwell_aff"    = dw_aff_df,
    "df_dwell16_aff"  = dw_aff16_df,
    "df_cmhc_scss"    = cmhc_scss_df,
    "df_dwell"        = dw_df,
    "df_dwell_ten"    = dw_ten_df,
    "df_dwell_coll"   = dw_coll_df,
    "df_dwell_str"    = dw_str_df,
    "df_nm"           = nm_df,
    "df_shelter"      = shelter_df,
    "df_unhoused"     = unhoused_df,
    "df_transit"      = transit_df,
    "df_prop_res"     = prop_res_df,
    "df_prop_own"     = prop_own_df,
    "df_chn_census"   = chn_census_df,
    "df_chn_hart"     = chn_hart_df,
    "df_stir_ppl"     = stir_ppl_df,
    "df_stir"         = stir_all_df,
    "df_cond"         = cond_all_df,
    "df_echn"         = echn_df,
    "df_flood"        = flood_df,
    "df_coastal"      = coastal_df,
    "df_humid"        = humid_df,
    "df_hot"          = hot_df,
    "df_cold"         = cold_df,
    "df_rain"         = rain_df,
    "df_clim_vuln"    = clim_vuln_df
  )
  
  #' Remove any pre-existing data sheets, then re-add as hidden so they
  #' feed Excel formula references without being visible to users.
  #' clean_utf8() is applied to every sheet to prevent illegal XML character
  #' errors from accented place names in French Canadian geographies.
  sheets_to_remove <- intersect(names(sheets), wb$sheet_names)
  for (s in sheets_to_remove) wb$remove_worksheet(s)
  
  for (sheet_name in names(sheets)) {
    wb$add_worksheet(sheet_name, visible = "hidden")
    wb$add_data(sheet = sheet_name, x = clean_utf8(sheets[[sheet_name]]))
  }
  
  wb$save(output_path)
  cat(sprintf("  Saved: %s\n", output_path))
  
  
  #'----------------------------------------------------------------------------
  # CLEAN UP (per-CSD objects only) ----
  #'----------------------------------------------------------------------------
  
  #' Remove CSD-specific objects before the next iteration to free memory.
  rm(
    geo_df, commute_in_all, commute_out_all, commute_df, commute_in_csds,
    commute_infl_df,
    wfh_df, lf_df, naics_df,
    pop_df, imm_vis_df, est_csd_df, est_cd_df, est_pr_df, comp_cd_df, disa_hr_df, census_misc_df,
    proj_df, proj_in_df, proj_out_df, comm_df,
    main_df, hh_ten_df,
    cost_index_df, tuition_index_df, avg_rent_df, cmhc_rms_df, sales_df,
    cimd_df, cisr_df, cisv_df, msdi_df, flood_df, coastal_df, humid_df,
    hot_df, cold_df, rain_df, clim_vuln_df,
    permits_df, inc_source_df, inc_char_df, inc_hart_df, inc_cat_df,
    mbm_df, wage_df, lwage_df, incdist_df,
    dw_aff_df, dw_aff16_df, cmhc_scss_df, dw_df, dw_ten_df,
    nm_df, shelter_df, unhoused_df, dw_coll_df, dw_str_df, transit_df,
    prop_res_df, prop_own_df,
    chn_census_df, chn_hart_df, stir_ppl_df, stir_all_df, cond_all_df, echn_df,
    sheets, wb
  )
  
  elapsed <- round(proc.time()[["elapsed"]] - csd_start, 1)
  cat(sprintf("  CSD %s complete. (%.1f seconds)\n", community, elapsed))
}

cat("\nAll CSDs processed.\n")