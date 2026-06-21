# ============================================================
# Project: CSI Daily Monitor
# Author: Portfolio version
# Purpose:
#   - Daily CSI monitoring (branch & agent level)
#   - Fair KPI comparison (today vs yesterday vs cumulative)
#   - Action-oriented metrics (response rate, inducible cases)
#   - Internal management & decision support
# Status:
#   Stable version (v1.4)
#2026/06/12:Enhanced data base access security
# ============================================================

library(readxl)
library(dplyr)
library(lubridate)
library(ggplot2)
library(gt)
library(scales)
library(htmltools)
library(webshot2)
library(pdftools)
library(pagedown)
library(forecast)
library(tidyr)
library(forcats)
library(psych)
library(ltm)
library(fmsb)
library(plotly)
library(prophet)
library(writexl)
library(stringr)
library(DBI)
library(RMariaDB)


get_script_dir <- function() {
  if (requireNamespace("rstudioapi", quietly = TRUE) && rstudioapi::isAvailable()) {
    path <- tryCatch(rstudioapi::getSourceEditorContext()$path, error = function(e) "")
    if (nzchar(path)) {
      return(dirname(normalizePath(path, winslash = "/", mustWork = FALSE)))
    }
  }

  file_arg <- grep("^--file=", commandArgs(trailingOnly = FALSE), value = TRUE)
  if (length(file_arg) > 0) {
    path <- sub("^--file=", "", file_arg[[1]])
    return(dirname(normalizePath(path, winslash = "/", mustWork = FALSE)))
  }

  getwd()
}

output_dir <- Sys.getenv("REPORT_OUTPUT_DIR", file.path(get_script_dir(), "outputs"))
dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)

get_required_env <- function(name) {
  value <- Sys.getenv(name, unset = NA_character_)
  if (is.na(value) || !nzchar(value)) {
    stop(sprintf("Missing required environment variable: %s", name), call. = FALSE)
  }
  value
}

load_daily_monitor_data <- function(period_df, period_df2) {
  mariadb_port <- as.integer(Sys.getenv("MARIADB_PORT", "3306"))
  if (is.na(mariadb_port)) {
    stop("MARIADB_PORT must be a valid integer", call. = FALSE)
  }

  table_name <- Sys.getenv("MARIADB_TABLE", "csi_feedback_clean")
  if (!str_detect(table_name, "^[A-Za-z0-9_]+$")) {
    stop("MARIADB_TABLE may contain only letters, numbers, and underscores", call. = FALSE)
  }

  con <- DBI::dbConnect(
    RMariaDB::MariaDB(),
    host = Sys.getenv("MARIADB_HOST", "127.0.0.1"),
    port = mariadb_port,
    user = get_required_env("MARIADB_USER"),
    password = get_required_env("MARIADB_PASSWORD"),
    dbname = Sys.getenv("MARIADB_DATABASE", "customer_experience")
  )

  on.exit(DBI::dbDisconnect(con), add = TRUE)

  df <- DBI::dbGetQuery(con, sprintf("
SELECT *
FROM `%s`
WHERE 填寫問卷時間 IS NOT NULL
  AND DATE(填寫問卷時間) BETWEEN '%s' AND '%s'
", table_name, period_df$df_start, period_df$df_end))

  df2 <- DBI::dbGetQuery(con, sprintf("
SELECT *
FROM `%s`
WHERE DATE(簡訊發送) BETWEEN '%s' AND '%s'

UNION ALL

SELECT *
FROM `%s`
WHERE DATE(LINE發送) BETWEEN '%s' AND '%s'
", table_name, period_df2$df2_start, period_df2$df2_end,
   table_name, period_df2$df2_start, period_df2$df2_end))

  list(df = df, df2 = df2)
}

#Date interval function
get_csi_period <- function(report_ym, run_date = Sys.Date()) {

  report_month <- lubridate::ymd(paste0(report_ym, "-01"))

  # df：CSI填寫問卷期間
  df_start <- report_month %m-% months(1) + days(20)   # n-1月/21
  df_end_limit <- report_month + days(19)              # n月/20
  df_end_run <- as.Date(run_date) - days(1)            # 當日-1

  df_end <- min(df_end_limit, df_end_run)

  # df2：回函率發送期間
  df2_start <- report_month                            # n月/1
  df2_end_limit <- ceiling_date(report_month, "month") - days(1)  # n月最後一天
  df2_end_run <- as.Date(run_date) - days(1)           # 當日-1

  df2_end <- min(df2_end_limit, df2_end_run)

  list(
    df_start = df_start,
    df_end = df_end,
    df2_start = df2_start,
    df2_end = df2_end
  )
}

#Data Acquiry
# 報表月份：CSI填寫問卷用
report_ym <- Sys.getenv("REPORT_YM", format(Sys.Date(), "%Y-%m"))
run_date <- Sys.Date()

period_df <- get_csi_period(report_ym, run_date = run_date)

current_ym <- format(run_date, "%Y-%m")
period_df2 <- get_csi_period(current_ym, run_date = run_date)

data_list <- load_daily_monitor_data(period_df, period_df2)

df <- data_list$df
df2 <- data_list$df2
# === 2. 時間格式處理 ===
df <- df %>%
  mutate(填寫問卷時間 = ymd_hms(填寫問卷時間))
df2 <- df2 %>%
  mutate(
    簡訊發送 = ymd_hms(`簡訊發送`),
    LINE發送 = ymd_hms(`LINE發送`),
    填寫問卷時間 = ymd_hms(填寫問卷時間))
df2 <- df2 %>%
  mutate(
    是否回覆 = ifelse(!is.na(填寫問卷時間), 1, 0),
    可加強勸誘 = ifelse(is.na(填寫問卷時間) & is.na(轉回CSI專案日期), 1, 0)
  )
df2_summary <- df2 %>%
  group_by(據點名稱) %>%
  summarise(
    發送數 = n(),
    回覆數 = sum(是否回覆 == 1, na.rm = TRUE),
    回函率 = 回覆數 / 發送數,
    可勸誘數 = sum(可加強勸誘 == 1, na.rm = TRUE),
    可勸誘比例 = 可勸誘數 / 發送數
  ) %>%
  arrange(desc(可勸誘比例))
df2_total <- df2 %>%
  summarise(
    據點名稱 = "合計",
    發送數 = n(),
    回覆數 = sum(是否回覆 == 1, na.rm = TRUE),
    回函率 = 回覆數 / 發送數,
    可勸誘數 = sum(可加強勸誘 == 1, na.rm = TRUE),
    可勸誘比例 = 可勸誘數 / 發送數
  )

df2_summary <- bind_rows(df2_summary, df2_total)
df2_SA_summary <- df2 %>%
  mutate(
    服務專員姓名 = trimws(as.character(服務專員姓名)),
    據點名稱 = as.character(據點名稱)
  ) %>%
  group_by(據點名稱, 服務專員姓名) %>%
  summarise(
    發送數 = n(),
    回覆數 = sum(是否回覆 == 1, na.rm = TRUE),
    回函率 = 回覆數 / 發送數,
    可勸誘數 = sum(可加強勸誘 == 1, na.rm = TRUE),
    可勸誘比例 = 可勸誘數 / 發送數,
    .groups = "drop"
  ) %>%
  arrange(desc(可勸誘比例))

# === 3. 分離「今日累計」與「昨日累計」資料 ===
latest_date <- as_date(max(df$填寫問卷時間, na.rm = TRUE))
yesterday_date <- latest_date - days(1)

# 今日累計（截至今日為止）
df_today <- df %>%
  filter(as_date(填寫問卷時間) <= latest_date)

# 昨日累計（截至昨日為止）
df_yesterday_raw <- df %>%
  filter(as_date(填寫問卷時間) <= yesterday_date)

# === 4. 統計今日與昨日資料 ===
summary_csi <- function(data) {
  data %>%
    group_by(據點名稱) %>%
    summarise(
      樣本數 = n(),
      落實件數 = sum(`A+落實` == 1, na.rm = TRUE),
      高滿意件數 = sum(高滿意 == 1, na.rm = TRUE),
      低滿意件數 = sum(低滿意 == 1, na.rm = TRUE)
    ) %>%
    mutate(
      落實比例 = 落實件數 / 樣本數,
      滿意比例 = 高滿意件數 / 樣本數,
      低滿意比例 = 低滿意件數 / 樣本數
    )
}

df_today_summary <- summary_csi(df_today)
df_yesterday_summary <- summary_csi(df_yesterday_raw)
today_total <- df_today_summary %>%
  summarise(
    據點名稱 = "合計",
    樣本數 = sum(樣本數, na.rm = TRUE),
    落實件數 = sum(落實件數, na.rm = TRUE),
    高滿意件數 = sum(高滿意件數, na.rm = TRUE),
    低滿意件數 = sum(低滿意件數, na.rm = TRUE)
  ) %>%
  mutate(
    落實比例 = 落實件數 / 樣本數,
    滿意比例 = 高滿意件數 / 樣本數,
    低滿意比例 = 低滿意件數 / 樣本數
  )

yesterday_total <- df_yesterday_summary %>%
  summarise(
    據點名稱 = "合計",
    樣本數 = sum(樣本數, na.rm = TRUE),
    落實件數 = sum(落實件數, na.rm = TRUE),
    高滿意件數 = sum(高滿意件數, na.rm = TRUE),
    低滿意件數 = sum(低滿意件數, na.rm = TRUE)
  ) %>%
  mutate(
    落實比例 = 落實件數 / 樣本數,
    滿意比例 = 高滿意件數 / 樣本數,
    低滿意比例 = 低滿意件數 / 樣本數
  )

df_today_summary <- bind_rows(df_today_summary, today_total)
df_yesterday_summary <- bind_rows(df_yesterday_summary, yesterday_total)

# === 5. 強制據點排序 ===
site_order <- c("Branch 01", "Branch 02", "Branch 03", "Branch 04", "Branch 05", "合計")

df_today_summary$據點名稱 <- factor(
  df_today_summary$據點名稱,
  levels = site_order
)
df_today_summary <- arrange(df_today_summary, 據點名稱)

df_yesterday_summary$據點名稱 <- factor(
  df_yesterday_summary$據點名稱,
  levels = site_order
)
df_yesterday_summary <- arrange(df_yesterday_summary, 據點名稱)

# === 6. 比對差異（主數據為今日，箭頭比較昨日） ===
df_final <- df_today_summary %>%
  left_join(df_yesterday_summary, by = "據點名稱", suffix = c("", "_昨")) %>%
  mutate(
    # 計算差異
    落實_diff = 落實比例 - 落實比例_昨,
    滿意_diff = 滿意比例 - 滿意比例_昨,
    低滿意_diff = 低滿意比例 - 低滿意比例_昨,

    # 顯示用欄位（主數據為今日、箭頭變化量為昨日比較）
    落實顯示 = sprintf(
      "<span style='color:%s;font-weight:bold;'>%s</span> <span style='font-size:10pt;color:#777;'>%s %+0.1f%%</span>",
      ifelse(落實比例 >= 0.97, '#00B050', '#E74C3C'),
      percent(落實比例, accuracy = 0.1),
      ifelse(落實_diff > 0, '▲', ifelse(落實_diff < 0, '▼', '—')),
      落實_diff * 100
    ),
    滿意顯示 = sprintf(
      "<span style='color:%s;font-weight:bold;'>%s</span> <span style='font-size:10pt;color:#777;'>%s %+0.1f%%</span>",
      ifelse(滿意比例 >= 0.97, '#00B050', '#E74C3C'),
      percent(滿意比例, accuracy = 0.1),
      ifelse(滿意_diff > 0, '▲', ifelse(滿意_diff < 0, '▼', '—')),
      滿意_diff * 100
    ),
    低滿意顯示 = sprintf(
      "<span style='color:%s;font-weight:bold;'>%s</span> <span style='font-size:10pt;color:#777;'>%s %+0.1f%%</span>",
      ifelse(低滿意比例 <= 0.01, '#00B050', '#E74C3C'),
      percent(低滿意比例, accuracy = 0.1),
      ifelse(低滿意_diff > 0, '▲', ifelse(低滿意_diff < 0, '▼', '—')),
      低滿意_diff * 100
    )
  )

# === 將 df2_summary（回函率 + 可勸誘數）整合進 df_final ===
df_final <- df_final %>%
  left_join(df2_summary, by = "據點名稱") %>%
  mutate(
    回函率顏色 = ifelse(回函率 < 0.3, "#E74C3C",
                   ifelse(回函率 < 0.5, "#F39C12", "#00B050")),
    勸誘顏色 = ifelse(可勸誘比例 > 0.4, "#E74C3C",
                  ifelse(可勸誘比例 > 0.2, "#F39C12", "#00B050"))
  )

report_month_text <- paste0(as.integer(format(lubridate::ymd(paste0(report_ym, "-01")), "%m")), "月份")

# === 7. 建立報表 ===
report <- df_final %>%
  dplyr::select(
    據點名稱,
    樣本數,
    落實件數, 落實顯示,
    高滿意件數, 滿意顯示,
    低滿意件數, 低滿意顯示
  ) %>%
  gt(rowname_col = "據點名稱") %>%
  tab_header(
    title = md("**CSI 每日戰報**"),
    subtitle = paste0("• ", format(Sys.Date(), "%Y/%m/%d"))
  ) %>%
  fmt_markdown(columns = c(落實顯示, 滿意顯示, 低滿意顯示)) %>%
  cols_label(
    樣本數 = md("樣本數"),
    落實件數 = md("100%件數"),
    落實顯示 = md("A⁺ 落實度"),
    高滿意件數 = md("高滿意件數"),
    滿意顯示 = md("滿意度"),
    低滿意件數 = md("低滿意件數"),
    低滿意顯示 = md("低滿意比例")
  ) %>%
  tab_options(
    table.font.names = "Microsoft JhengHei",
    data_row.padding = px(5)
  )

# === 9. 視覺樣式 ===
htmltools::browsable(
  tagList(
    tags$style(HTML("
      body {
        background: linear-gradient(to bottom, #ffffff, #fff7d9);
        font-family: 'Microsoft JhengHei';
      }
      table.gt_table {
        border-collapse: collapse;
        border-radius: 12px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.15);
      }
      .gt_title {
        font-weight: 900;
        color: #000;
        letter-spacing: 1px;
      }
      .gt_subtitle {
        color: #555;
        font-weight: bold;
      }
      .card .label {
    color: #000000;     /* 黑色 */
    font-weight: 700;   /* 稍微加粗，像 KPI 標籤 */
      }

    ")),
    report
  )
)

# === 代號與顏色設定 ===
branch_codes  <- c("B01", "B02", "B03", "B04", "B05", "Total")
branch_names  <- c("Branch 01", "Branch 02", "Branch 03", "Branch 04", "Branch 05", "合計")
branch_colors <- c("#001F60", "#A93226", "#1E8449", "#884EA0", "#2471A3", "#2C3E50")
card_order <- c("Branch 01", "Branch 02", "Branch 03", "Branch 04", "Branch 05", "合計")

# === 生成單張卡 ===
generate_card <- function(code, branch, data_row, color) {

  # ===== 顏色條件設定 =====
  color_lu <- ifelse(data_row$落實比例 >= 0.98, "#00B050",
                     ifelse(data_row$落實比例 < 0.97, "#C00000", "#000000"))
  color_sat <- ifelse(data_row$滿意比例 >= 0.98, "#00B050",
                      ifelse(data_row$滿意比例 < 0.97, "#C00000", "#000000"))
  color_low <- ifelse(data_row$低滿意比例 > 0.01, "#C00000",
                      ifelse(data_row$低滿意比例 == 0, "#00B050", "#000000"))

  # ===== 變化率顏色 =====
  diff_color_lu  <- ifelse(data_row$落實_diff > 0, "#00B050",
                           ifelse(data_row$落實_diff < 0, "#E74C3C", "#777"))
  diff_color_sat <- ifelse(data_row$滿意_diff > 0, "#00B050",
                           ifelse(data_row$滿意_diff < 0, "#E74C3C", "#777"))
  diff_color_low <- ifelse(data_row$低滿意_diff > 0, "#E74C3C",
                           ifelse(data_row$低滿意_diff < 0, "#00B050", "#777"))

  # ===== 回函率顏色邏輯 =====
  rate_color <- ifelse(is.na(data_row$回函率), "#777",
                       ifelse(data_row$回函率 < 0.4, "#E74C3C",
                              ifelse(data_row$回函率 < 0.45, "#000000", "#00B050")))

  # ===== 防呆：把所有 NA → 0 並轉 numeric =====
  safe_num <- function(x) {
    ifelse(is.na(as.numeric(x)), 0, as.numeric(x))
  }

  # ===== 組 HTML =====
  sprintf("
  <div class='card' style='border-left:6px solid %s;'>
    <div class='branch'>
      <span class='branch-code'>%s</span> %s
    </div>
    <div class='stats'>

      <div class='stat'>
        <div class='label'>樣本數</div>
        <div class='value'>%d</div>
      </div>

      <div class='stat'>
        <div class='label'>A⁺ 落實度</div>
        <div class='value' style='color:%s;'>%.1f%%</div>
        <span class='diff' style='color:%s;'>%s %+.1f%%</span>
        <div class='count'>(%d件)</div>
      </div>

      <div class='stat'>
        <div class='label'>滿意度</div>
        <div class='value' style='color:%s;'>%.1f%%</div>
        <span class='diff' style='color:%s;'>%s %+.1f%%</span>
        <div class='count'>(%d件)</div>
      </div>

      <div class='stat'>
        <div class='label'>低滿意比例</div>
        <div class='value' style='color:%s;'>%.1f%%</div>
        <span class='diff' style='color:%s;'>%s %+.1f%%</span>
        <div class='count'>(%d件)</div>
      </div>

      <div class='stat'>
        <div class='label'>回函率</div>
        <div class='value' style='color:%s;'>%s</div>
        <div class='count'>(%d件)</div>
      </div>

      <div class='stat'>
        <div class='label'>可加強勸誘</div>
        <div class='value' style='color:#1E8449;'>%d件</div>
        <div class='count'>(%.1f%%)</div>
      </div>

    </div>
  </div>",
          # === 主內容 ===
          color, code, branch,
          safe_num(data_row$樣本數),
          # 落實度
          color_lu, safe_num(data_row$落實比例 * 100), diff_color_lu,
          ifelse(data_row$落實_diff > 0, "▲", ifelse(data_row$落實_diff < 0, "▼", "▬")),
          safe_num(data_row$落實_diff * 100), safe_num(data_row$落實件數),
          # 滿意度
          color_sat, safe_num(data_row$滿意比例 * 100), diff_color_sat,
          ifelse(data_row$滿意_diff > 0, "▲", ifelse(data_row$滿意_diff < 0, "▼", "▬")),
          safe_num(data_row$滿意_diff * 100), safe_num(data_row$高滿意件數),
          # 低滿意
          color_low, safe_num(data_row$低滿意比例 * 100), diff_color_low,
          ifelse(data_row$低滿意_diff > 0, "▲", ifelse(data_row$低滿意_diff < 0, "▼", "▬")),
          safe_num(data_row$低滿意_diff * 100), safe_num(data_row$低滿意件數),
          # 回函率與勸誘
          rate_color, scales::percent(safe_num(data_row$回函率), accuracy = 0.1),
          safe_num(data_row$回覆數),
          safe_num(data_row$可勸誘數), safe_num(data_row$可勸誘比例 * 100)
  )
}

# === 依固定順序抓每一列資料 ===
data_rows <- lapply(card_order, function(site) {
  df_final %>% filter(as.character(據點名稱) == site)
})

# === 組合全部卡片 ===
cards_html <- paste0(
  mapply(
    generate_card,
    branch_codes,
    branch_names,
    data_rows,
    branch_colors,
    SIMPLIFY = FALSE
  ),
  collapse = "\n"
)

html_content <- sprintf("
<!DOCTYPE html>
<html lang='zh-Hant'>
<head>
<meta charset='UTF-8'>
<title>%s CSI 每日戰報</title>
<style>
  body {
    font-family: 'Microsoft JhengHei', sans-serif;
    background: linear-gradient(to bottom, #ffffff, #fff7d9);
    padding: 15px 25px;
    margin: 0;
  }

  /* ===== 標題區 ===== */
  .header {
  text-align: left;
  position: relative;
  margin-bottom: 25px;
  padding-left: 25px;
}

.report-title {
  font-size: 36px;
  font-weight: 900;
  letter-spacing: 1px;
  display: inline-block;
}

  .report-date {
  position: absolute;
  right: 25px;
  top: 10px;
  font-size: 16px;
  font-weight: 600;
  color: #666;
  text-align: right;
  line-height: 1.5;
}

.period-text {
  font-size: 13px;
  font-weight: 500;
  color: #777;
}

  /* ===== 卡片樣式 ===== */
  .card {
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #fff;
    border-radius: 14px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.1);
    padding: 10px 20px;
    margin-bottom: 8px;
  }

  .branch {
  display: flex;
  align-items: center;
  gap: 12px;
  font-size: 20px;
  font-weight: 900;
  flex-shrink: 0;
  width: 150px;
  white-space: nowrap;
}

  .stats {
    display: flex;
    align-items: flex-end;
    justify-content: space-between;
    width: 100%%;
    gap: 32px;
    text-align: center;
  }

  .stat {
    position: relative;
    min-width: 90px;
  }

  /* === 樣本數對齊調整 === */
  .stat:first-child {
    transform: translateY(-15px);
  }

  .label {
    font-size: 11px;
    color: #777;
    margin-bottom: 2px;
  }

  .value {
    font-size: 22px;
    font-weight: bold;
    line-height: 1;
    display: inline-block;
    margin-right: 2px;
  }

  .diff {
    display: inline-block;
    font-size: 12px;
    font-weight: 600;
    margin-left: 3px;
    vertical-align: top;
    position: relative;
    top: -2px;
  }

  .count {
    font-size: 11px;
    color: #777;
    line-height: 1.1;
  }

  /* hover 效果 */
  .card:hover {
    background: #f8f9fa;
    transform: translateY(-1px);
    transition: 0.2s;
  }
</style>
</head>
<body>
  <div class='header'>
    <div class='report-title'>%s CSI 每日戰報</div>
   <div class='report-date'>
  <div>%s</div>
  <div class='period-text'>調查期間：%s ~ %s</div>
</div>
  </div>
  %s
</body>
</html>",
report_month_text,
report_month_text,
format(Sys.Date(), "%Y/%m/%d"),
format(period_df$df_start, "%Y/%m/%d"),
format(period_df$df_end, "%Y/%m/%d"),
cards_html)               # 4️⃣ 報表內容

# === 寫出檔案 ===
file_name <- sprintf("%s CSI_每日戰報_%s.html",
                     report_month_text,
                     format(Sys.Date(), "%Y-%m-%d"))
file_name <- file.path(output_dir, file_name)
writeLines(html_content, file_name)
browseURL(file_name)

# === 1️⃣ 據點 + 專員層級統計 ===
summary_csi_agent <- function(data) {
  data %>%
    group_by(據點名稱, 服務專員姓名) %>%
    summarise(
      樣本數 = n(),
      落實件數 = sum(`A+落實` == 1, na.rm = TRUE),
      高滿意件數 = sum(高滿意 == 1, na.rm = TRUE),
      低滿意件數 = sum(低滿意 == 1, na.rm = TRUE)
    ) %>%
    mutate(
      落實比例 = 落實件數 / 樣本數,
      滿意比例 = 高滿意件數 / 樣本數,
      低滿意比例 = 低滿意件數 / 樣本數
    )
}

# === 2️⃣ 今日 / 昨日彙總 ===
df_agent_today <- summary_csi_agent(df_today)
df_agent_yesterday <- summary_csi_agent(df_yesterday_raw)

# === 3️⃣ 合併 + 計算差異 ===
df_agent_final <- df_agent_today %>%
  left_join(df_agent_yesterday, by = c("據點名稱", "服務專員姓名"), suffix = c("", "_昨")) %>%
  mutate(
    落實_diff = 落實比例 - 落實比例_昨,
    滿意_diff = 滿意比例 - 滿意比例_昨,
    低滿意_diff = 低滿意比例 - 低滿意比例_昨,

    落實顯示 = sprintf(
      "<span style='color:%s;font-weight:bold;'>%s</span>
       <span style='font-size:10pt;color:#777;'>%s %+0.1f%%</span>",
      ifelse(落實比例 >= 0.98, '#00B050',
             ifelse(落實比例 < 0.97, '#E74C3C', '#000000')),
      scales::percent(落實比例, accuracy = 0.1),
      ifelse(落實_diff > 0, '▲', ifelse(落實_diff < 0, '▼', '▬')),
      落實_diff * 100
    ),
    滿意顯示 = sprintf(
      "<span style='color:%s;font-weight:bold;'>%s</span>
       <span style='font-size:10pt;color:#777;'>%s %+0.1f%%</span>",
      ifelse(滿意比例 >= 0.98, '#00B050',
             ifelse(滿意比例 < 0.97, '#E74C3C', '#000000')),
      scales::percent(滿意比例, accuracy = 0.1),
      ifelse(滿意_diff > 0, '▲', ifelse(滿意_diff < 0, '▼', '▬')),
      滿意_diff * 100
    ),
    低滿意顯示 = sprintf(
      "<span style='color:%s;font-weight:bold;'>%s</span>
       <span style='font-size:10pt;color:#777;'>%s %+0.1f%%</span>",
      ifelse(低滿意比例 > 0.01, '#E74C3C',
             ifelse(低滿意比例 == 0, '#00B050', '#000000')),
      scales::percent(低滿意比例, accuracy = 0.1),
      ifelse(低滿意_diff > 0, '▲', ifelse(低滿意_diff < 0, '▼', '▬')),
      低滿意_diff * 100
    )
  )
# === 🔹 計算專員獎勵金 ===
df_agent_final <- df_agent_final %>%
  mutate(
    # 落實獎勵
    落實獎勵 = case_when(
      落實比例 > 0.97 & 落實件數 >= 40 ~ 3000,
      落實比例 > 0.97 & 落實件數 >= 30 ~ 2500,
      落實比例 > 0.97 ~ 2000,
      落實比例 <= 0.95 ~ -2000,
      落實比例 <= 0.97 ~ -1000,
      TRUE ~ 0
    ),
    # 滿意獎勵
    滿意獎勵 = case_when(
      滿意比例 <0.97 ~ -2000,
      滿意比例 >= 0.95 & 高滿意件數 >= 50 ~ 5000,
      滿意比例 >= 0.95 & 高滿意件數 >= 40 ~ 3000,
      滿意比例 >= 0.95 & 高滿意件數 >= 30 ~ 2000,
      TRUE ~ 0
    ) - 低滿意件數 * 2000,
    # 總獎勵
    總獎勵金 = 落實獎勵 + 滿意獎勵
  )

df_sa_score_avg <- df %>%
  mutate(
    服務專員姓名 = trimws(as.character(服務專員姓名)),
    據點名稱 = as.character(據點名稱)
  ) %>%
  group_by(據點名稱, 服務專員姓名) %>%
  summarise(
    滿意度平均分 = mean(滿意度總分, na.rm = TRUE),
    .groups = "drop"
  )

df_agent_final <- df_agent_final %>%
  mutate(
    服務專員姓名 = trimws(as.character(服務專員姓名)),
    據點名稱 = as.character(據點名稱)
  ) %>%
  left_join(
    df2_SA_summary %>%
      rename(
        發送數_回函 = 發送數,
        回覆數_回函 = 回覆數,
        回函率_專員 = 回函率,
        可勸誘數_專員 = 可勸誘數,
        可勸誘比例_專員 = 可勸誘比例
      ),
    by = c("據點名稱", "服務專員姓名")
  )

df_agent_final <- df_agent_final %>%
  left_join(
    df_sa_score_avg,
    by = c("據點名稱", "服務專員姓名")
  )

df_agent_final <- df_agent_final %>%
  mutate(
    回函率獎勵 = case_when(

      #回函率 ≥ 45% → 直接獎勵
      !is.na(回函率_專員) & 回函率_專員 >= 0.45 & 回覆數_回函 >=35 ~ 2000,

      #回函率 ≥ 45% → 直接獎勵
      !is.na(回函率_專員) & 回函率_專員 >= 0.40 & 回覆數_回函 >=35 ~ 1000,

      # 其餘低回函率 → 全扣
      !is.na(回函率_專員) & 回函率_專員 < 0.40 ~ -2000,

      # 其他（40–45% 或無資料）
      TRUE ~ 0
    ),

    # 最終總獎勵（唯一對外指標）
    總獎勵金 = 落實獎勵 + 滿意獎勵 + 回函率獎勵
  )

# === 4️⃣ 輸出表格（據點分組 + 專員列表）===
report_agent <- df_agent_final %>%
  dplyr::select(
    據點名稱,
    服務專員姓名,
    樣本數,
    落實件數, 落實顯示,
    高滿意件數, 滿意顯示,
    低滿意件數, 低滿意顯示,
    發送數_回函, 回覆數_回函, 回函率_專員, 可勸誘數_專員,
    回函率獎勵, 總獎勵金
  ) %>%
  gt(groupname_col = "據點名稱", rowname_col = "服務專員姓名") %>%
  tab_header(
    title = md("**服務專員 CSI 成績戰報**"),
    subtitle = paste0("• ", format(Sys.Date(), "%Y/%m/%d"))
  ) %>%
  fmt_markdown(columns = c(落實顯示, 滿意顯示, 低滿意顯示)) %>%
  fmt_percent(columns = 回函率_專員, decimals = 1) %>%
  fmt_number(columns = c(回函率獎勵, 總獎勵金), decimals = 0, use_seps = TRUE) %>%
  cols_label(
    發送數_回函 = md("回函發送"),
    回覆數_回函 = md("回函回覆"),
    回函率_專員 = md("回函率"),
    可勸誘數_專員 = md("可勸誘"),
    回函率獎勵 = md("回函獎勵"),
    總獎勵金 = md("總獎勵金")
  ) %>%
  data_color(
    columns = c(回函率獎勵, 總獎勵金),
    colors = scales::col_bin(
      bins = c(-Inf, -1, 0, Inf),
      palette = c(
        "#E74C3C",  # 負數 → 紅
        "#7F8C8D",  # 0 → 灰
        "#1F618D"   # 正數 → 藍綠
      )
    )
  )

report_agent

# === 專員報表 HTML ===
agent_cards <- df_agent_final %>%
  mutate(
    # 顏色判斷與符號
    落實色 = ifelse(落實比例 >= 0.98, "#00B050",
                 ifelse(落實比例 < 0.97, "#E74C3C", "#000000")),
    滿意色 = ifelse(滿意比例 >= 0.98, "#00B050",
                 ifelse(滿意比例 < 0.97, "#E74C3C", "#000000")),
    低滿意色 = ifelse(低滿意比例 > 0.01, "#E74C3C",
                  ifelse(低滿意比例 == 0, "#00B050", "#000000")),
    落實符 = ifelse(落實_diff > 0, "▲", ifelse(落實_diff < 0, "▼", "▬")),
    滿意符 = ifelse(滿意_diff > 0, "▲", ifelse(滿意_diff < 0, "▼", "▬")),
    低滿意符 = ifelse(低滿意_diff > 0, "▲", ifelse(低滿意_diff < 0, "▼", "▬"))
  )

# === 排序順序 ===
site_order <- c("Branch 01", "Branch 02", "Branch 03", "Branch 04", "Branch 05")

# === 專員報表資料準備 ===
agent_cards <- df_agent_final %>%
  mutate(
    據點名稱 = factor(據點名稱, levels = site_order),

    # === 原本 CSI 顏色/符號 ===
    落實色 = ifelse(落實比例 >= 0.98, "#00B050",
                 ifelse(落實比例 < 0.97, "#E74C3C", "#000000")),
    滿意色 = ifelse(滿意比例 >= 0.98, "#00B050",
                 ifelse(滿意比例 < 0.97, "#E74C3C", "#000000")),
    低滿意色 = ifelse(低滿意比例 > 0.01, "#E74C3C",
                  ifelse(低滿意比例 == 0, "#00B050", "#000000")),
    落實符 = ifelse(落實_diff > 0, "▲", ifelse(落實_diff < 0, "▼", "▬")),
    滿意符 = ifelse(滿意_diff > 0, "▲", ifelse(滿意_diff < 0, "▼", "▬")),
    低滿意符 = ifelse(低滿意_diff > 0, "▲", ifelse(低滿意_diff < 0, "▼", "▬")),

    # === 回函率顯示（專員） ===
    發送數_回函 = ifelse(is.na(發送數_回函), 0, 發送數_回函),
    回覆數_回函 = ifelse(is.na(回覆數_回函), 0, 回覆數_回函),
    回函率_專員 = ifelse(發送數_回函 > 0, 回覆數_回函 / 發送數_回函, NA_real_),

    回函色 = ifelse(is.na(回函率_專員), "#777777",
                 ifelse(回函率_專員 >= 0.45, "#00B050",
                        ifelse(回函率_專員 <  0.40, "#E74C3C", "#000000"))),

    回函顯示 = ifelse(
      is.na(回函率_專員),
      "—",
      paste0(scales::percent(回函率_專員, accuracy = 0.1),
             " (", 回覆數_回函, "/", 發送數_回函, ")")
    ),

    # === 金額顏色（回函獎勵 / 總獎勵金）===
    回函獎勵色 = ifelse(回函率獎勵 > 0, "#1F618D",
                   ifelse(回函率獎勵 < 0, "#E74C3C", "#7F8C8D")),
    總獎勵色   = ifelse(總獎勵金 > 0, "#002060",
                    ifelse(總獎勵金 < 0, "#E74C3C", "#7F8C8D"))
  ) %>%
  arrange(據點名稱, 服務專員姓名)

# === HTML架構 ===
html_agent_cards <- paste0(
  "<html><head>
  <meta charset='UTF-8'>
  <style>
    body {
      background: linear-gradient(to bottom, #ffffff, #fff7d9);
      font-family: 'Microsoft JhengHei';
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
      background-color: #fff7d9 !important;
    }
    .container {
      width: 95%;
      margin: 20px auto;
    }
    .title {
      text-align: center;
      font-weight: 900;
      font-size: 28px;
      margin-bottom: 5px;
    }
    .subtitle {
      text-align: center;
      color: #555;
      font-size: 13px;
      margin-bottom: 20px;
    }
    .site-header {
  margin-top: 35px;
  margin-bottom: 10px;
  font-size: 22px;
  font-weight: 900;
  color: #002060;
  border-left: 10px solid #002060;
  padding-left: 10px;
}

@media print {
  .site-header-break {
    break-before: page;
    page-break-before: always;
  }
}

    .card {
  background: white;
  border-radius: 12px;
  box-shadow: 0 2px 6px rgba(0,0,0,0.15);
  margin: 10px 0;
  padding: 12px 18px;
  border-left: 6px solid #002060;

  display: grid;
  grid-template-columns: 130px 95px 95px 95px 130px 95px;
  align-items: center;
  column-gap: 18px;
}

.agent {
  font-weight: 700;
  font-size: 18px;
  color: #002060;
  white-space: nowrap;
}

.metric {
  text-align: center;
  min-width: 0;
}

.metric h4 {
  margin: 0 0 4px 0;
  font-size: 13px;
  color: #555;
  white-space: nowrap;
}

.metric span.value {
  display: block;
  font-size: 20px;
  font-weight: 800;
  line-height: 1.1;
  white-space: nowrap;
}

.diff {
  display: block;
  font-size: 12px;
  color: #777;
  white-space: nowrap;
  margin-top: 2px;
}

    /* ✅ 每個廠別自動分頁 */
    .page-break {
      page-break-after: always;
    }
   @media print {
  body {
    zoom: 0.95;
  }

  .container {
    width: 100%;
    margin: 0 auto;
  }

  .card {
    page-break-inside: avoid;
    break-inside: avoid;
    grid-template-columns: 120px 88px 88px 88px 120px 88px;
    column-gap: 18px;
    padding: 14px 20px;
  }

  .agent {
    font-size: 16px;
  }

  .metric h4 {
    font-size: 12px;
  }

  .metric span.value {
    font-size: 16px;
    letter-spacing: 0.5px;
  }

  .diff {
    font-size: 11px;
  }

  .page-break {
    page-break-after: always;
  }
}
  </style></head><body>
  <div class='container'>
    <div class='title'>CSI 每日戰報 — 服務專員</div>
    <div class='subtitle'>• ", format(Sys.Date(), "%Y/%m/%d"), "</div>"
)

# === 依據廠別順序分組 ===
valid_sites <- site_order[site_order %in% unique(as.character(agent_cards$據點名稱))]

for (i in seq_along(valid_sites)) {
  site <- valid_sites[i]
  site_data <- agent_cards %>% filter(據點名稱 == site)
  if (nrow(site_data) == 0) next

  header_class <- if (i == 1) "site-header" else "site-header site-header-break"

  html_agent_cards <- paste0(
    html_agent_cards,
    "<div class='", header_class, "'>", site, "</div>"
  )

  for (j in 1:nrow(site_data)) {
    r <- site_data[j,]

    html_agent_cards <- paste0(
      html_agent_cards,
      sprintf("
      <div class='card'>
        <div class='agent'>%s</div>

        <div class='metric'>
          <h4>A⁺ 落實度</h4>
          <span class='value' style='color:%s;'>%s</span>
          <span class='diff'>%s %+0.1f%%</span>
        </div>

        <div class='metric'>
          <h4>滿意度</h4>
          <span class='value' style='color:%s;'>%s</span>
          <span class='diff'>%s %+0.1f%%</span>
        </div>

        <div class='metric'>
          <h4>低滿意比例</h4>
          <span class='value' style='color:%s;'>%s</span>
          <span class='diff'>%s %+0.1f%%</span>
        </div>

        <div class='metric'>
          <h4>回函率</h4>
          <span class='value' style='color:%s;'>%s</span>
        </div>

        <div class='metric'>
          <h4>總獎勵金</h4>
          <span class='value' style='color:%s;'>$%s</span>
        </div>
      </div>",
              r$服務專員姓名,
              r$落實色, percent(r$落實比例, accuracy = 0.1), r$落實符, r$落實_diff * 100,
              r$滿意色, percent(r$滿意比例, accuracy = 0.1), r$滿意符, r$滿意_diff * 100,
              r$低滿意色, percent(r$低滿意比例, accuracy = 0.1), r$低滿意符, r$低滿意_diff * 100,
              r$回函色, r$回函顯示,
              r$總獎勵色, formatC(r$總獎勵金, format = "d", big.mark = ",")
      )
    )
  }

  total_bonus <- sum(site_data$總獎勵金, na.rm = TRUE)

  html_agent_cards <- paste0(
    html_agent_cards,
    sprintf("
    <div class='card' style='background:#002060;color:white;text-align:right;'>
      <div style='grid-column:1 / span 5;font-weight:900;font-size:18px;'>合計獎勵金</div>
      <div style='font-weight:900;font-size:20px;'>$%s</div>
    </div>",
            formatC(total_bonus, format = "d", big.mark = ","))
  )
}

# === 寫出檔案 ===
html_file <- file.path(output_dir, "CSI_Agent_Daily_Report.html")
writeLines(html_agent_cards, html_file)
browseURL(html_file)

# === 指定輸出 PDF 檔案名稱 ===
pdf_file <- file.path(output_dir, "CSI_服務專員每日實績.pdf")

# === 輸出設定 ===
chrome_print(
  input = html_file,
  output = pdf_file,
  wait = 3
)

cat("✅ 已完成輸出 PDF：", pdf_file, "\n")


df_daily <- df %>%
  mutate(日期 = as_date(填寫問卷時間)) %>%
  group_by(據點名稱, 日期) %>%
  summarise(
    樣本數 = n(),
    落實比例 = mean(`A+落實`, na.rm = TRUE),
    滿意比例 = mean(`高滿意`, na.rm = TRUE),
    低滿意比例 = mean(`低滿意`, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  arrange(據點名稱, 日期)

df_cum <- df_daily %>%
  group_by(據點名稱) %>%
  arrange(日期) %>%
  mutate(
    累計樣本 = cumsum(樣本數),
    累計落實 = cumsum(樣本數 * 落實比例),
    累計高滿意 = cumsum(樣本數 * 滿意比例),
    累計低滿意 = cumsum(樣本數 * 低滿意比例),
    累計落實比例 = 累計落實 / 累計樣本,
    累計滿意比例 = 累計高滿意 / 累計樣本,
    累計低滿意比例 = 累計低滿意 / 累計樣本
  )

# === 定義：累計版 ARIMA 趨勢繪圖函式 ===
plot_arima_trend_cum <- function(df, y_col, title, h = 7, freq = 7, lb_lag = 10,
                                 risk_direction = c("low", "high"),
                                 threshold = NULL) {
  risk_direction <- match.arg(risk_direction)

  out <- list()
  branches <- unique(df$據點名稱)

  for (branch in branches) {
    df_branch <- df %>% filter(據點名稱 == branch) %>% arrange(日期)

    df_branch <- df_branch %>%
      mutate(
        累計樣本 = cumsum(樣本數),
        累計值   = cumsum(.data[[y_col]] * 樣本數),
        累計比例 = ifelse(累計樣本 > 0, 累計值 / 累計樣本, NA_real_)
      ) %>%
      filter(!is.na(累計比例))

    if (nrow(df_branch) < 1) next

    ts_data <- ts(df_branch$累計比例, frequency = freq)
    fit <- auto.arima(ts_data)
    fc  <- forecast(fit, h = h)

    df_forecast <- data.frame(
      日期 = seq(max(df_branch$日期) + days(1), by = "day", length.out = h),
      預測 = as.numeric(fc$mean),
      下界 = as.numeric(fc$lower[, 2]),
      上界 = as.numeric(fc$upper[, 2])
    )

    # === 找風險最高日 ===
    risk_day <- if (risk_direction == "low") {
      df_forecast %>% slice_min(order_by = 預測, n = 1, with_ties = FALSE)
    } else {
      df_forecast %>% slice_max(order_by = 預測, n = 1, with_ties = FALSE)
    }
    is_breach <- if (is.null(threshold)) {
      rep(FALSE, nrow(risk_day))
    } else if (risk_direction == "low") {
      risk_day$預測 < threshold
    } else {
      risk_day$預測 > threshold
    }
    risk_day <- risk_day %>%
      mutate(
        is_breach = is_breach
      )
    risk_day <- risk_day %>%
      mutate(
        risk_label = sprintf(
          "%s 最高風險日\n%s\n預測: %s",
          ifelse(is_breach, "🔴", "🟡"),
          format(日期, "%m/%d"),
          scales::percent(預測, accuracy = 0.1)
        )
      )
    y_limits <- if (risk_direction == "high") {
      threshold_ref <- if (is.null(threshold)) numeric(0) else threshold * 2
      y_max <- max(c(df_branch$累計比例, df_forecast$上界, threshold_ref), na.rm = TRUE)
      c(0, min(1.02, max(0.05, y_max * 1.1)))
    } else {
      c(0.8, 1.02)
    }

    p <- ggplot() +
      geom_line(
        data = df_branch,
        aes(x = 日期, y = 累計比例),
        color = "#002060", linewidth = 1.2
      ) +
      geom_ribbon(
        data = df_forecast,
        aes(x = 日期, ymin = 下界, ymax = 上界),
        fill = "#9EC5FE", alpha = 0.4
      ) +
      geom_line(
        data = df_forecast,
        aes(x = 日期, y = 預測),
        color = "#0056B3", linewidth = 1.2, linetype = "dashed"
      ) +
      geom_point(
        data = risk_day,
        aes(x = 日期, y = 預測),
        color = ifelse(risk_day$is_breach, "#E74C3C", "#F39C12"),
        size = 3
      ) +
      geom_text(
        data = risk_day,
        aes(x = 日期, y = 預測, label = risk_label),
        vjust = -0.8,
        size = 3.5,
        family = "Microsoft JhengHei"
      ) +
      scale_y_continuous(
        labels = percent_format(accuracy = 1)
      ) +
      coord_cartesian(ylim = y_limits) +
      labs(
        title = paste0(branch, " — ", title, "（累計）"),
        subtitle = paste0("ARIMA ", h, "日預測（累計資料 ", nrow(df_branch), " 筆）"),
        x = "日期", y = "累計比例"
      ) +
      theme_minimal(base_family = "Microsoft JhengHei") +
      theme(
        plot.title = element_text(size = 16, face = "bold", color = "#002060"),
        plot.subtitle = element_text(size = 11, color = "#555"),
        axis.title = element_text(size = 12, face = "bold"),
        axis.text = element_text(size = 10)
      )

    if (!is.null(threshold)) {
      p <- p + geom_hline(
        yintercept = threshold,
        color = "#E74C3C",
        linetype = "dotted",
        linewidth = 1
      )
    }

    acc <- accuracy(fit)
    lb  <- Box.test(residuals(fit), lag = lb_lag, type = "Ljung")

    out[[branch]] <- list(
      plot = p,
      fit = fit,
      forecast = df_forecast,
      accuracy = acc,
      ljungbox = lb,
      risk_day = risk_day
    )
  }

  out
}


plots_lu_cum <- plot_arima_trend_cum(df_daily, "落實比例", "A⁺ 落實度趨勢",risk_direction = "low",threshold = 0.97)
plots_sat_cum <- plot_arima_trend_cum(df_daily, "滿意比例", "滿意度趨勢",risk_direction = "low",threshold = 0.97)
plots_low_cum <- plot_arima_trend_cum(df_daily, "低滿意比例", "低滿意比例趨勢",risk_direction = "high",threshold = 0.02)

# 模型明細圖保留在 plots_*_cum 中，需要時再手動 print()。

#"Prophet" Prediction
# === 定義：累計版 Prophet 趨勢繪圖函式 ===
plot_prophet_trend_cum <- function(df, y_col, title, h = 7,
                                   risk_direction = c("low", "high"),
                                   threshold = NULL) {
  risk_direction <- match.arg(risk_direction)

  out <- list()
  branches <- unique(df$據點名稱)

  for (branch in branches) {
    df_branch <- df %>%
      filter(據點名稱 == branch) %>%
      arrange(日期)

    df_branch <- df_branch %>%
      mutate(
        累計樣本 = cumsum(樣本數),
        累計值   = cumsum(.data[[y_col]] * 樣本數),
        累計比例 = ifelse(累計樣本 > 0, 累計值 / 累計樣本, NA_real_)
      ) %>%
      filter(!is.na(累計比例))

    if (nrow(df_branch) < 7) next

    prophet_df <- df_branch %>%
      transmute(
        ds = as.Date(日期),
        y  = as.numeric(累計比例)
      )

    fit <- prophet(
      prophet_df,
      growth = "linear",
      yearly.seasonality = FALSE,
      weekly.seasonality = TRUE,
      daily.seasonality = FALSE,
      seasonality.mode = "additive"
    )

    future <- make_future_dataframe(
      fit,
      periods = h,
      freq = "day"
    )

    forecast_df <- predict(fit, future)

    df_forecast <- forecast_df %>%
      select(ds, yhat, yhat_lower, yhat_upper) %>%
      filter(ds > max(prophet_df$ds)) %>%
      slice(1:h) %>%
      rename(
        日期 = ds,
        預測 = yhat,
        下界 = yhat_lower,
        上界 = yhat_upper
      ) %>%
      mutate(
        預測 = pmin(pmax(預測, 0), 1),
        下界 = pmin(pmax(下界, 0), 1),
        上界 = pmin(pmax(上界, 0), 1)
      )

    df_fitted <- forecast_df %>%
      select(ds, yhat) %>%
      filter(ds <= max(prophet_df$ds)) %>%
      rename(
        日期 = ds,
        fitted = yhat
      )

    eval_df <- prophet_df %>%
      rename(日期 = ds, actual = y) %>%
      left_join(df_fitted, by = "日期")

    rmse <- sqrt(mean((eval_df$actual - eval_df$fitted)^2, na.rm = TRUE))
    mae  <- mean(abs(eval_df$actual - eval_df$fitted), na.rm = TRUE)
mape <- mean(abs((eval_df$actual - eval_df$fitted) / eval_df$actual),
             na.rm = TRUE) * 100

metrics <- data.frame(
  branch = branch,
  RMSE = rmse,
  MAE = mae,
  MAPE = mape
)

# === 找風險最高日 ===
risk_day <- if (risk_direction == "low") {
  df_forecast %>% slice_min(order_by = 預測, n = 1, with_ties = FALSE)
} else {
  df_forecast %>% slice_max(order_by = 預測, n = 1, with_ties = FALSE)
}

is_breach <- if (is.null(threshold)) {
  rep(FALSE, nrow(risk_day))
} else if (risk_direction == "low") {
  risk_day$預測 < threshold
} else {
  risk_day$預測 > threshold
}

risk_day <- risk_day %>%
  mutate(
    is_breach = is_breach,
    risk_label = sprintf(
      "⚠ 最高風險日\n%s\n預測: %s",
      format(日期, "%m/%d"),
      scales::percent(預測, accuracy = 0.1)
    )
  )
y_limits <- if (risk_direction == "high") {
  threshold_ref <- if (is.null(threshold)) numeric(0) else threshold * 2
  y_max <- max(c(df_branch$累計比例, df_forecast$上界, threshold_ref), na.rm = TRUE)
  c(0, min(1.02, max(0.05, y_max * 1.1)))
} else {
  c(0.8, 1.02)
}

p <- ggplot() +
  geom_line(
    data = df_branch,
    aes(x = 日期, y = 累計比例),
    color = "#002060",
    linewidth = 1.2
  ) +
  geom_ribbon(
    data = df_forecast,
    aes(x = 日期, ymin = 下界, ymax = 上界),
    fill = "#9EC5FE",
    alpha = 0.4
  ) +
  geom_line(
    data = df_forecast,
    aes(x = 日期, y = 預測),
    color = "#0056B3",
    linewidth = 1.2,
    linetype = "dashed"
  ) +
  geom_point(
    data = risk_day,
    aes(x = 日期, y = 預測),
    color = ifelse(risk_day$is_breach, "#E74C3C", "#F39C12"),
    size = 3
  ) +
  geom_text(
    data = risk_day,
    aes(x = 日期, y = 預測, label = risk_label),
    vjust = -0.8,
    size = 3.5,
    family = "Microsoft JhengHei"
  ) +
  scale_y_continuous(
    labels = percent_format(accuracy = 1)
  ) +
  coord_cartesian(ylim = y_limits) +
  labs(
    title = paste0(branch, " — ", title, "（累計）"),
    subtitle = paste0("Prophet ", h, "日預測（累計資料 ", nrow(df_branch), " 筆）"),
    x = "日期",
    y = "累計比例"
  ) +
  theme_minimal(base_family = "Microsoft JhengHei") +
  theme(
    plot.title = element_text(size = 16, face = "bold", color = "#002060"),
    plot.subtitle = element_text(size = 11, color = "#555"),
    axis.title = element_text(size = 12, face = "bold"),
    axis.text = element_text(size = 10)
  )

if (!is.null(threshold)) {
  p <- p + geom_hline(
    yintercept = threshold,
    color = "#E74C3C",
    linetype = "dotted",
    linewidth = 1
  )
}

out[[branch]] <- list(
  plot = p,
  fit = fit,
  forecast = df_forecast,
  fitted = df_fitted,
  metrics = metrics,
  risk_day = risk_day
)
  }

  out
}

plots_lu_prophet <- plot_prophet_trend_cum(
  df_daily,
  "落實比例",
  "A⁺ 落實度趨勢",
  risk_direction = "low",
  threshold = 0.97   # ❗一定要有
)
plots_sat_prophet <- plot_prophet_trend_cum(
  df_daily,
  "滿意比例",
  "滿意度趨勢",
  risk_direction = "low",
  threshold = 0.97   # ❗一定要有
)
plots_low_prophet <- plot_prophet_trend_cum(
  df_daily,
  "低滿意比例",
  "低滿意比例趨勢",
  risk_direction = "high",
  threshold = 0.02   # 經銷商門檻，低滿意比例越低越好
)

make_cum_actual_df <- function(df, y_col) {
  df %>%
    arrange(據點名稱, 日期) %>%
    group_by(據點名稱) %>%
    mutate(
      累計樣本 = cumsum(樣本數),
      累計值 = cumsum(.data[[y_col]] * 樣本數),
      實際 = ifelse(累計樣本 > 0, 累計值 / 累計樣本, NA_real_)
    ) %>%
    ungroup() %>%
    filter(!is.na(實際)) %>%
    select(據點名稱, 日期, 實際)
}

combine_forecast_result <- function(model_name, result_list, actual_df, metric_label,
                                    risk_direction, threshold) {
  bind_rows(lapply(names(result_list), function(branch) {
    result <- result_list[[branch]]
    if (is.null(result) || is.null(result$forecast)) return(NULL)

    forecast_df <- result$forecast %>%
      mutate(
        據點名稱 = branch,
        模型 = model_name,
        指標 = metric_label,
        risk_direction = risk_direction,
        threshold = threshold,
        日期 = as.Date(日期)
      ) %>%
      select(據點名稱, 模型, 指標, risk_direction, threshold, 日期, 預測, 下界, 上界)

    actual_branch <- actual_df %>%
      filter(據點名稱 == branch) %>%
      mutate(
        模型 = model_name,
        指標 = metric_label,
        risk_direction = risk_direction,
        threshold = threshold
      ) %>%
      select(據點名稱, 模型, 指標, risk_direction, threshold, 日期, 實際)

    full_join(forecast_df, actual_branch,
              by = c("據點名稱", "模型", "指標", "risk_direction", "threshold", "日期"))
  }))
}

make_trend_combo_plotly <- function(trend_df) {
  combo_levels <- trend_df %>%
    distinct(據點名稱, 指標, 模型) %>%
    arrange(據點名稱, 指標, 模型) %>%
    mutate(combo = paste(據點名稱, 指標, 模型, sep = "｜"))

  fig <- plot_ly()
  trace_meta <- tibble(combo = character(), trace_id = integer())
  trace_id <- 0

  for (combo in combo_levels$combo) {
    parts <- strsplit(combo, "｜", fixed = TRUE)[[1]]
    branch <- parts[[1]]
    metric <- parts[[2]]
    model <- parts[[3]]

    d <- trend_df %>%
      filter(據點名稱 == branch, 指標 == metric, 模型 == model) %>%
      arrange(日期)

    actual_df <- d %>% filter(!is.na(實際))
    forecast_df <- d %>% filter(!is.na(預測))
    threshold <- d$threshold[which(!is.na(d$threshold))[1]]
    visible <- combo == combo_levels$combo[[1]]

    if (nrow(actual_df) > 0) {
      trace_id <- trace_id + 1
      fig <- fig %>%
        add_lines(
          data = actual_df,
          x = ~日期,
          y = ~實際,
          name = "實際累計",
          line = list(color = "#002060", width = 3),
          visible = visible,
          hovertemplate = "<b>%{x}</b><br>實際累計: %{y:.1%}<extra></extra>"
        )
      trace_meta <- bind_rows(trace_meta, tibble(combo = combo, trace_id = trace_id))
    }

    if (nrow(forecast_df) > 0) {
      trace_id <- trace_id + 1
      fig <- fig %>%
        add_ribbons(
          data = forecast_df,
          x = ~日期,
          ymin = ~下界,
          ymax = ~上界,
          name = "預測區間",
          fillcolor = "rgba(158,197,254,0.35)",
          line = list(color = "rgba(158,197,254,0)"),
          visible = visible,
          hoverinfo = "skip",
          showlegend = FALSE
        )
      trace_meta <- bind_rows(trace_meta, tibble(combo = combo, trace_id = trace_id))

      trace_id <- trace_id + 1
      fig <- fig %>%
        add_lines(
          data = forecast_df,
          x = ~日期,
          y = ~預測,
          name = paste0(model, " 預測"),
          line = list(color = "#0056B3", width = 3, dash = "dash"),
          visible = visible,
          hovertemplate = "<b>%{x}</b><br>預測: %{y:.1%}<extra></extra>"
        )
      trace_meta <- bind_rows(trace_meta, tibble(combo = combo, trace_id = trace_id))

      risk_direction <- d$risk_direction[which(!is.na(d$risk_direction))[1]]
      risk_day <- if (risk_direction == "high") {
        forecast_df %>% slice_max(order_by = 預測, n = 1, with_ties = FALSE)
      } else {
        forecast_df %>% slice_min(order_by = 預測, n = 1, with_ties = FALSE)
      }
      risk_day <- risk_day %>%
        mutate(
          is_breach = case_when(
            is.na(threshold) ~ FALSE,
            risk_direction == "high" ~ 預測 > threshold,
            TRUE ~ 預測 < threshold
          ),
          risk_label = paste0(
            ifelse(is_breach, "🔴", "🟡"),
            " 最高風險日<br>",
            format(日期, "%m/%d"),
            "<br>",
            scales::percent(預測, accuracy = 0.1)
          )
        )

      trace_id <- trace_id + 1
      fig <- fig %>%
        add_trace(
          data = risk_day,
          x = ~日期,
          y = ~預測,
          type = "scatter",
          mode = "markers+text",
          name = "最高風險日",
          marker = list(
            color = ifelse(risk_day$is_breach, "#E74C3C", "#F39C12"),
            size = 11
          ),
          text = ~risk_label,
          textposition = "top center",
          textfont = list(size = 11, color = "#333333"),
          visible = visible,
          hovertemplate = "<b>最高風險日</b><br>%{x}<br>預測: %{y:.1%}<extra></extra>"
        )
      trace_meta <- bind_rows(trace_meta, tibble(combo = combo, trace_id = trace_id))
    }

    if (!is.na(threshold)) {
      x_range <- range(d$日期, na.rm = TRUE)
      threshold_df <- tibble(日期 = x_range, threshold = threshold)

      trace_id <- trace_id + 1
      fig <- fig %>%
        add_lines(
          data = threshold_df,
          x = ~日期,
          y = ~threshold,
          name = "門檻",
          line = list(color = "#E74C3C", width = 2, dash = "dot"),
          visible = visible,
          hovertemplate = "門檻: %{y:.1%}<extra></extra>"
        )
      trace_meta <- bind_rows(trace_meta, tibble(combo = combo, trace_id = trace_id))
    }
  }

  buttons <- lapply(combo_levels$combo, function(combo_name) {
    vis <- rep(FALSE, trace_id)
    ids <- trace_meta %>%
      filter(.data$combo == combo_name) %>%
      pull(trace_id)
    vis[ids] <- TRUE

    list(
      method = "update",
      args = list(
        list(visible = vis),
        list(
          title = combo_name,
          xaxis = list(title = "日期", autorange = TRUE),
          yaxis = list(title = "累計比例", tickformat = ".0%", autorange = TRUE)
        )
      ),
      label = combo_name
    )
  })

  fig %>%
    layout(
      title = combo_levels$combo[[1]],
      xaxis = list(title = "日期"),
      yaxis = list(title = "累計比例", tickformat = ".0%", autorange = TRUE),
      legend = list(orientation = "h", x = 0, y = -0.2),
      updatemenus = list(
        list(
          type = "dropdown",
          x = 1.02,
          y = 1,
          xanchor = "left",
          yanchor = "top",
          buttons = buttons
        )
      )
    )
}

trend_combo_df <- bind_rows(
  combine_forecast_result(
    "ARIMA", plots_lu_cum, make_cum_actual_df(df_daily, "落實比例"),
    "A⁺ 落實度", "low", 0.97
  ),
  combine_forecast_result(
    "ARIMA", plots_sat_cum, make_cum_actual_df(df_daily, "滿意比例"),
    "滿意度", "low", 0.97
  ),
  combine_forecast_result(
    "ARIMA", plots_low_cum, make_cum_actual_df(df_daily, "低滿意比例"),
    "低滿意比例", "high", 0.02
  ),
  combine_forecast_result(
    "Prophet", plots_lu_prophet, make_cum_actual_df(df_daily, "落實比例"),
    "A⁺ 落實度", "low", 0.97
  ),
  combine_forecast_result(
    "Prophet", plots_sat_prophet, make_cum_actual_df(df_daily, "滿意比例"),
    "滿意度", "low", 0.97
  ),
  combine_forecast_result(
    "Prophet", plots_low_prophet, make_cum_actual_df(df_daily, "低滿意比例"),
    "低滿意比例", "high", 0.02
  )
)
fig_trend_combo <- make_trend_combo_plotly(trend_combo_df)

# RStudio 互動開發時才顯示預覽；Rscript pipeline 跳過 graphics device 繪圖。
if (interactive()) {
  print(fig_trend_combo)
}

# Prophet 明細圖與 metrics 保留在 plots_*_prophet 中，需要時再手動 print()。

#三項指標
# === 1. 題目欄位 ===
question_cols <- grep("^[0-9]+-", names(df), value = TRUE)
score_questions <- question_cols[!grepl("原因|改進|建議|類別", question_cols)]

# === 2. 建立 point-biserial correlation 函數 ===
point_biserial <- function(x, y) {
  # 只處理 0/1 二元變數
  if (length(unique(na.omit(y))) != 2) return(NA_real_)
  y <- as.numeric(y)
  x1 <- x[y == 1]
  x0 <- x[y == 0]
  m1 <- mean(x1, na.rm = TRUE)
  m0 <- mean(x0, na.rm = TRUE)
  s <- sd(x, na.rm = TRUE)
  p <- mean(y, na.rm = TRUE)
  q <- 1 - p
  if (is.na(s) || s == 0) return(NA_real_)
  (m1 - m0) / s * sqrt(p * q)
}

# === 3. 資料整理 ===
df_scored <- df %>%
  dplyr::select(all_of(c("A+落實", "高滿意", "低滿意", score_questions))) %>%
  mutate(across(everything(), as.numeric))

# === 4. 通用分析函數 ===
calc_corr_auto <- function(df, target_var) {
  target <- df[[target_var]]
  res <- lapply(score_questions, function(q) {
    x <- df[[q]]
    r <- suppressWarnings(point_biserial(x, target))
    tibble(題項 = q, 相關係數 = r)
  }) %>% bind_rows()

  res %>% arrange(desc(相關係數))
}

# === 5. 三組分析 ===
corr_lu  <- calc_corr_auto(df_scored, "A+落實")
corr_sat <- calc_corr_auto(df_scored, "高滿意")
corr_low <- calc_corr_auto(df_scored, "低滿意")

# === 6. 視覺化 ===
plot_top_bottom <- function(corr_df, title_text, n = 5) {
  corr_df <- corr_df %>% filter(!is.na(相關係數))
  top_df <- corr_df %>% slice_head(n = n)
  bottom_df <- corr_df %>% slice_tail(n = n)

  p_top <- ggplot(top_df, aes(x = reorder(題項, 相關係數), y = 相關係數, fill = 相關係數)) +
    geom_col() +
    coord_flip() +
    scale_fill_gradient2(low = "blue", mid = "white", high = "red", midpoint = 0) +
    labs(title = paste0(title_text, " 強項 Top", n),
         x = "題項", y = "相關係數") +
    theme_minimal(base_family = "Microsoft JhengHei") +
    theme(plot.title = element_text(hjust = 0.5, face = "bold"))

  p_bottom <- ggplot(bottom_df, aes(x = reorder(題項, 相關係數), y = 相關係數, fill = 相關係數)) +
    geom_col() +
    coord_flip() +
    scale_fill_gradient2(low = "blue", mid = "white", high = "red", midpoint = 0) +
    labs(title = paste0(title_text, " 弱項 Bottom", n),
         x = "題項", y = "相關係數") +
    theme_minimal(base_family = "Microsoft JhengHei") +
    theme(plot.title = element_text(hjust = 0.5, face = "bold"))

  list(強項圖 = p_top, 弱項圖 = p_bottom)
}

# === 7. 繪圖 ===
plots_lu  <- plot_top_bottom(corr_lu,  "A⁺落實度")
plots_sat <- plot_top_bottom(corr_sat, "高滿意")
plots_low <- plot_top_bottom(corr_low, "低滿意")

# 圖形已保留在 plots_lu / plots_sat / plots_low，需要時再手動 print()。

# === 差異分析函數 ===
calc_gap <- function(df, target_var) {
  target <- df[[target_var]]

  gap_df <- lapply(score_questions, function(q) {
    x <- df[[q]]
    m1 <- mean(x[target == 1], na.rm = TRUE)
    m0 <- mean(x[target == 0], na.rm = TRUE)
    gap <- m1 - m0
    tibble(題項 = q, 差異值 = gap)
  }) %>% bind_rows()

  gap_df %>% arrange(差異值)
}

# === 計算三組 ===
gap_lu  <- calc_gap(df_scored, "A+落實")
gap_sat <- calc_gap(df_scored, "高滿意")
gap_low <- calc_gap(df_scored, "低滿意")  # 這組是反向（低滿意者=1）

# === 視覺化：分數落差最低（弱項） ===
plot_gap_bottom <- function(gap_df, title_text, n = 5) {
  gap_df <- gap_df %>% filter(!is.na(差異值))
  bottom_df <- gap_df %>% slice_min(order_by = 差異值, n = n)

  ggplot(bottom_df, aes(x = reorder(題項, 差異值), y = 差異值, fill = 差異值)) +
    geom_col() +
    coord_flip() +
    scale_fill_gradient2(low = "red", mid = "white", high = "blue", midpoint = 0) +
    labs(title = paste0(title_text, " 弱項 Bottom", n, "（低滿意群體分數較低）"),
         x = "題項", y = "分數差（高滿意−低滿意）") +
    theme_minimal(base_family = "Microsoft JhengHei") +
    theme(plot.title = element_text(hjust = 0.5, face = "bold", size = 14))
}

# 需要時可手動呼叫 plot_gap_bottom()。

# === 1️⃣ 整體資料三軸對照表 ===
corr_all <- corr_lu %>%
  rename(Aplus落實 = 相關係數) %>%
  full_join(corr_sat %>% rename(高滿意 = 相關係數), by = "題項") %>%
  full_join(corr_low %>% rename(低滿意 = 相關係數), by = "題項") %>%
  pivot_longer(cols = c(Aplus落實, 高滿意, 低滿意),
               names_to = "指標", values_to = "相關係數") %>%
  filter(!is.na(相關係數))

# === 2️⃣ 整體對照圖 ===
p_all <- ggplot(corr_all, aes(x = fct_reorder(題項, 相關係數), y = 相關係數, fill = 指標)) +
  geom_col(position = position_dodge(width = 0.8)) +
  coord_flip() +
  scale_fill_manual(values = c("Aplus落實" = "#0072B2", "高滿意" = "#009E73", "低滿意" = "#D55E00")) +
  labs(title = "CSI 問卷題項三軸對照圖（整體樣本）",
       subtitle = "顯示各題項對三個核心指標的點二列相關係數",
       x = "題項", y = "相關係數",
       fill = "指標") +
  theme_minimal(base_family = "Microsoft JhengHei") +
  theme(
    plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
    plot.subtitle = element_text(size = 12, hjust = 0.5),
    legend.position = "top"
  )

# 需要時可手動 print(p_all)。

# === 3️⃣ 各廠據點版本 ===
branches <- unique(df$據點名稱)
plots_branch <- list()

for (b in branches) {
  df_branch <- df %>% filter(據點名稱 == b)

  # 重算三項 point-biserial
  df_scored_b <- df_branch %>%
    dplyr::select(all_of(c("A+落實", "高滿意", "低滿意", score_questions))) %>%
    mutate(across(everything(), as.numeric))

  corr_lu_b  <- calc_corr_auto(df_scored_b, "A+落實") %>% rename(Aplus落實 = 相關係數)
  corr_sat_b <- calc_corr_auto(df_scored_b, "高滿意") %>% rename(高滿意 = 相關係數)
  corr_low_b <- calc_corr_auto(df_scored_b, "低滿意") %>% rename(低滿意 = 相關係數)

  corr_b <- corr_lu_b %>%
    full_join(corr_sat_b, by = "題項") %>%
    full_join(corr_low_b, by = "題項") %>%
    pivot_longer(cols = c(Aplus落實, 高滿意, 低滿意),
                 names_to = "指標", values_to = "相關係數") %>%
    filter(!is.na(相關係數))

  p_b <- ggplot(corr_b, aes(x = fct_reorder(題項, 相關係數), y = 相關係數, fill = 指標)) +
    geom_col(position = position_dodge(width = 0.8)) +
    coord_flip() +
    scale_fill_manual(values = c("Aplus落實" = "#0072B2", "高滿意" = "#009E73", "低滿意" = "#D55E00")) +
    labs(title = paste0(b, " — 三軸對照圖"),
         subtitle = "A⁺落實、高滿意、低滿意三指標點二列相關係數",
         x = "題項", y = "相關係數") +
    theme_minimal(base_family = "Microsoft JhengHei") +
    theme(
      plot.title = element_text(size = 15, face = "bold", hjust = 0.5),
      legend.position = "top"
    )

  plots_branch[[b]] <- p_b
}

# === 顯示例子 ===
#plots_branch[["Branch 01"]]
#plots_branch[["Branch 02"]]
#plots_branch[["Branch 03"]]
#plots_branch[["Branch 04"]]
#plots_branch[["Branch 05"]]

make_corr_switch_plotly <- function(corr_long, title_prefix = "三軸對照圖") {

  # 固定指標順序與顯示名
  metric_levels <- c("Aplus落實", "低滿意", "高滿意")

  df <- corr_long %>%
    filter(指標 %in% metric_levels) %>%
    mutate(
      指標 = factor(指標, levels = metric_levels),
      相關係數 = as.numeric(相關係數)
    )

  # 針對每個指標：先排序，並存下「題項順序」（用於 categoryarray）
  df_list <- list()
  y_orders <- list()

  for (m in metric_levels) {
    d <- df %>%
      filter(指標 == m) %>%
      arrange(desc(相關係數))

    df_list[[m]] <- d
    # ✅ 這個順序就是「由高到低」，我們要它顯示在上到下
    y_orders[[m]] <- rev(d$題項)  # plotly 類別軸需要反轉，才能讓最大在最上面
  }

  fig <- plot_ly()

  for (i in seq_along(metric_levels)) {
    m <- metric_levels[i]
    d <- df_list[[m]]

    fig <- fig %>%
      add_trace(
        data = d,
        type = "bar",
        orientation = "h",
        x = ~相關係數,
        y = ~題項,                 # ✅ 用原字串題項，不用 factor
        name = m,
        visible = (i == 1),
        hovertemplate = paste0(
          "<b>題項：</b>%{y}<br>",
          "<b>指標：</b>", m, "<br>",
          "<b>相關係數：</b>%{x:.3f}<extra></extra>"
        )
      )
  }

  # ✅ Dropdown：除了 visible，也同步更新 yaxis 的 categoryarray
  buttons <- lapply(seq_along(metric_levels), function(i) {
    m <- metric_levels[i]
    vis <- rep(FALSE, length(metric_levels)); vis[i] <- TRUE

    list(
      method = "update",
      args = list(
        list(visible = vis),
        list(
          title = paste0(title_prefix, "｜", m, ""),
          yaxis = list(
            title = "題項（依所選指標排序）",
            categoryorder = "array",
            categoryarray = y_orders[[m]]   # ✅ 關鍵：鎖定順序
          ),
          xaxis = list(title = "點二列相關係數")
        )
      ),
      label = m
    )
  })

  # 初始顯示 Aplus落實 的順序
  fig %>%
    layout(
      title = paste0(title_prefix, "｜Aplus落實"),
      margin = list(l = 220, r = 40, t = 80, b = 60),
      xaxis = list(title = "點二列相關係數", zeroline = TRUE),
      yaxis = list(
        title = "題項（依所選指標排序）",
        categoryorder = "array",
        categoryarray = y_orders[["Aplus落實"]]
      ),
      updatemenus = list(
        list(
          type = "dropdown",
          x = 1.02, y = 1,
          xanchor = "left", yanchor = "top",
          buttons = buttons
        )
      )
    )
}

# --- 2) 一次產生「每個廠」的 plotly ---
site_order <- c("Branch 01", "Branch 02", "Branch 03", "Branch 04", "Branch 05")

plots_branch_switch <- list()

for (b in site_order) {
  df_branch <- df %>% filter(據點名稱 == b)

  # 這段跟你原本一致：先轉成可算相關的 numeric 資料
  df_scored_b <- df_branch %>%
    dplyr::select(all_of(c("A+落實", "高滿意", "低滿意", score_questions))) %>%
    mutate(across(everything(), as.numeric))

  corr_lu_b  <- calc_corr_auto(df_scored_b, "A+落實") %>% rename(Aplus落實 = 相關係數)
  corr_sat_b <- calc_corr_auto(df_scored_b, "高滿意") %>% rename(高滿意 = 相關係數)
  corr_low_b <- calc_corr_auto(df_scored_b, "低滿意") %>% rename(低滿意 = 相關係數)

  corr_b <- corr_lu_b %>%
    full_join(corr_sat_b, by = "題項") %>%
    full_join(corr_low_b, by = "題項") %>%
    pivot_longer(cols = c(Aplus落實, 低滿意, 高滿意),
                 names_to = "指標", values_to = "相關係數") %>%
    filter(!is.na(相關係數))

  plots_branch_switch[[b]] <- make_corr_switch_plotly(
    corr_b,
    title_prefix = paste0(b, "｜題項相關度")
  )
}

# 各據點互動圖已保留在 plots_branch_switch，需要時再手動 print()。

# === 廠別代號與名稱 ===
branch_codes <- c("B01", "B02", "B03", "B04", "B05")
branch_names <- c("Branch 01", "Branch 02", "Branch 03", "Branch 04", "Branch 05")

# === 篩出可加強勸誘案件 ===
df_inducible <- df2 %>%
  filter(可加強勸誘 == 1)

# === 固定依 B01~B05 分 sheet ===
export_list <- lapply(seq_along(branch_codes), function(i) {
  code <- branch_names[i]

  df_inducible %>%
    filter(據點名稱 == code)
})

# === sheet 名稱 ===
names(export_list) <- paste0(branch_codes, "_", branch_names)

# === 輸出 Excel ===
write_xlsx(
  export_list,
  path = file.path(
    output_dir,
    sprintf("可加強勸誘名單_依廠別_%s.xlsx", format(Sys.Date(), "%Y%m%d"))
  )
)

df_agent_daily <- df %>%
  mutate(
    日期 = as_date(填寫問卷時間),
    據點名稱 = as.character(據點名稱),
    服務專員姓名 = trimws(as.character(服務專員姓名))
  ) %>%
  filter(!is.na(服務專員姓名), 服務專員姓名 != "") %>%
  group_by(據點名稱, 服務專員姓名, 日期) %>%
  summarise(
    樣本數 = n(),
    落實件數 = sum(`A+落實` == 1, na.rm = TRUE),
    高滿意件數 = sum(高滿意 == 1, na.rm = TRUE),
    低滿意件數 = sum(低滿意 == 1, na.rm = TRUE),
    落實比例 = 落實件數 / 樣本數,
    滿意比例 = 高滿意件數 / 樣本數,
    低滿意比例 = 低滿意件數 / 樣本數,
    .groups = "drop"
  ) %>%
  arrange(據點名稱, 服務專員姓名, 日期)

df_agent_cum <- df_agent_daily %>%
  group_by(據點名稱, 服務專員姓名) %>%
  arrange(日期, .by_group = TRUE) %>%
  mutate(
    累計樣本 = cumsum(樣本數),
    累計落實件數 = cumsum(落實件數),
    累計高滿意件數 = cumsum(高滿意件數),
    累計低滿意件數 = cumsum(低滿意件數),
    累計落實比例 = 累計落實件數 / 累計樣本,
    累計滿意比例 = 累計高滿意件數 / 累計樣本,
    累計低滿意比例 = 累計低滿意件數 / 累計樣本
  ) %>%
  ungroup()

# === 統計每位專員累計樣本數，供篩選 ===
agent_list <- df_agent_cum %>%
  group_by(據點名稱, 服務專員姓名) %>%
  summarise(
    累計樣本數 = max(累計樣本, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(
    key = paste(據點名稱, 服務專員姓名, sep = "｜")
  ) %>%
  arrange(據點名稱, desc(累計樣本數), 服務專員姓名)

make_agent_csi_dropdown_plot <- function(min_n = 10) {

  # 只保留樣本數達門檻的專員
  valid_agents <- agent_list %>%
    filter(累計樣本數 >= min_n)

  if (nrow(valid_agents) == 0) {
    message("沒有符合樣本數門檻的專員。")
    return(NULL)
  }

  fig <- plot_ly()

  trace_meta <- data.frame(
    agent_key = character(),
    metric = character(),
    trace_id = integer(),
    stringsAsFactors = FALSE
  )

  trace_id <- 0

  metric_order <- c("累計落實比例", "累計滿意比例", "累計低滿意比例")
  metric_label_map <- c(
    "累計落實比例" = "落實度",
    "累計滿意比例" = "滿意度",
    "累計低滿意比例" = "低滿意比例"
  )

  # === 逐專員加三條線 ===
  for (i in seq_len(nrow(valid_agents))) {
    site  <- valid_agents$據點名稱[i]
    agent <- valid_agents$服務專員姓名[i]
    key   <- valid_agents$key[i]
    n_obs <- valid_agents$累計樣本數[i]

    df_one <- df_agent_cum %>%
      filter(據點名稱 == site, 服務專員姓名 == agent) %>%
      select(
        日期,
        累計落實比例,
        累計滿意比例,
        累計低滿意比例
      )

    for (m in metric_order) {
      trace_id <- trace_id + 1

      fig <- fig %>%
        add_lines(
          data = df_one,
          x = ~日期,
          y = as.formula(paste0("~", m)),
          name = metric_label_map[[m]],
          visible = ifelse(i == 1, TRUE, FALSE),   # 初始只顯示第一位專員
          hovertemplate = paste0(
            "<b>據點：</b>", site, "<br>",
            "<b>專員：</b>", agent, "<br>",
            "<b>樣本數：</b>", n_obs, "<br>",
            "<b>日期：</b>%{x}<br>",
            "<b>指標：</b>", metric_label_map[[m]], "<br>",
            "<b>累計比例：</b>%{y:.1%}<extra></extra>"
          )
        )

      trace_meta <- bind_rows(
        trace_meta,
        data.frame(
          agent_key = key,
          metric = m,
          trace_id = trace_id,
          stringsAsFactors = FALSE
        )
      )
    }
  }

  # === dropdown 按鈕：每次顯示一位專員的三條線 ===
  buttons <- lapply(seq_len(nrow(valid_agents)), function(i) {
    key  <- valid_agents$key[i]
    site <- valid_agents$據點名稱[i]
    agent <- valid_agents$服務專員姓名[i]
    n_obs <- valid_agents$累計樣本數[i]

    vis <- rep(FALSE, trace_id)
    ids <- trace_meta %>%
      filter(agent_key == key) %>%
      pull(trace_id)
    vis[ids] <- TRUE

    list(
      method = "update",
      args = list(
        list(visible = vis),
        list(
          title = paste0(site, "｜", agent, "｜CSI 指標歷史累計推移（n=", n_obs, "）")
        )
      ),
      label = paste0(agent, " (", site, ")")
    )
  })

  # === 初始標題 ===
  init_site  <- valid_agents$據點名稱[1]
  init_agent <- valid_agents$服務專員姓名[1]
  init_n     <- valid_agents$累計樣本數[1]

  fig %>%
    layout(
      title = paste0(init_site, "｜", init_agent, "｜CSI 指標歷史累計推移（n=", init_n, "）"),
      xaxis = list(title = "日期"),
      yaxis = list(
        title = "累計比例",
        tickformat = ".0%",
        range = c(0, 1.0) ,
        x = 2.0,
        y = 1.5
      ),
      legend = list(
        orientation = "V",
        x = 1.0,
        y = 0.5
      ),
      updatemenus = list(
        list(
          type = "dropdown",
          x = 1.02,
          y = 0.95,
          xanchor = "left",
          yanchor = "top",
          buttons = buttons
        )
      )
    )
}

fig_agent_csi <- make_agent_csi_dropdown_plot(min_n = 10)
# 需要時可手動 print(fig_agent_csi)。

# === 輸出清單 ===
export_list <- list(
  問卷明細 = df,
  回函率明細 = df2
)

# === 輸出 Excel ===
write_xlsx(
  export_list,
  path = file.path(
    output_dir,
    sprintf("CSI明細_%s.xlsx", format(Sys.Date(), "%Y%m%d"))
  )
)
