r"""
=====================================================================================
MARQUEE INVESTOR PORTFOLIO  +  GICS SECTOR SCREENER
=====================================================================================
One script. Put the two CSV files in a folder, run it, and you get the formatted
workbook plus an interactive HTML screener. When new data arrives, replace the CSV
files and run again. No code changes needed.

-------------------------------------------------------------------------------------
STEP 1 - WHERE TO PUT THE CSV FILES
-------------------------------------------------------------------------------------
  Create a folder, put BOTH CSV files in it, and write that folder path into DATA_DIR
  below. The folder can be named anything.

      screener/
        marquee_screener.py                  <- this script
        data/                                <- DATA_DIR   (put the CSV files here)
          <any name>.csv                         1) All Stocks screener export
          <any name>.csv                         2) Dataroma marquee export
        output/                              <- OUT_DIR    (created automatically)
          Marquee_Investor_Portfolio_Consolidated.xlsx
          Sector_Screener.html

  COLUMN NAMES CAN DRIFT TOO. If TradingView renames a column, for example switching a
  margin from "Annual" to "Trailing 12 months", the script finds the new name on its
  own and says so in the run log. A metric dropped from the export entirely just has its
  filter switched off for that run. Only the identity columns (Symbol, Description,
  Country, Sector, Industry, Price, Market capitalization) are truly required.

  FILENAMES DO NOT MATTER. The script opens every CSV in the folder and works out
  which is which from the column headers, so you can rename or re-export the files
  as often as you like. The All Stocks export is recognised by columns such as
  "Market capitalization" and "Price to earnings ratio"; the Dataroma export by its
  ownership columns such as "Ownership count" and "Hold Price".
  If an older copy of either export is still sitting in the folder, the one with the
  latest date in its filename is used (modified time breaks ties, or decides when the
  name carries no date) and the older one is reported and skipped. Any CSV that matches
  neither export is ignored.

  The Dataroma half can also be scraped for you instead of exported by hand - see
  SCRAPE_DATAROMA below. In that case only the All Stocks CSV needs to be in the folder.

-------------------------------------------------------------------------------------
STEP 2 - RUN IT
-------------------------------------------------------------------------------------
  Once DATA_DIR is set:

      python3 marquee_screener.py

  Or pass the folders on the command line instead:

      python3 marquee_screener.py --data-dir "D:/work/screener/data" --out-dir "D:/work/screener/output"

  In Jupyter, Spyder or a VS Code notebook, run it as a cell like this:

      %run marquee_screener.py

  The notebook is detected automatically and its own arguments are ignored, so the
  DATA_DIR and OUT_DIR values below are what get used. Windows paths need the r prefix:

      DATA_DIR = r'C:\Users\LENOVO\Desktop\Stock Research\CSV'
      OUT_DIR  = r'C:\Users\LENOVO\Desktop\Stock Research\CSV\OUTPUT'

-------------------------------------------------------------------------------------
WHAT TO CHANGE AND WHERE  (seven places, no need to read the rest of the code)
-------------------------------------------------------------------------------------
  [0] DATA_DIR / OUT_DIR   Where the CSV files live and where output goes. Top of file.
  [1] SCREEN_UNIVERSE      'all' screens the full 30,000 stock universe.
                           'marquee' screens only the superinvestor holdings.
  [2] DEFAULT_FILTERS      The base threshold set. Since every sector now has its own
                           complete block, this only applies to a sector you add to
                           SECTOR_MAP without giving it a block. All four margins are
                           here - gross, operating, net and free cash flow - and the
                           FCF margin filter is switched on.
  [3] SECTOR_OVERRIDES     THE MAIN ONE. Every GICS sector has its own block, and each
                           block lists all fifteen metrics explicitly. Nothing is
                           hidden and nothing arrives silently from the defaults.
                           Change a value directly, or write None to switch that
                           filter off. Banks never clear ROA above 10% and utilities
                           never clear debt to equity below 2, so a few sectors carry
                           values that differ from the defaults - those lines are
                           marked "# <- changed from default".
                           Nothing passing? Open that sector's block, loosen a value
                           and run again.
                           'All Sectors' is the first block. It screens the whole
                           universe in one go and feeds the "Screen - All Sectors"
                           sheet, and it is editable exactly like the rest.
  [4] BLANK_FAILS /        True rejects any stock missing a screened value.
      MIN_DATA_COVERAGE    False ignores the blank and judges the stock on what it
                           does report. Altman Z covers about 76% of the universe and
                           Piotroski about 63%, so True returns far fewer names. When
                           BLANK_FAILS is False, MIN_DATA_COVERAGE is what stops a
                           thinly covered stock from passing by default.
  [5] SECTOR_MAP           The GICS name against its TradingView sector list. Your
                           mapping is already in place. Add a new sector here.

  You do not need to open Excel to check the thresholds. The "Screen Settings" sheet
  lists the final threshold set for every sector, with overrides flagged in a
  different colour.

-------------------------------------------------------------------------------------
THE INTERACTIVE HTML
-------------------------------------------------------------------------------------
  Sector_Screener.html lands in the output folder. Double-click it in a browser. No
  server and no internet connection needed, because the data is embedded in the file.
  There you can change every threshold live, switch sector (including All Sectors),
  toggle US against non-US, and download the result as CSV.
  To push the screener to a live website on every run, switch GITHUB_UPLOAD on and set
  a token. Setup steps are written above that setting.

  Changes made in the HTML are temporary. Once you settle on a threshold set, use the
  "Copy as Python" button and paste the block into SECTOR_OVERRIDES below. Only then
  does it become permanent in the workbook.
=====================================================================================
"""

import argparse, datetime, glob, json, math, os, re, sys, time
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# =====================================================================================
# [0] FOLDERS  and  [1] UNIVERSE
# =====================================================================================
# Write your own folder here. The script uses these, so nothing needs to be typed on
# the command line.
#   Windows:      DATA_DIR = r'C:\Users\Shobhit\Desktop\screener\data'
#   Mac / Linux:  DATA_DIR = '/Users/shobhit/Desktop/screener/data'
DATA_DIR = 'data'
OUT_DIR  = 'output'

SCREEN_UNIVERSE = 'all'    # 'all' = full universe, 'marquee' = superinvestor holdings only

# --------------------------------------------------------------------------------------
# Dataroma scraping. The marquee holdings can be pulled straight off dataroma.com with
# Selenium instead of exporting a CSV by hand.
#   'never'   use the dataroma CSV already sitting in DATA_DIR
#   'auto'    scrape unless the newest dataroma CSV is already from today, so the marquee
#             date always moves to the date of the run. If the scrape fails, the newest
#             existing file is used and its own date is shown, never a date it does not have.
#   'always'  scrape on every run
# Needs Chrome plus:  pip install selenium
# The scraped file is written into DATA_DIR as dataroma_portfolio_YYYY-MM-DD.csv, so the
# normal auto-detection picks it up like any other export.
SCRAPE_DATAROMA = 'auto'
DATAROMA_MAX_AGE_DAYS = 0     # 0 = refresh on every run that does not already have today's file
DATAROMA_PAGES = 15
# Keep only the newest Dataroma file in DATA_DIR. Every older dataroma_portfolio_*.csv, and any
# other CSV recognised as a Dataroma export, is deleted once the newest one is chosen, so the
# folder never fills up. The workflow commits the deletion along with everything else.
KEEP_ONLY_LATEST_DATAROMA = True        # the scraper stops early on its own if a page comes back empty

# Which threshold set to run.
#   'base'   one set of bars for every sector: the median company of the whole market,
#            written into SECTOR_OVERRIDES. The like-for-like view.
#   'tuned'  each sector gets its own bars: the median company of that sector, computed
#            from the data on each run, see [3b] SECTOR-TUNED RULES. (SECTOR_PRESETS below
#            is the older hand-set version, kept for reference only.)
# Switch this one word and run again. Both sets are fully editable further down.
THRESHOLD_MODE = 'tuned'

# How the "Screen - All Sectors" sheet is built.
#   True   every row is judged against ITS OWN sector's block, and the results are then
#          combined into one list. A bank is judged on the Financials rules, a REIT on the
#          Real Estate rules, and so on. The combined list is exactly the union of the
#          eleven sector sheets. Rows whose TradingView sector maps to no GICS sector
#          (Miscellaneous, Government) are judged on the 'All Sectors' block instead.
#   False  one single rule set - the 'All Sectors' block - is applied to all 30,000 rows.
#          Stricter and like-for-like, but it wipes out sectors that structurally cannot
#          clear a market-wide rule.
ALL_SECTORS_USES_SECTOR_RULES = True

# Put the data date into the output filenames, e.g.
# Marquee_Investor_Portfolio_Consolidated_2026-08-10.xlsx. The date is read from the CSV
# filename, and from the file's modified date if the name carries no date.
# Set False for stable names, which is easier if you re-upload the HTML to a website.
DATE_IN_FILENAME = True

# Optional: a local folder that is also a git repository for a website, typically a
# GitHub Pages repo. When this is set, the screener is copied there as index.html on
# every run, so publishing is just: run this script, then git add / commit / push.
# Leave it as '' to switch the feature off.
#   PUBLISH_DIR = r'C:\Users\LENOVO\Desktop\mics-screener'
PUBLISH_DIR = ''

# --------------------------------------------------------------------------------------
# Deep analysis app. When this is set, every ticker in the HTML screener becomes a link
# to a Streamlit app that pulls live data for that one company and runs the forensic
# scores. Leave it as '' and the tickers stay plain text.
# Streamlit is the right home for that page because the fetch happens server side -
# a static page cannot call Yahoo directly, since Yahoo sends no CORS header.
#   DEEP_APP_URL = 'https://your-app.streamlit.app'
DEEP_APP_URL = 'https://mics-deep-analysis-dcxigpgnhvhywvq2599eh3.streamlit.app'
# The deep analysis scorecard, loaded from the deep app's own repo on each run so the export
# and the deep page always use the same scoring. A local deep_scoring.py next to this script
# is used first if present; if neither is reachable the export simply leaves those columns out.
DEEP_SCORING_URL = 'https://raw.githubusercontent.com/wealthmics/mics-deep-analysis/main/scoring.py'

# --------------------------------------------------------------------------------------
# GitHub upload. With this on, the screener is pushed straight to the repo at the end of
# every run, so the live site updates without opening a browser.
#
# One-time setup:
#   1. github.com -> click your avatar -> Settings (your account, not the repo)
#   2. Developer settings -> Personal access tokens -> Fine-grained tokens
#   3. Generate new token. Repository access: only your screener repo.
#      Permissions -> Repository permissions -> Contents: Read and write. Nothing else.
#   4. Copy the token (starts with github_pat_) and set it as an environment variable
#      called GITHUB_TOKEN, or paste it below.
#
# Windows, to set it permanently, run once in Command Prompt then reopen your editor:
#      setx GITHUB_TOKEN github_pat_xxxxxxxxxxxxxxxx
#
# A token is a password. Never share it, never screenshot it, and never let this script
# sit in a public repo with a real token pasted in. If one ever leaks, revoke it on the
# same GitHub page and generate a new one.
GITHUB_UPLOAD = False  # the workflow commits instead
GITHUB_REPO   = 'wealthmics/Stock-Screener'   # owner/repo
GITHUB_BRANCH = 'main'
GITHUB_PATH   = 'index.html'                  # the name GitHub Pages serves
GITHUB_XLSX   = 'Marquee_Investor_Portfolio_Consolidated.xlsx'   # '' to skip the workbook
# the deep analysis app fetches these two from the Pages site, so they have to be committed
GITHUB_FEED   = ['stocks.json', 'sector_stats.json']
# The token is written in below as you asked. Because it sits inside this file, do not
# upload this .py to the public Stock-Screener repo - GitHub's secret scanner will find it
# and disable the token. To move it out later, set the line back to '' and run once in
# Command Prompt:   setx GITHUB_TOKEN github_pat_your_token_here   then restart Jupyter.
# When the line is empty the script reads the GITHUB_TOKEN environment variable instead.
GITHUB_TOKEN  = ''

# How the HTML is written.
#   False  one self-contained file, about 6 MB. Double-click and it works, no internet,
#          no server. Best for emailing the file or keeping it on your own machine.
#   True   a light index.html plus a separate data.json. The page appears instantly and
#          the data loads behind a short "loading" message. This is what a website wants,
#          including GitHub Pages. It will NOT work by double-clicking from your desktop,
#          because browsers block local file reads.
# Both files are written to OUT_DIR, and to PUBLISH_DIR as index.html + data.json when set.
WEB_SPLIT = False

# =====================================================================================
# [4] MISSING DATA POLICY
# =====================================================================================
BLANK_FAILS = False        # True = strict, a blank value rejects the stock

# With BLANK_FAILS = False, a row that reports nothing at all would technically clear every
# filter. This guard stops that: a stock must report at least this share of the active
# filters to qualify. 0.6 means 60%. Set it to 0 to switch the guard off.
MIN_DATA_COVERAGE = 0.6

# =====================================================================================
# [2] DEFAULT FILTERS  ---  the base rule set
#     None = filter off.  '_min' means at least this value, '_max' means at most this value.
# =====================================================================================
DEFAULT_FILTERS = {
    'pe_max'            : 20,      # Price to earnings (TTM), at most
    'roe_min'           : 10,      # Return on equity %
    'roic_min'          : 10,      # Return on invested capital %
    'roce_min'          : 10,      # Return on capital employed %
    'roa_min'           : 10,      # Return on assets %
    'gross_margin_min'  : 10,      # Gross margin %
    'op_margin_min'     : 10,      # Operating margin %
    'net_margin_min'    : 10,      # Net margin %
    'fcf_margin_min'    : 10,      # Free cash flow margin %
    'altman_min'        : 2.99,    # Altman Z, above 2.99 is the safe zone
    'piotroski_min'     : 7,       # Piotroski F score out of 9
    'current_ratio_min' : 1,       # Current ratio
    'de_max'            : 2,       # Debt to equity, at most
    'int_cov_min'       : 3,       # Interest coverage
    'mcap_min'          : None,    # put a number here to drop micro caps
}

# =====================================================================================
# [3] SECTOR THRESHOLDS  ---  two complete sets, switched by THRESHOLD_MODE above
# =====================================================================================
#     Every block below lists all fifteen metrics explicitly. Nothing is hidden and
#     nothing arrives silently from somewhere else. Change a value directly, or write
#     None to switch that filter off.
#
#     SECTOR_OVERRIDES  is the 'base' set. Every sector gets the same bars: the median
#                       company of the whole market (companies above $300M). A pass means
#                       better than the typical listed company on every measure, whatever
#                       the sector. Banks, property and utilities rarely clear it. These are
#                       written numbers, so they stay put until you edit them.
#     SECTOR_PRESETS    the older hand-set version, kept for reference only. The 'tuned'
#                       set is computed on every run, see [3b].
#
#     Switch between them with THRESHOLD_MODE at the top. Edit either one freely.
# =====================================================================================
SECTOR_OVERRIDES = {

    # ---- All Sectors ---------------------------------------------------------------
    # The median company of the whole market, 18,269 companies above $300M, 22 October
    # 2026 data. Every sector below uses exactly these numbers.
    'All Sectors': {
        'pe_max'            : 19.4  ,
        'roe_min'           : 8.4   ,
        'roic_min'          : 5.9   ,
        'roce_min'          : None  ,
        'roa_min'           : 3.4   ,
        'gross_margin_min'  : 28.7  ,
        'op_margin_min'     : 10.2  ,
        'net_margin_min'    : 7.5   ,
        'fcf_margin_min'    : 6.7   ,
        'altman_min'        : 3.03  ,
        'piotroski_min'     : 6     ,
        'current_ratio_min' : 1.65  ,
        'de_max'            : 0.46  ,
        'int_cov_min'       : 5.3   ,
        'mcap_min'          : 300e6 ,
    },

    # ---- Information Technology ----------------------------------------------------
    'Information Technology': {
        'pe_max'            : 19.4  ,
        'roe_min'           : 8.4   ,
        'roic_min'          : 5.9   ,
        'roce_min'          : None  ,
        'roa_min'           : 3.4   ,
        'gross_margin_min'  : 28.7  ,
        'op_margin_min'     : 10.2  ,
        'net_margin_min'    : 7.5   ,
        'fcf_margin_min'    : 6.7   ,
        'altman_min'        : 3.03  ,
        'piotroski_min'     : 6     ,
        'current_ratio_min' : 1.65  ,
        'de_max'            : 0.46  ,
        'int_cov_min'       : 5.3   ,
        'mcap_min'          : 300e6 ,
    },

    # ---- Communication Services ----------------------------------------------------
    'Communication Services': {
        'pe_max'            : 19.4  ,
        'roe_min'           : 8.4   ,
        'roic_min'          : 5.9   ,
        'roce_min'          : None  ,
        'roa_min'           : 3.4   ,
        'gross_margin_min'  : 28.7  ,
        'op_margin_min'     : 10.2  ,
        'net_margin_min'    : 7.5   ,
        'fcf_margin_min'    : 6.7   ,
        'altman_min'        : 3.03  ,
        'piotroski_min'     : 6     ,
        'current_ratio_min' : 1.65  ,
        'de_max'            : 0.46  ,
        'int_cov_min'       : 5.3   ,
        'mcap_min'          : 300e6 ,
    },

    # ---- Consumer Discretionary ----------------------------------------------------
    'Consumer Discretionary': {
        'pe_max'            : 19.4  ,
        'roe_min'           : 8.4   ,
        'roic_min'          : 5.9   ,
        'roce_min'          : None  ,
        'roa_min'           : 3.4   ,
        'gross_margin_min'  : 28.7  ,
        'op_margin_min'     : 10.2  ,
        'net_margin_min'    : 7.5   ,
        'fcf_margin_min'    : 6.7   ,
        'altman_min'        : 3.03  ,
        'piotroski_min'     : 6     ,
        'current_ratio_min' : 1.65  ,
        'de_max'            : 0.46  ,
        'int_cov_min'       : 5.3   ,
        'mcap_min'          : 300e6 ,
    },

    # ---- Consumer Staples ----------------------------------------------------------
    'Consumer Staples': {
        'pe_max'            : 19.4  ,
        'roe_min'           : 8.4   ,
        'roic_min'          : 5.9   ,
        'roce_min'          : None  ,
        'roa_min'           : 3.4   ,
        'gross_margin_min'  : 28.7  ,
        'op_margin_min'     : 10.2  ,
        'net_margin_min'    : 7.5   ,
        'fcf_margin_min'    : 6.7   ,
        'altman_min'        : 3.03  ,
        'piotroski_min'     : 6     ,
        'current_ratio_min' : 1.65  ,
        'de_max'            : 0.46  ,
        'int_cov_min'       : 5.3   ,
        'mcap_min'          : 300e6 ,
    },

    # ---- Health Care ---------------------------------------------------------------
    'Health Care': {
        'pe_max'            : 19.4  ,
        'roe_min'           : 8.4   ,
        'roic_min'          : 5.9   ,
        'roce_min'          : None  ,
        'roa_min'           : 3.4   ,
        'gross_margin_min'  : 28.7  ,
        'op_margin_min'     : 10.2  ,
        'net_margin_min'    : 7.5   ,
        'fcf_margin_min'    : 6.7   ,
        'altman_min'        : 3.03  ,
        'piotroski_min'     : 6     ,
        'current_ratio_min' : 1.65  ,
        'de_max'            : 0.46  ,
        'int_cov_min'       : 5.3   ,
        'mcap_min'          : 300e6 ,
    },

    # ---- Financials ----------------------------------------------------------------
    'Financials': {
        'pe_max'            : 19.4  ,
        'roe_min'           : 8.4   ,
        'roic_min'          : 5.9   ,
        'roce_min'          : None  ,
        'roa_min'           : 3.4   ,
        'gross_margin_min'  : 28.7  ,
        'op_margin_min'     : 10.2  ,
        'net_margin_min'    : 7.5   ,
        'fcf_margin_min'    : 6.7   ,
        'altman_min'        : 3.03  ,
        'piotroski_min'     : 6     ,
        'current_ratio_min' : 1.65  ,
        'de_max'            : 0.46  ,
        'int_cov_min'       : 5.3   ,
        'mcap_min'          : 300e6 ,
    },

    # ---- Energy --------------------------------------------------------------------
    'Energy': {
        'pe_max'            : 19.4  ,
        'roe_min'           : 8.4   ,
        'roic_min'          : 5.9   ,
        'roce_min'          : None  ,
        'roa_min'           : 3.4   ,
        'gross_margin_min'  : 28.7  ,
        'op_margin_min'     : 10.2  ,
        'net_margin_min'    : 7.5   ,
        'fcf_margin_min'    : 6.7   ,
        'altman_min'        : 3.03  ,
        'piotroski_min'     : 6     ,
        'current_ratio_min' : 1.65  ,
        'de_max'            : 0.46  ,
        'int_cov_min'       : 5.3   ,
        'mcap_min'          : 300e6 ,
    },

    # ---- Materials -----------------------------------------------------------------
    'Materials': {
        'pe_max'            : 19.4  ,
        'roe_min'           : 8.4   ,
        'roic_min'          : 5.9   ,
        'roce_min'          : None  ,
        'roa_min'           : 3.4   ,
        'gross_margin_min'  : 28.7  ,
        'op_margin_min'     : 10.2  ,
        'net_margin_min'    : 7.5   ,
        'fcf_margin_min'    : 6.7   ,
        'altman_min'        : 3.03  ,
        'piotroski_min'     : 6     ,
        'current_ratio_min' : 1.65  ,
        'de_max'            : 0.46  ,
        'int_cov_min'       : 5.3   ,
        'mcap_min'          : 300e6 ,
    },

    # ---- Industrials ---------------------------------------------------------------
    'Industrials': {
        'pe_max'            : 19.4  ,
        'roe_min'           : 8.4   ,
        'roic_min'          : 5.9   ,
        'roce_min'          : None  ,
        'roa_min'           : 3.4   ,
        'gross_margin_min'  : 28.7  ,
        'op_margin_min'     : 10.2  ,
        'net_margin_min'    : 7.5   ,
        'fcf_margin_min'    : 6.7   ,
        'altman_min'        : 3.03  ,
        'piotroski_min'     : 6     ,
        'current_ratio_min' : 1.65  ,
        'de_max'            : 0.46  ,
        'int_cov_min'       : 5.3   ,
        'mcap_min'          : 300e6 ,
    },

    # ---- Utilities -----------------------------------------------------------------
    'Utilities': {
        'pe_max'            : 19.4  ,
        'roe_min'           : 8.4   ,
        'roic_min'          : 5.9   ,
        'roce_min'          : None  ,
        'roa_min'           : 3.4   ,
        'gross_margin_min'  : 28.7  ,
        'op_margin_min'     : 10.2  ,
        'net_margin_min'    : 7.5   ,
        'fcf_margin_min'    : 6.7   ,
        'altman_min'        : 3.03  ,
        'piotroski_min'     : 6     ,
        'current_ratio_min' : 1.65  ,
        'de_max'            : 0.46  ,
        'int_cov_min'       : 5.3   ,
        'mcap_min'          : 300e6 ,
    },

    # ---- Real Estate ---------------------------------------------------------------
    'Real Estate': {
        'pe_max'            : 19.4  ,
        'roe_min'           : 8.4   ,
        'roic_min'          : 5.9   ,
        'roce_min'          : None  ,
        'roa_min'           : 3.4   ,
        'gross_margin_min'  : 28.7  ,
        'op_margin_min'     : 10.2  ,
        'net_margin_min'    : 7.5   ,
        'fcf_margin_min'    : 6.7   ,
        'altman_min'        : 3.03  ,
        'piotroski_min'     : 6     ,
        'current_ratio_min' : 1.65  ,
        'de_max'            : 0.46  ,
        'int_cov_min'       : 5.3   ,
        'mcap_min'          : 300e6 ,
    },
}


# ---- the optional sector-tuned set, used when THRESHOLD_MODE = 'tuned' --------------
SECTOR_PRESETS = {

    # ---- All Sectors -----------------------------------------------------------
    # One relaxed set for the whole market. Sits roughly at the cross-sector median of
    # the marquee holdings, so the strict base rules do not wipe out entire sectors.
    'All Sectors': {
        'pe_max'            : 25    ,   # <- differs from the base set
        'roe_min'           : 10    ,
        'roic_min'          : None  ,   # <- differs from the base set
        'roce_min'          : None  ,   # <- differs from the base set
        'roa_min'           : 3     ,   # <- differs from the base set
        'gross_margin_min'  : 10    ,
        'op_margin_min'     : 10    ,
        'net_margin_min'    : 5     ,   # <- differs from the base set
        'fcf_margin_min'    : 5     ,   # <- differs from the base set
        'altman_min'        : 2     ,   # <- differs from the base set
        'piotroski_min'     : 5     ,   # <- differs from the base set
        'current_ratio_min' : 1     ,
        'de_max'            : 1.5   ,   # <- differs from the base set
        'int_cov_min'       : 3     ,
        'mcap_min'          : None  ,
    },

    # ---- Information Technology ------------------------------------------------
    # Marquee medians here: P/E 39.6, ROE 12.8, ROA 6.1, Altman 4.5. A P/E cap of 20
    # only admits 40% of the sector and ROA of 10% only 35%, so both are relaxed.
    'Information Technology': {
        'pe_max'            : 45    ,   # <- differs from the base set
        'roe_min'           : 12    ,   # <- differs from the base set
        'roic_min'          : None  ,   # <- differs from the base set
        'roce_min'          : None  ,   # <- differs from the base set
        'roa_min'           : 5     ,   # <- differs from the base set
        'gross_margin_min'  : 10    ,
        'op_margin_min'     : 10    ,
        'net_margin_min'    : 7     ,   # <- differs from the base set
        'fcf_margin_min'    : 10    ,
        'altman_min'        : 3     ,   # <- differs from the base set
        'piotroski_min'     : 6     ,   # <- differs from the base set
        'current_ratio_min' : 1.2   ,   # <- differs from the base set
        'de_max'            : 1     ,   # <- differs from the base set
        'int_cov_min'       : 4     ,   # <- differs from the base set
        'mcap_min'          : None  ,
    },

    # ---- Communication Services ------------------------------------------------
    # The thinnest sector - only 18 marquee names and weak fundamentals
    # (median ROE 7.1, Altman 1.5, interest coverage 0.9). Everything is set low deliberately;
    # tightening anything here empties the sheet.
    'Communication Services': {
        'pe_max'            : 25    ,   # <- differs from the base set
        'roe_min'           : 5     ,   # <- differs from the base set
        'roic_min'          : None  ,   # <- differs from the base set
        'roce_min'          : None  ,   # <- differs from the base set
        'roa_min'           : 1.5   ,   # <- differs from the base set
        'gross_margin_min'  : 10    ,
        'op_margin_min'     : 6     ,   # <- differs from the base set
        'net_margin_min'    : 3     ,   # <- differs from the base set
        'fcf_margin_min'    : 3     ,   # <- differs from the base set
        'altman_min'        : 1     ,   # <- differs from the base set
        'piotroski_min'     : 4     ,   # <- differs from the base set
        'current_ratio_min' : 0.8   ,   # <- differs from the base set
        'de_max'            : 2     ,
        'int_cov_min'       : 0.8   ,   # <- differs from the base set
        'mcap_min'          : None  ,
    },

    # ---- Consumer Discretionary ------------------------------------------------
    # Marquee medians: P/E 21.6, ROE 12.1, operating margin 8.0, net margin 4.4.
    # Volume business, so the margin bars sit below the base 10%.
    'Consumer Discretionary': {
        'pe_max'            : 28    ,   # <- differs from the base set
        'roe_min'           : 10    ,
        'roic_min'          : None  ,   # <- differs from the base set
        'roce_min'          : None  ,   # <- differs from the base set
        'roa_min'           : 3     ,   # <- differs from the base set
        'gross_margin_min'  : 10    ,
        'op_margin_min'     : 7     ,   # <- differs from the base set
        'net_margin_min'    : 3.5   ,   # <- differs from the base set
        'fcf_margin_min'    : 4     ,   # <- differs from the base set
        'altman_min'        : 2.2   ,   # <- differs from the base set
        'piotroski_min'     : 6     ,   # <- differs from the base set
        'current_ratio_min' : 0.9   ,   # <- differs from the base set
        'de_max'            : 1.5   ,   # <- differs from the base set
        'int_cov_min'       : 3     ,
        'mcap_min'          : None  ,
    },

    # ---- Consumer Staples ------------------------------------------------------
    # Marquee medians: P/E 23.0, ROE 12.2, net margin 4.5. Same thin-margin logic
    # as retail - net margin at 10% removes almost the whole sector.
    'Consumer Staples': {
        'pe_max'            : 28    ,   # <- differs from the base set
        'roe_min'           : 10    ,
        'roic_min'          : None  ,   # <- differs from the base set
        'roce_min'          : None  ,   # <- differs from the base set
        'roa_min'           : 3     ,   # <- differs from the base set
        'gross_margin_min'  : 10    ,
        'op_margin_min'     : 8     ,   # <- differs from the base set
        'net_margin_min'    : 3     ,   # <- differs from the base set
        'fcf_margin_min'    : 4     ,   # <- differs from the base set
        'altman_min'        : 2     ,   # <- differs from the base set
        'piotroski_min'     : 6     ,   # <- differs from the base set
        'current_ratio_min' : 1     ,
        'de_max'            : 1.5   ,   # <- differs from the base set
        'int_cov_min'       : 3     ,
        'mcap_min'          : None  ,
    },

    # ---- Health Care -----------------------------------------------------------
    # Marquee medians: P/E 25.4, ROE 5.5, gross margin 60.3. Pharma carries heavy R&D
    # so ROE and ROA run low while gross margin runs very high.
    'Health Care': {
        'pe_max'            : 30    ,   # <- differs from the base set
        'roe_min'           : 8     ,   # <- differs from the base set
        'roic_min'          : None  ,   # <- differs from the base set
        'roce_min'          : None  ,   # <- differs from the base set
        'roa_min'           : 3     ,   # <- differs from the base set
        'gross_margin_min'  : 10    ,
        'op_margin_min'     : 12    ,   # <- differs from the base set
        'net_margin_min'    : 6     ,   # <- differs from the base set
        'fcf_margin_min'    : 8     ,   # <- differs from the base set
        'altman_min'        : 2.5   ,   # <- differs from the base set
        'piotroski_min'     : 5     ,   # <- differs from the base set
        'current_ratio_min' : 1.5   ,   # <- differs from the base set
        'de_max'            : 1     ,   # <- differs from the base set
        'int_cov_min'       : 4     ,   # <- differs from the base set
        'mcap_min'          : None  ,
    },

    # ---- Financials ------------------------------------------------------------
    # Banks: Altman Z, current ratio, debt to equity, interest coverage, gross margin and
    # FCF margin are all meaningless here and stay off. Median ROA is 1.5%, not 10%. Median
    # P/E is 14.2, so the cheapness bar can be tighter than the base rule.
    'Financials': {
        'pe_max'            : 16    ,   # <- differs from the base set
        'roe_min'           : 13    ,   # <- differs from the base set
        'roic_min'          : None  ,   # <- differs from the base set
        'roce_min'          : None  ,   # <- differs from the base set
        'roa_min'           : 1.5   ,   # <- differs from the base set
        'gross_margin_min'  : None  ,   # <- differs from the base set
        'op_margin_min'     : 25    ,   # <- differs from the base set
        'net_margin_min'    : 15    ,   # <- differs from the base set
        'fcf_margin_min'    : None  ,   # <- differs from the base set
        'altman_min'        : None  ,   # <- differs from the base set
        'piotroski_min'     : 5     ,   # <- differs from the base set
        'current_ratio_min' : None  ,   # <- differs from the base set
        'de_max'            : None  ,   # <- differs from the base set
        'int_cov_min'       : None  ,   # <- differs from the base set
        'mcap_min'          : None  ,
    },

    # ---- Energy ----------------------------------------------------------------
    # Marquee medians: P/E 10.6, ROE 13.9, ROA 6.9, operating margin 16.2. Cyclical, so
    # the P/E bar tightens while the margin bars loosen.
    'Energy': {
        'pe_max'            : 15    ,   # <- differs from the base set
        'roe_min'           : 10    ,
        'roic_min'          : None  ,   # <- differs from the base set
        'roce_min'          : None  ,   # <- differs from the base set
        'roa_min'           : 4     ,   # <- differs from the base set
        'gross_margin_min'  : 10    ,
        'op_margin_min'     : 12    ,   # <- differs from the base set
        'net_margin_min'    : 6     ,   # <- differs from the base set
        'fcf_margin_min'    : 5     ,   # <- differs from the base set
        'altman_min'        : 1.8   ,   # <- differs from the base set
        'piotroski_min'     : 6     ,   # <- differs from the base set
        'current_ratio_min' : 0.9   ,   # <- differs from the base set
        'de_max'            : 1     ,   # <- differs from the base set
        'int_cov_min'       : 3     ,
        'mcap_min'          : None  ,
    },

    # ---- Materials -------------------------------------------------------------
    # Marquee medians: P/E 21.7, ROE 9.2, ROA 4.4. Capital heavy commodity business.
    'Materials': {
        'pe_max'            : 25    ,   # <- differs from the base set
        'roe_min'           : 9     ,   # <- differs from the base set
        'roic_min'          : None  ,   # <- differs from the base set
        'roce_min'          : None  ,   # <- differs from the base set
        'roa_min'           : 4     ,   # <- differs from the base set
        'gross_margin_min'  : 10    ,
        'op_margin_min'     : 12    ,   # <- differs from the base set
        'net_margin_min'    : 5     ,   # <- differs from the base set
        'fcf_margin_min'    : 6     ,   # <- differs from the base set
        'altman_min'        : 2.5   ,   # <- differs from the base set
        'piotroski_min'     : 6     ,   # <- differs from the base set
        'current_ratio_min' : 1.5   ,   # <- differs from the base set
        'de_max'            : 1     ,   # <- differs from the base set
        'int_cov_min'       : 4     ,   # <- differs from the base set
        'mcap_min'          : None  ,
    },

    # ---- Industrials -----------------------------------------------------------
    # Marquee medians: P/E 27.8, ROE 13.7, ROA 4.9, operating margin 12.4.
    'Industrials': {
        'pe_max'            : 30    ,   # <- differs from the base set
        'roe_min'           : 13    ,   # <- differs from the base set
        'roic_min'          : None  ,   # <- differs from the base set
        'roce_min'          : None  ,   # <- differs from the base set
        'roa_min'           : 4     ,   # <- differs from the base set
        'gross_margin_min'  : 10    ,
        'op_margin_min'     : 12    ,   # <- differs from the base set
        'net_margin_min'    : 7     ,   # <- differs from the base set
        'fcf_margin_min'    : 8     ,   # <- differs from the base set
        'altman_min'        : 2.5   ,   # <- differs from the base set
        'piotroski_min'     : 6     ,   # <- differs from the base set
        'current_ratio_min' : 1.4   ,   # <- differs from the base set
        'de_max'            : 1     ,   # <- differs from the base set
        'int_cov_min'       : 4     ,   # <- differs from the base set
        'mcap_min'          : None  ,
    },

    # ---- Utilities -------------------------------------------------------------
    # Only 25 marquee names. Median FCF margin is negative (-7.4) because of capex and
    # median Altman is 1.1, so both are off or set very low. Median operating margin is 21.9,
    # which lets that bar sit higher than the base rule.
    'Utilities': {
        'pe_max'            : 28    ,   # <- differs from the base set
        'roe_min'           : 8     ,   # <- differs from the base set
        'roic_min'          : None  ,   # <- differs from the base set
        'roce_min'          : None  ,   # <- differs from the base set
        'roa_min'           : 2     ,   # <- differs from the base set
        'gross_margin_min'  : 10    ,
        'op_margin_min'     : 15    ,   # <- differs from the base set
        'net_margin_min'    : 8     ,   # <- differs from the base set
        'fcf_margin_min'    : None  ,   # <- differs from the base set
        'altman_min'        : 0.8   ,   # <- differs from the base set
        'piotroski_min'     : 5     ,   # <- differs from the base set
        'current_ratio_min' : None  ,   # <- differs from the base set
        'de_max'            : 3     ,   # <- differs from the base set
        'int_cov_min'       : 1.5   ,   # <- differs from the base set
        'mcap_min'          : None  ,
    },

    # ---- Real Estate -----------------------------------------------------------
    # REITs: median Altman 1.4 and current ratio 1.1, so those stay off. Median operating
    # margin 19.3 and net margin 13.2 are high, so those bars go above the base rule.
    'Real Estate': {
        'pe_max'            : 32    ,   # <- differs from the base set
        'roe_min'           : 7     ,   # <- differs from the base set
        'roic_min'          : None  ,   # <- differs from the base set
        'roce_min'          : None  ,   # <- differs from the base set
        'roa_min'           : 3     ,   # <- differs from the base set
        'gross_margin_min'  : 10    ,
        'op_margin_min'     : 18    ,   # <- differs from the base set
        'net_margin_min'    : 10    ,
        'fcf_margin_min'    : None  ,   # <- differs from the base set
        'altman_min'        : None  ,   # <- differs from the base set
        'piotroski_min'     : 5     ,   # <- differs from the base set
        'current_ratio_min' : None  ,   # <- differs from the base set
        'de_max'            : 2     ,
        'int_cov_min'       : 1.5   ,   # <- differs from the base set
        'mcap_min'          : None  ,
    },

}


# =====================================================================================
# [3b] SECTOR-TUNED RULES  ---  what 'tuned' means now
# =====================================================================================
# Every sector gets its own bars, computed from its own companies on each run: each rule
# sits at the median company of that sector, so a company passes only if it beats the
# typical peer in its own sector on every measure. No bar is shared across sectors.
# Every sector is judged only on the measures that describe that kind of business
# (SECTOR_RULES below).
#
# P/E is the one rule measured against a narrower group: the sector's quality peers, those
# with ROIC of at least 10% (ROE for banks, property and utilities). Loss makers and thin
# earners have inflated trailing P/Es and would otherwise drag the bar up.
#
# Percentiles are taken among companies above MCAP_FLOOR only, since those are the only
# ones that can pass.
#
# RULE_TARGETS: key -> (percentile, limit)
#   percentile  the share of the sector a company has to beat. 50 = the median.
#               None = use the limit as a fixed bar, the same for every sector.
#   limit       for '_min' rules the bar is never below it, for '_max' rules never above it.
#               None = no limit.
MCAP_FLOOR = 300e6           # US dollars. A fixed liquidity floor, never moved by the slider.

# Quality peers, the group P/E is compared against. ROE is used where ROIC means little.
PEER_ROIC_MIN = 10
PEER_BY_ROE = {'Financials', 'Real Estate', 'Utilities'}

RULE_TARGETS = {k: (50, None) for k in (
    'pe_max', 'roe_min', 'roic_min', 'roce_min', 'roa_min', 'gross_margin_min', 'op_margin_min',
    'net_margin_min', 'fcf_margin_min', 'altman_min', 'piotroski_min', 'current_ratio_min',
    'de_max', 'int_cov_min')}
# Fewer rules make a sector easier to pass: with 13 rules at the median about 3% of a sector
# gets through, with 5 rules about 10% does. Sectors judged on fewer rules therefore get a
# higher percentile on each, so passing is about as hard everywhere (2.5% to 3.5% of the
# sector, 22 October 2026 data). Financials 5 rules -> 70th, Real Estate and Utilities
# 9 rules -> 55th. The run log shows each sector's pass count; recheck these after big data changes.
SECTOR_PERCENTILE = {'Financials': 70, 'Real Estate': 55, 'Utilities': 55}
# Per sector replacements for RULE_TARGETS, built from SECTOR_PERCENTILE. Add hand entries
# the same way, e.g. SECTOR_TARGETS['Energy']['de_max'] = (40, None).
SECTOR_TARGETS = {g: {k: (t, None) for k in RULE_TARGETS} for g, t in SECTOR_PERCENTILE.items()}

# Every rule is used unless there is a stated reason not to (RULE_NOT_USED_WHY below).
_CORE = ['pe_max', 'roic_min', 'roe_min', 'roa_min', 'gross_margin_min', 'op_margin_min',
         'fcf_margin_min', 'net_margin_min', 'de_max', 'int_cov_min', 'current_ratio_min',
         'altman_min', 'piotroski_min', 'mcap_min']
SECTOR_RULES = {
    'All Sectors'            : _CORE,
    'Information Technology' : _CORE,
    'Communication Services' : _CORE,
    'Consumer Discretionary' : _CORE,
    'Consumer Staples'       : _CORE,
    'Health Care'            : _CORE,
    'Energy'                 : _CORE,
    'Materials'              : _CORE,
    'Industrials'            : _CORE,
    'Financials'             : ['pe_max', 'roe_min', 'roa_min', 'op_margin_min', 'net_margin_min',
                                'mcap_min'],
    'Real Estate'            : ['pe_max', 'roic_min', 'roe_min', 'roa_min', 'gross_margin_min',
                                'op_margin_min', 'fcf_margin_min', 'de_max', 'int_cov_min',
                                'mcap_min'],
    'Utilities'              : ['pe_max', 'roic_min', 'roe_min', 'roa_min', 'op_margin_min',
                                'net_margin_min', 'de_max', 'int_cov_min', 'piotroski_min',
                                'mcap_min'],
}
# Per rule shifts of the percentile, in points. Earnings in energy, materials and property
# swing with the cycle, so a low P/E at the top of the cycle misleads: their P/E bar sits at
# the 60th percentile of quality peers instead of the median.
RULE_TARGET_SHIFT = {
    ('Energy', 'pe_max'): -10, ('Materials', 'pe_max'): -10, ('Real Estate', 'pe_max'): -10,
}
# Why a rule is left out, shown on the web page next to "not used" so the choice is visible.
# (short label shown in the table, full reason shown on hover and under the rule).
# Coverage figures are the share of companies above $300M that TradingView reports the
# figure for, 22 October 2026 export.
_SAME_AS_ROIC = ('same as ROIC', 'Measures the same thing as ROIC, operating profit over the capital '
                 'in the business. Using both would count one idea twice.')
RULE_NOT_USED_WHY = {
    'All': {
        'roce_min': _SAME_AS_ROIC,
    },
    'Financials': {
        'roic_min': ('capital is the product', "A lender's invested capital is mostly money it lends "
                     'on, so ROIC has no clear meaning. ROE and ROA are the standard tests for banks '
                     'and insurers.'),
        'roce_min': ('capital is the product', "A lender's capital employed is mostly money it lends "
                     'on, so ROCE has no clear meaning. ROE and ROA are used instead.'),
        'gross_margin_min': ('no cost of goods', 'Banks and insurers have no cost of goods sold, so '
                             'there is no gross margin. TradingView reports it for only 31% of them.'),
        'fcf_margin_min': ('lending swings cash flow', 'Loans made and deposits taken run through '
                           'operating cash flow, so free cash flow moves with the loan book, not '
                           'with business quality.'),
        'current_ratio_min': ('no current split', "A bank's balance sheet is not split into current "
                              'and long-term items, so the ratio says nothing about liquidity.'),
        'de_max': ('debt is the raw material', 'Borrowing is how lenders make money, so a high debt to '
                   'equity is normal, not a warning. Capital ratios would be the right test and the '
                   'export does not carry them.'),
        'int_cov_min': ('interest is the cost of sales', 'For a lender, interest paid is the main cost '
                        'of doing business, so interest cover does not measure safety. TradingView '
                        'reports it for only 24% of them.'),
        'altman_min': ('not built for lenders', 'The Altman model was built on manufacturers and '
                       'misreads financial balance sheets. TradingView reports it for only 34% of them.'),
        'piotroski_min': ('too little data', 'TradingView reports it for only 6% of financial companies, '
                          'so the rule would wave almost everyone through unchecked.'),
    },
    'Real Estate': {
        'roce_min': ('too little data', 'TradingView reports it for only 19% of property companies, '
                     'and ROIC already covers it.'),
        'net_margin_min': ('revaluations distort', 'Property revaluations and depreciation swing net '
                           'profit without any change in rent, so net margin misleads. Operating '
                           'margin is used instead.'),
        'current_ratio_min': ('long-term assets', 'Landlords hold buildings for decades and fund them '
                              'long term, so short-term liquidity says little about risk. Debt and '
                              'interest cover do that job.'),
        'altman_min': ('not built for property', 'The Altman model penalises any asset-heavy, '
                       'debt-funded balance sheet, so the typical landlord scores 1.1 and would look '
                       'distressed.'),
        'piotroski_min': ('too little data', 'TradingView reports it for only 18% of property '
                          'companies, so the rule would wave most of them through unchecked.'),
    },
    'Utilities': {
        'roce_min': _SAME_AS_ROIC,
        'gross_margin_min': ('fuel passes through', 'Fuel and power costs are passed to customers '
                             'through tariffs, so gross margin moves with commodity prices, not with '
                             'the business.'),
        'fcf_margin_min': ('capex recovered via tariffs', 'Utilities spend heavily on grids and plants '
                           'and recover it through regulated tariffs over decades, so free cash flow '
                           'is often negative for healthy companies in building years.'),
        'current_ratio_min': ('thin by design', 'Steady, regulated cash flows let utilities run thin '
                              'working capital safely, so a low current ratio is not a warning sign.'),
        'altman_min': ('not built for utilities', 'The Altman model penalises asset-heavy, debt-funded '
                       'balance sheets, which is how every utility is built: the typical one scores 1.3.'),
    },
}


def not_used_why():
    """For the web page: {sector: {rule: [short, long]}} for every rule a sector leaves out."""
    out = {}
    for g in [ALL_LABEL] + list(SECTOR_MAP):
        used = SECTOR_RULES.get(g, _CORE)
        out[g] = {k: list(RULE_NOT_USED_WHY.get(g, {}).get(k) or RULE_NOT_USED_WHY['All'].get(k)
                          or ('not used', 'Not part of this sector\'s rules.'))
                  for k in DEFAULT_FILTERS if k not in used}
    return out


# Hand overrides on top of the computed bars, written as values: {'Energy': {'pe_max': 12}}
TUNED_OVERRIDES = {}

COMPUTED_TUNED = {}          # filled in at run time from the universe
DISTRIBUTIONS = {}           # 21 point distribution per sector and measure, for the web page
MARKET_DIST = {}             # the same, whole market and every company, for the base set
POSITIVE_ONLY = {'pe_max', 'de_max'}   # a loss or negative equity never passes these


def _measure(df, key):
    s = pd.to_numeric(df[FILTER_SPECS[key][0]], errors='coerce').dropna()
    if key in POSITIVE_ONLY:
        s = s[s > 0]
    return s


def rule_target(sector, key):
    """(percentile, limit) for one rule in one sector, after sector replacements and shifts."""
    t, lim = SECTOR_TARGETS.get(sector, {}).get(key, RULE_TARGETS.get(key, (50, None)))
    if t is not None:
        t += RULE_TARGET_SHIFT.get((sector, key), 0)
    return t, lim


def quality_peers(df):
    """Rows earning at least PEER_ROIC_MIN on ROIC, or on ROE in PEER_BY_ROE sectors."""
    roic = pd.to_numeric(df[FILTER_SPECS['roic_min'][0]], errors='coerce')
    roe = pd.to_numeric(df[FILTER_SPECS['roe_min'][0]], errors='coerce')
    by_roe = df['GICS Sector'].isin(PEER_BY_ROE)
    return df[(by_roe & (roe >= PEER_ROIC_MIN)) | (~by_roe & (roic >= PEER_ROIC_MIN))]


def compute_tuned(universe):
    """Set every tuned bar from the data, rule by rule, see RULE_TARGETS."""
    COMPUTED_TUNED.clear()
    DISTRIBUTIONS.clear()
    MARKET_DIST.clear()
    mcap = pd.to_numeric(universe[FILTER_SPECS['mcap_min'][0]], errors='coerce')
    investable = universe[mcap >= MCAP_FLOOR]
    for key in DEFAULT_FILTERS:
        s = _measure(investable, key)
        if len(s) >= 20:
            MARKET_DIST[key] = [round(float(s.quantile(q / 20)), 6) for q in range(21)]
    for sector in [ALL_LABEL] + list(SECTOR_MAP):
        sub = investable if sector == ALL_LABEL else investable[investable['GICS Sector'] == sector]
        dist, block = {}, {}
        peers = quality_peers(sub)
        for key in DEFAULT_FILTERS:
            s = _measure(peers if key == 'pe_max' else sub, key)
            if len(s) >= 20:
                dist[key] = [round(float(s.quantile(q / 20)), 6) for q in range(21)]
            if key not in SECTOR_RULES.get(sector, _CORE):
                block[key] = None
                continue
            if key == 'mcap_min':
                block[key] = MCAP_FLOOR
                continue
            t, lim = rule_target(sector, key)
            if t is None:
                block[key] = lim
                continue
            if len(s) < 20:
                block[key] = lim
                continue
            q = (1 - t / 100) if FILTER_SPECS[key][1] == 'max' else t / 100
            v = float(s.quantile(min(max(q, 0.0), 1.0)))
            if lim is not None:
                v = max(v, lim) if FILTER_SPECS[key][1] == 'min' else min(v, lim)
            if key == 'piotroski_min':
                v = float(math.ceil(v))
            block[key] = round(v, 4)
        block.update(TUNED_OVERRIDES.get(sector, {}))
        COMPUTED_TUNED[sector] = block
        DISTRIBUTIONS[sector] = dist


def fixed_rules():
    """Per sector, the rules held at a fixed number, which the strictness slider leaves alone."""
    return {g: [k for k in SECTOR_RULES.get(g, _CORE) if k != 'mcap_min' and rule_target(g, k)[0] is None]
            for g in [ALL_LABEL] + list(SECTOR_MAP)}


# =====================================================================================
# [5] GICS  <-  TradingView SECTOR MAPPING
# =====================================================================================
SECTOR_MAP = {
    'Information Technology' : ['Electronic technology', 'Technology services'],
    'Communication Services' : ['Communications'],
    'Consumer Discretionary' : ['Consumer durables', 'Consumer services', 'Retail trade'],
    'Consumer Staples'       : ['Consumer non-durables', 'Distribution services'],
    'Health Care'            : ['Health technology', 'Health services'],
    'Financials'             : ['Finance'],
    'Energy'                 : ['Energy minerals'],
    'Materials'              : ['Non-energy minerals', 'Process industries'],
    'Industrials'            : ['Producer manufacturing', 'Industrial services',
                                'Commercial services', 'Transportation'],
    'Utilities'              : ['Utilities'],
    'Real Estate'            : [],   # no separate TradingView sector - see the note below
}

# TradingView has no separate Real Estate sector, so REITs mostly sit inside Finance.
# This pattern pulls them out by industry name onto the Real Estate sheet and removes
# them from the Financials sheet.
REAL_ESTATE_INDUSTRY = re.compile(r'real estate|reit', re.I)
PULL_REITS_OUT_OF_FINANCE = True     # set False to leave the Real Estate sheet empty

# Industries TradingView files under the wrong sector for GICS purposes. Aerospace & defense
# sits inside "Electronic technology", which would put defense contractors in IT.
INDUSTRY_TO_GICS = {'Aerospace & defense': 'Industrials'}

ALL_LABEL = 'All Sectors'      # the combined, no-sector-split screen

# Excel sheet names cannot exceed 31 characters
SHEET_NAMES = {
    ALL_LABEL                : 'Screen - All Sectors',
    'Information Technology' : 'Screen - Info Technology',
    'Communication Services' : 'Screen - Communication Svcs',
    'Consumer Discretionary' : 'Screen - Cons Discretionary',
    'Consumer Staples'       : 'Screen - Cons Staples',
    'Health Care'            : 'Screen - Health Care',
    'Financials'             : 'Screen - Financials',
    'Energy'                 : 'Screen - Energy',
    'Materials'              : 'Screen - Materials',
    'Industrials'            : 'Screen - Industrials',
    'Utilities'              : 'Screen - Utilities',
    'Real Estate'            : 'Screen - Real Estate',
}

# =====================================================================================
# ==============  ENGINE BELOW THIS LINE - normally nothing to change here  ============
# =====================================================================================

# filter key -> (source column, direction, pretty label)
FILTER_SPECS = {
    'pe_max'            : ('Price to earnings ratio',                          'max', 'P/E (TTM)'),
    'roe_min'           : ('Return on equity %, Trailing 12 months',           'min', 'ROE %'),
    'roic_min'          : ('Return on invested capital %, Trailing 12 months',  'min', 'ROIC %'),
    'roce_min'          : ('Return on capital employed %, Trailing 12 months',  'min', 'ROCE %'),
    'roa_min'           : ('Return on assets %, Trailing 12 months',            'min', 'ROA %'),
    'gross_margin_min'  : ('Gross margin %, Trailing 12 months',                            'min', 'Gross margin %'),
    'op_margin_min'     : ('Operating margin %, Trailing 12 months',                        'min', 'Operating margin %'),
    'net_margin_min'    : ('Net margin %, Trailing 12 months',                              'min', 'Net margin %'),
    'fcf_margin_min'    : ('Free cash flow margin %, Trailing 12 months',                   'min', 'FCF margin %'),
    'altman_min'        : ('Altman Z-score, Trailing 12 months',                'min', 'Altman Z'),
    'piotroski_min'     : ('Piotroski F-score, Trailing 12 months',             'min', 'Piotroski F'),
    'current_ratio_min' : ('Current ratio, Quarterly',                          'min', 'Current ratio'),
    'de_max'            : ('Debt to equity ratio, Quarterly',                   'max', 'Debt / Equity'),
    'int_cov_min'       : ('Interest coverage, Trailing 12 months',                         'min', 'Interest coverage'),
    'mcap_min'          : ('Market capitalization',                             'min', 'Market cap'),
}

PCT, NUM, X, BIG, TXT = '0.0"%";(0.0"%");-', '#,##0.00', '0.0"x"', '#,##0', '@'
Z, F = '0.00', '0"/9"'
USD = '$#,##0.00'

thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
H1F = Font(name='Arial', size=10, bold=True, color='FFFFFF')
H2F = Font(name='Arial', size=9, bold=True, color='FFFFFF')
BF = Font(name='Arial', size=9)
BOLD = Font(name='Arial', size=9, bold=True)
CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center')
RIGHT = Alignment(horizontal='right')
STRIPE = PatternFill('solid', fgColor='F2F5FA')

NAVY, GOLD, PLUM = '1F3864', 'BF8F00', '4B2E83'


# ------------------------------------------------------------------ data loading -----
DATE_RE = re.compile(r'(20\d{2})[-_. ]?(\d{2})[-_. ]?(\d{2})')


def file_date(path):
    """The data date, taken from the filename if it carries one, else the file's own date."""
    m = DATE_RE.search(os.path.basename(path))
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    return datetime.date.fromtimestamp(os.path.getmtime(path))


def file_rank(path):
    """Sort key for 'which file is newest'. The date written in the filename wins, and the
    modified time only breaks ties. On GitHub Actions every file gets the same modified time
    at checkout, so ranking by modified time alone picks an arbitrary file."""
    return (file_date(path), os.path.getmtime(path))


def pretty_date(d):
    return f'{d.day} {d:%B %Y}'


def _header(path):
    """Column names of a CSV, lower-cased, without reading the whole file."""
    try:
        return [str(c).strip().lower().replace('\n', ' ')
                for c in pd.read_csv(path, nrows=0).columns]
    except Exception:
        return []


def classify(path):
    """Work out what a CSV is from its columns, not its filename.

    Returns 'universe', 'marquee' or None. This is why the files can be named
    anything - rename them freely, re-export them, keep the date suffix or drop it.
    """
    cols = _header(path)
    if not cols:
        return None
    joined = ' | '.join(cols)
    # The All Stocks export carries the screener's financial columns.
    universe_hits = sum(k in joined for k in
                        ('market capitalization', 'price to earnings', 'return on equity',
                         'net margin', 'country or region', 'sector'))
    # The Dataroma export is a short ownership table: ticker, name, weight, holder count.
    marquee_hits = sum(k in joined for k in
                       ('ownership', 'hold price', 'portfolio', '52 week', '52w', 'max %'))
    has_pct_col = any(c.strip().startswith('%') or c.strip('  ') in ('%', '%\u25bc') for c in cols)
    if universe_hits >= 3:
        return 'universe'
    if (marquee_hits >= 2 or (has_pct_col and 'stock' in joined)) and len(cols) <= 15:
        return 'marquee'
    return None


def find_files(data_dir):
    """Auto-detect the universe and marquee CSVs in a folder by their contents.

    Filenames are never used to decide what a file is. If several files of the same
    kind are present, the most recently modified one wins, so an old export left in
    the folder is harmless.
    """
    if not os.path.isdir(data_dir):
        sys.exit(f"\n[ERROR] The folder does not exist:\n        {os.path.abspath(data_dir)}\n"
                 f"        Fix: correct DATA_DIR at the top of the script.\n")
    csvs = sorted(glob.glob(os.path.join(data_dir, '*.csv')))
    found = {'universe': [], 'marquee': [], 'unknown': []}
    for path in csvs:
        found[classify(path) or 'unknown'].append(path)

    def newest(paths):
        return max(paths, key=file_rank) if paths else None

    uni, mrq = newest(found['universe']), newest(found['marquee'])
    if uni and mrq:
        for kind, chosen in (('universe', uni), ('marquee', mrq)):
            extra = [os.path.basename(x) for x in found[kind] if x != chosen]
            if extra:
                print(f'      note: {len(extra)} older {kind} file(s) ignored -> {extra}')
        if found['unknown']:
            print(f"      note: ignored {[os.path.basename(x) for x in found['unknown']]} "
                  f"(columns match neither export)")
        return uni, mrq

    listing = '\n'.join(
        f'          {os.path.basename(p):<45} detected as: {classify(p) or "unrecognised"}'
        for p in csvs) or '          no CSV files at all'
    missing = 'All Stocks screener export' if not uni else 'Dataroma marquee export'
    sys.exit(f"\n[ERROR] Could not find the {missing} in this folder.\n"
             f"        Looked in : {os.path.abspath(data_dir)}\n"
             f"        Saw       :\n{listing}\n"
             f"        The two files are identified by their columns, not their names.\n"
             f"        The All Stocks export needs columns such as 'Market capitalization' and\n"
             f"        'Price to earnings ratio'. The Dataroma export needs its ownership\n"
             f"        columns such as 'Ownership count' and 'Hold Price'.\n"
             f"        Fix       : re-export the missing file, or check DATA_DIR.\n")


def to_money(s):
    if pd.isna(s):
        return np.nan
    return pd.to_numeric(str(s).replace('$', '').replace(',', '').strip(), errors='coerce')


SUFFIX = (r'\b(incorporated|inc|corporation|corp|company|co|limited|ltd|plc|llc|lp|nv|n\.v|sa|s\.a'
          r'|ag|ab|asa|as|a/s|oyj|spa|s\.p\.a|holdings?|holding|group|grp|the|trust|reit|adr|ads'
          r'|sponsored|cl|class|series)\b')


def norm_name(n):
    s = str(n).lower().replace('&', ' and ')
    s = re.sub(r'[^a-z0-9/\. ]', ' ', s)
    s = re.sub(r'\.', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    for _ in range(4):
        s = re.sub(SUFFIX, ' ', s)
        s = re.sub(r'\s+', ' ', s).strip()
    for _ in range(2):
        s = re.sub(r'\b(a|b|c|k|h|i|ii)\b$', '', s).strip()
    return re.sub(r'\s+', ' ', s).strip()


# Hand-checked ADR / alternate share class map: dataroma ticker -> (universe ticker, country)
ALIAS = {
    'TSM': ('2330', 'Taiwan'), 'NVO': ('NOVO_B', 'Denmark'), 'NSRGY': ('NESN', 'Switzerland'),
    'DB': ('DBK', 'Germany'), 'SNY': ('SAN', 'France'), 'FMS': ('FME', 'Germany'),
    'LZAGY': ('LONN', 'Switzerland'), 'NGG': ('NG.', 'United Kingdom'),
    'RYCEY': ('RR.', 'United Kingdom'), 'LOGI': ('LOGN', 'Switzerland'),
    'FLUIF': ('FDR', 'Spain'), 'HMRZF': ('HM_B', 'Sweden'), 'PHJMF': ('HMSP', 'Indonesia'),
    'DVDCF': ('CPR', None), 'BUD': ('ABI', 'Belgium'), 'BUDFF': ('ABI', 'Belgium'),
    'ABEV': ('ABEV3', 'Brazil'), 'ITUB': ('ITUB3', 'Brazil'), 'VALE': ('VALE3', 'Brazil'),
    'EMBJ': ('EMBJ3', 'Brazil'), 'YPF': ('YPFD', 'Argentina'), 'TEO': ('TECO2', 'Argentina'),
    'CIB': ('CIBEST', 'Colombia'), 'FMX': ('FEMSA/UB', 'Mexico'), 'ASR': ('ASUR/B', 'Mexico'),
    'ZTO': ('2057', 'China'), 'SMFG': ('8316', 'Japan'), 'PBR.A': ('PETR3', 'Brazil'),
    'UHAL.B': ('UHAL', 'United States'), 'FWONK': ('FWONA', 'United States'),
    'LLYVK': ('LLYVA', 'United States'), 'GLIBK': ('GLIBA', 'United States'),
    'GLBAV-OLD': ('GLIBA', 'United States'), 'HTZWW': ('HTZ', 'United States'),
    'AUROW': ('AUR', 'United States'), 'ASPSW': ('ASPS', 'Luxembourg'),
    'ASPSZ': ('ASPS', 'Luxembourg'), 'ILLUU': ('ILLU', 'United States'),
}

ETF_TICKERS = set(
    'SPY VOO IWM GLD SGOV EWY FRDM QQQ DIA IVV VTI AGG TLT HYBB BIL SHV IEF LQD SLV IAU VEA VWO '
    'EFA IEMG JEPI SCHD VIG DVY XLF XLE XLK XLV XLI XLP XLU XLB XLY XLC VNQ IJH IJR MDY RSP ITOT '
    'SPLG VXUS BND VCIT VCSH JNK USHY EMB PFF ARKK SMH SOXX IBIT FBTC GBTC ETHE MOAT KWEB GDXJ '
    'DINT MAVF AMLP IWF'.split())
ETF_HINT = re.compile(r'\b(etf|ishares|spdr|vanguard|index|invesco|proshares|select sector|treasury'
                      r'|gold shares|fund|msci|s&p 500|russell|kraneshares|vaneck|alerian'
                      r'|davis select)\b', re.I)


ISIN_COL_HINT = re.compile(r'isin', re.I)


def find_isin_column(df):
    """The ISIN column, whatever TradingView decided to call it, or None."""
    for c in df.columns:
        if ISIN_COL_HINT.search(str(c)):
            return c
    return None


# ------------------------------------------------------------------ column resolver ---
# TradingView renames columns from time to time, most often by switching the period, for
# example "Gross margin %, Annual" becoming "Gross margin %, Trailing 12 months". The
# resolver below lets the script survive that without any code change. For every column
# the script needs, it looks in the CSV for
#   1. the exact name
#   2. the same name with different spacing or capital letters
#   3. a known synonym for the identity columns (Ticker for Symbol, Name for Description)
#   4. the same metric with a different period, preferring TTM, then Annual, then Quarterly
# Whatever it finds is copied under the name the script uses, and the run log says what
# was swapped. A metric that is missing entirely becomes a blank column and its filter is
# switched off for the run, so a dropped column can never crash the build or empty a screen.

PERIOD_PREFERENCE = ['trailing 12 months', 'annual', 'fiscal year', 'quarterly',
                     'fiscal quarter', 'most recent quarter']
PERIOD_SHORT = {'trailing 12 months': 'TTM', 'annual': 'Annual', 'fiscal year': 'Annual',
                'quarterly': 'Quarterly', 'fiscal quarter': 'Quarterly',
                'most recent quarter': 'Quarterly'}
PERIOD_WORDS = {'TTM': 'trailing twelve months', 'Annual': 'annual',
                'Quarterly': 'latest quarter'}
IDENTITY_SYNONYMS = {
    'Symbol': ['Ticker', 'Symbol name'],
    'Description': ['Name', 'Company', 'Company name'],
    'Country or region of registration': ['Country', 'Country or region'],
    'Market capitalization': ['Market cap'],
    'Market capitalization - Currency': ['Market cap - Currency'],
    '52 week high': ['High, 52 weeks', '52 Week High', '52-week high', 'Price - 52 week high',
                     'High 52 weeks', '52 weeks high'],
    'Relative volume': ['Relative volume, 10 days', 'Relative Volume', 'Relative volume 10 days',
                        'Relative volume at time'],
}
# without these the build cannot work at all, so their absence stops the run
OPTIONAL_COLS = {'Performance %, 1 month', 'Performance %, 3 months', '52 week high',
                 'Relative volume'}
ESSENTIAL_COLS = ['Symbol', 'Description', 'Country or region of registration',
                  'Sector', 'Industry', 'Price', 'Market capitalization']
COLUMN_SOURCE = {}     # name the script uses -> CSV column it was actually read from
MISSING_COLS = set()   # names the CSV does not carry in any form


def _norm_col(c):
    return re.sub(r'\s+', ' ', str(c)).strip().lower()


def _split_col(c):
    """'Gross margin %, Annual - Currency' -> ('gross margin %', 'annual', ' - currency')."""
    n = _norm_col(c)
    tail = ''
    if n.endswith(' - currency'):
        n, tail = n[:-len(' - currency')], ' - currency'
    base, period = n.rsplit(', ', 1) if ', ' in n else (n, '')
    return base, period, tail


def canonical_columns():
    """Every source column the script reads, gathered from its own column lists."""
    cols = set(ESSENTIAL_COLS) | {'Price - Currency', 'Market capitalization - Currency'}
    cols |= {spec[0] for spec in FILTER_SPECS.values()}
    cols |= {src for src, _ in DEEP_FEED_FIELDS}
    cols |= {src for src, _ in HTML_COLS}
    for _, _, items in metric_groups(''):
        cols |= {src for _, src, _ in items}
    derived = {'ISIN', 'GICS Sector', 'Marquee Flag', 'Marquee Wt %', 'Investors Holding'}
    return sorted(c for c in cols if c not in derived and not c.startswith('__'))


def period_of(name):
    """The period a column really covers in this run: TTM, Annual, Quarterly or ''."""
    _, p, _ = _split_col(COLUMN_SOURCE.get(name, name))
    return PERIOD_SHORT.get(p, '')


def period_words(name, fallback):
    return PERIOD_WORDS.get(period_of(name), fallback)


def resolve_columns(a):
    """Map whatever the CSV calls each column onto the names the script uses."""
    COLUMN_SOURCE.clear()
    MISSING_COLS.clear()
    present = {_norm_col(c): c for c in a.columns}
    by_base = {}
    for c in a.columns:
        b, p, t = _split_col(c)
        by_base.setdefault((b, t), []).append((p, c))

    def rank(item):
        p = item[0]
        return PERIOD_PREFERENCE.index(p) if p in PERIOD_PREFERENCE else len(PERIOD_PREFERENCE)

    swapped = []
    for canon in canonical_columns():
        if canon in a.columns:
            COLUMN_SOURCE[canon] = canon
            continue
        hit = present.get(_norm_col(canon))
        if hit is None:
            for alt in IDENTITY_SYNONYMS.get(canon, []):
                hit = present.get(_norm_col(alt))
                if hit is not None:
                    break
        if hit is None:
            b, p0, t = _split_col(canon)
            # only reporting periods are interchangeable (TTM, annual, quarterly). A 1 month
            # return is not a stand-in for a year to date return, so those never swap.
            cands = [c for c in by_base.get((b, t), []) if c[0] in PERIOD_PREFERENCE] \
                if p0 in PERIOD_PREFERENCE else []
            if cands:
                hit = sorted(cands, key=rank)[0][1]
        if hit is not None:
            a[canon] = a[hit]
            COLUMN_SOURCE[canon] = hit
            swapped.append((canon, hit))
        else:
            a[canon] = np.nan
            MISSING_COLS.add(canon)

    lost = [c for c in ESSENTIAL_COLS if c in MISSING_COLS]
    if lost:
        sys.exit(f"\n[ERROR] The All Stocks CSV is missing columns the script cannot run "
                 f"without: {lost}\n        Columns it does have: {list(a.columns)}\n"
                 f"        Fix: add them back in the TradingView screener and export again.\n")

    n_total = len(canonical_columns())
    print(f'      column check: {n_total - len(swapped) - len(MISSING_COLS)} of {n_total} '
          f'found as named, {len(swapped)} matched under a new name, '
          f'{len(MISSING_COLS - OPTIONAL_COLS)} missing')
    for canon, hit in swapped:
        print(f"        using '{hit}'  for  '{canon}'")
    optional_missing = sorted(MISSING_COLS & OPTIONAL_COLS)
    if optional_missing:
        print(f'        optional momentum columns not in this export: {optional_missing}. '
              f'Add them in TradingView to switch those momentum signals on.')
    for canon in sorted(MISSING_COLS - OPTIONAL_COLS):
        uses = [FILTER_SPECS[k][2] for k in FILTER_SPECS if FILTER_SPECS[k][0] == canon]
        tail = f' - the {uses[0]} filter is switched off this run' if uses else ''
        print(f"        [WARN] not in this export: '{canon}'{tail}")

    # relabel the deep feed periods so the app shows what the numbers really cover
    for src, key in DEEP_FEED_FIELDS:
        if DEEP_FEED_PERIODS.get(key) in ('TTM', 'Annual', 'Quarterly') and period_of(src):
            DEEP_FEED_PERIODS[key] = period_of(src)
    return a


# ------------------------------------------------------------------ ISIN memory -------
# TradingView exports do not always carry the ISIN column, and the deep analysis app is
# keyed by ISIN. So every ISIN the script ever sees is remembered in DATA_DIR/isin_map.json,
# keyed by ticker and country, and filled back in on runs whose export has no ISIN column.
# The memory is also seeded from the last published stocks.json, so ISINs from older
# exports are recovered automatically. In GitHub Actions the data folder is committed, so
# the memory survives from run to run.
ISIN_MEMORY_FILE = 'isin_map.json'
_ISIN_SHAPE = re.compile(r'^[A-Z]{2}[A-Z0-9]{9}[0-9]$')


def _isin_key(symbol, country):
    """Ticker plus country. Numeric tickers are normalised (2330.0 -> 2330, 005930 -> 5930)
    because older feeds stored them as numbers and lost the leading zeros."""
    t = str(symbol).strip().upper()
    if re.fullmatch(r'\d+(\.0+)?', t):
        t = str(int(float(t)))
    return f'{t}|{str(country).strip()}'


def fill_isins(a, data_dir, feed_dirs=()):
    path = os.path.join(data_dir, ISIN_MEMORY_FILE)
    memory = {}
    try:
        with open(path) as fh:
            memory = json.load(fh).get('map', {})
    except Exception:
        pass
    seeded = 0
    for folder in [d for d in feed_dirs if d]:
        try:
            with open(os.path.join(folder, 'stocks.json')) as fh:
                old = json.load(fh).get('stocks', {})
        except Exception:
            continue
        for key, e in old.items():
            isin = str(e.get('isin') or key).upper()
            if _ISIN_SHAPE.match(isin) and e.get('ticker') and e.get('country'):
                k = _isin_key(e['ticker'], e['country'])
                if k not in memory:
                    memory[k] = isin
                    seeded += 1
    keys = [_isin_key(s, c) for s, c in zip(a['Symbol'], a['Country or region of registration'])]
    learned = 0
    for k, isin in zip(keys, a['ISIN']):
        if isinstance(isin, str) and memory.get(k) != isin:
            memory[k] = isin
            learned += 1
    missing = a['ISIN'].isna()
    filled = pd.Series(keys, index=a.index)[missing].map(memory)
    a.loc[missing, 'ISIN'] = filled
    n_filled = int(filled.notna().sum())
    try:
        with open(path, 'w') as fh:
            json.dump({'updated': datetime.date.today().isoformat(), 'count': len(memory),
                       'map': memory}, fh, separators=(',', ':'), sort_keys=True)
    except Exception as e:
        print(f'      [WARN] could not save {ISIN_MEMORY_FILE}: {e}')
    print(f'      ISIN memory: {len(memory)} known ({learned} learned from this export, '
          f'{seeded} recovered from the last stocks.json), {n_filled} filled in, '
          f'{int(a["ISIN"].isna().sum())} rows still without one')
    return a


def load_universe(path, feed_dirs=()):
    a = pd.read_csv(path)
    a.columns = [re.sub(r'\s+', ' ', str(c)).strip() for c in a.columns]
    a = resolve_columns(a)
    a['_sym'] = a['Symbol'].astype(str).str.strip()
    a['_norm'] = a['Description'].map(norm_name)
    a['_us'] = (a['Country or region of registration'] == 'United States').astype(int)
    isin_col = find_isin_column(a)
    if isin_col:
        a['ISIN'] = a[isin_col].astype(str).str.strip().str.upper()
        good = a['ISIN'].str.fullmatch(r'[A-Z]{2}[A-Z0-9]{9}[0-9]', na=False)
        a.loc[~good, 'ISIN'] = None
        print(f"      ISIN column '{isin_col}' found - {int(good.sum())} of {len(a)} rows "
              f"carry a valid ISIN")
    else:
        a['ISIN'] = None
        print('      no ISIN column in this export - filling from the ISIN memory')
    a = fill_isins(a, os.path.dirname(os.path.abspath(path)), feed_dirs)
    return a


def load_marquee(path):
    d = pd.read_csv(path)
    d.columns = ['Symbol', 'Stock', 'Portfolio %', 'Ownership Count', 'Hold Price', 'Max %',
                 'Current Price', '52W Low', '% Above 52W Low', '52W High']
    for c in ('Hold Price', 'Current Price', '52W Low', '52W High'):
        d[c] = d[c].map(to_money)
    d = d.drop_duplicates(subset=['Symbol', 'Stock']).reset_index(drop=True)
    d['Symbol'] = d['Symbol'].astype(str).str.strip()
    d['Stock'] = d['Stock'].astype(str).str.strip()
    d['_norm'] = d['Stock'].map(norm_name)
    return d



# ------------------------------------------------------------------ dataroma scrape ---
def prune_old_marquee(data_dir, keep):
    """Delete every Dataroma file except the one in use."""
    if not KEEP_ONLY_LATEST_DATAROMA or not keep:
        return
    keep = os.path.abspath(keep)
    gone = []
    for path in glob.glob(os.path.join(data_dir, '*.csv')):
        if os.path.abspath(path) == keep:
            continue
        name = os.path.basename(path)
        if name.lower().startswith('dataroma_portfolio') or classify(path) == 'marquee':
            try:
                os.remove(path)
                gone.append(name)
            except OSError as e:
                print(f'      [WARN] could not delete {name}: {e}')
    if gone:
        print(f'      removed {len(gone)} older Dataroma file(s), keeping only '
              f'{os.path.basename(keep)}')


def scrape_dataroma(data_dir):
    """Pull the superinvestor aggregate holdings off dataroma.com into a CSV.

    Selenium is imported here rather than at the top of the file so the script still runs
    for anyone who only has the CSV exports and never scrapes.
    """
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait
    except ImportError:
        sys.exit('\n[ERROR] Selenium is not installed, so Dataroma cannot be scraped.\n'
                 '        Fix: pip install selenium     (Chrome must also be installed)\n'
                 "        Or set SCRAPE_DATAROMA = 'never' and put the CSV in DATA_DIR "
                 'by hand.\n')

    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--window-size=1400,1000')
    options.add_experimental_option('prefs', {
        'profile.managed_default_content_settings.images': 2,
        'profile.default_content_setting_values.notifications': 2,
    })

    url = 'https://www.dataroma.com/m/g/portfolio.php?L={}'
    headers, rows_out = [], []
    driver = None
    try:
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 20)
        for page in range(1, DATAROMA_PAGES + 1):
            got = 0
            for attempt in (1, 2, 3):
                try:
                    driver.get(url.format(page))
                    table = wait.until(EC.presence_of_element_located((By.ID, 'grid')))
                    trs = table.find_elements(By.TAG_NAME, 'tr')
                    if not trs:
                        break
                    if not headers:
                        headers = [td.text.strip()
                                   for td in trs[0].find_elements(By.TAG_NAME, 'td')]
                    for tr in trs[1:]:
                        cols = [td.text.strip() for td in tr.find_elements(By.TAG_NAME, 'td')]
                        if cols and len(cols) == len(headers):
                            rows_out.append(cols)
                            got += 1
                    break
                except Exception as e:
                    if attempt == 3:
                        print(f'      page {page} failed after 3 tries, skipped ({e.__class__.__name__})')
                    continue
            print(f'      page {page:>2} -> {got} holdings')
            if got == 0 and page > 1:
                print('      empty page, stopping early')
                break
    finally:
        if driver is not None:
            driver.quit()

    if not headers or len(rows_out) < 100:
        sys.exit(f'\n[ERROR] The Dataroma scrape returned only {len(rows_out)} rows, which '
                 f'means the page layout changed or the site blocked the request.\n'
                 f"        Fix: set SCRAPE_DATAROMA = 'never' and export the CSV by hand "
                 f'for now.\n')

    df = pd.DataFrame(rows_out, columns=headers).drop_duplicates()
    out = os.path.join(data_dir, f'dataroma_portfolio_{datetime.date.today():%Y-%m-%d}.csv')
    df.to_csv(out, index=False)
    print(f'      scraped {len(df)} unique holdings -> {os.path.basename(out)}')
    return out


def dataroma_needs_scrape(data_dir):
    """Decide whether to scrape, based on SCRAPE_DATAROMA and how old the local file is."""
    if SCRAPE_DATAROMA == 'always':
        return True, 'SCRAPE_DATAROMA is always'
    if SCRAPE_DATAROMA == 'never':
        return False, ''
    existing = [p for p in glob.glob(os.path.join(data_dir, '*.csv'))
                if classify(p) == 'marquee']
    if not existing:
        return True, 'no dataroma CSV in the folder'
    newest = max(existing, key=file_rank)
    age = (datetime.date.today() - file_date(newest)).days
    if age > DATAROMA_MAX_AGE_DAYS:
        return True, f'{os.path.basename(newest)} is {age} days old'
    return False, f'{os.path.basename(newest)} is {age} days old, fresh enough'



# ------------------------------------------------------------------ matching ---------
def match_marquee(a, d):
    by_sym, by_norm = a.groupby('_sym').groups, a.groupby('_norm').groups

    def pick(cands, px, country=None):
        if len(cands) == 0:
            return None
        c = cands
        if country is not None:
            cc = c[c['Country or region of registration'] == country]
            if len(cc):
                c = cc
        if len(c) == 1:
            return c.index[0]
        if not pd.isna(px):
            t = c.copy()
            t['_pd'] = (pd.to_numeric(t['Price'], errors='coerce') - px).abs() / px
            close = t[t['_pd'] < 0.02]
            if len(close):
                return close.sort_values('_pd').index[0]
        if c['_us'].max() == 1:
            c = c[c['_us'] == 1]
        return c.sort_values('Market capitalization', ascending=False).index[0]

    recs = []
    for _, r in d.iterrows():
        sym, px = r['Symbol'], r['Current Price']
        idx, method = None, 'No match'
        if sym in ALIAS:
            tgt, ctry = ALIAS[sym]
            if tgt in by_sym:
                idx = pick(a.loc[by_sym[tgt]], px, ctry)
                method = 'Verified cross-listing / share-class alias'
        if idx is None:
            for cs, tag in [(sym, 'Exact ticker'),
                            (re.sub(r'-OLD$', '', sym), 'Ticker (legacy -OLD stripped)'),
                            (sym.replace('.', '_'), 'Ticker variant'),
                            (sym.replace('.', '-'), 'Ticker variant')]:
                if cs in by_sym:
                    j = pick(a.loc[by_sym[cs]], px)
                    if j is not None:
                        idx, method = j, tag
                        break
        if idx is None and r['_norm'] and r['_norm'] in by_norm:
            j = pick(a.loc[by_norm[r['_norm']]], px)
            if j is not None:
                idx, method = j, 'Company name match'

        rec = dict(r)
        rec['Match Method'] = method
        if idx is not None:
            m = a.loc[idx]
            rec['Matched Symbol'], rec['Matched Company'] = m['Symbol'], m['Description']
            for c in a.columns:
                if not c.startswith('_'):
                    rec['A_' + c] = m[c]
            p = pd.to_numeric(m['Price'], errors='coerce')
            same_px = (not pd.isna(px)) and (not pd.isna(p)) and abs(p - px) / px < 0.03
            alt_class = (m['Symbol'] != sym) and (not pd.isna(px)) and (not pd.isna(p)) \
                        and abs(p - px) / px > 0.20
            if method == 'Exact ticker' and same_px:
                rec['Match Confidence'] = 'High'
            elif method == 'Exact ticker':
                rec['Match Confidence'] = 'High' if m['Price - Currency'] == 'USD' else 'Medium'
            elif method in ('Verified cross-listing / share-class alias', 'Company name match'):
                rec['Match Confidence'] = 'High' if same_px else 'Medium'
            else:
                rec['Match Confidence'] = 'Medium'
            if m['Price - Currency'] == 'USD' and alt_class:
                rec['Data Basis'] = (f"Alternate share class ({m['Symbol']}) - company ratios "
                                     f"identical, per-share figures relate to {m['Symbol']}")
            elif m['Price - Currency'] == 'USD':
                rec['Data Basis'] = ('US line item' if (pd.isna(px) or pd.isna(p)
                                     or abs(p - px) / px <= 0.20)
                                     else 'US line item - screener price differs from the marquee '
                                          'list quote, verify before use')
            else:
                rec['Data Basis'] = (f"Primary listing in {m['Country or region of registration']} "
                                     f"({m['Price - Currency']}) - ratios comparable, price and FCF "
                                     f"in local currency")
        else:
            rec['Matched Symbol'] = None
            if sym in ETF_TICKERS or ETF_HINT.search(r['Stock']):
                rec['Reason Not Identified'] = 'ETF / index fund / trust - no company-level financials exist'
            elif sym.endswith('-OLD'):
                rec['Reason Not Identified'] = 'Legacy or superseded ticker (-OLD) - not in the current equity universe'
            elif re.search(r'(WTS|W$|U$|Z$)', sym) or 'WTS' in r['Stock'].upper():
                rec['Reason Not Identified'] = 'Warrant / unit / non-common security - no underlying match found'
            else:
                rec['Reason Not Identified'] = 'Company not present in the All Stocks universe (ADR-only, OTC or delisted)'
        recs.append(rec)
    return pd.DataFrame(recs)


# ------------------------------------------------------------------ screening --------
def screen_order():
    """All Sectors first, then the eleven GICS sectors."""
    return [ALL_LABEL] + list(SECTOR_MAP)


def active_set():
    """The threshold dictionary THRESHOLD_MODE points at."""
    if THRESHOLD_MODE not in ('base', 'tuned'):
        sys.exit(f"[ERROR] THRESHOLD_MODE must be 'base' or 'tuned', not {THRESHOLD_MODE!r}.")
    return COMPUTED_TUNED if THRESHOLD_MODE == 'tuned' else SECTOR_OVERRIDES


def sector_filters(gics):
    """DEFAULT_FILTERS plus that sector's block from the active set = the final rules."""
    f = dict(DEFAULT_FILTERS)
    f.update(active_set().get(gics, {}))
    return f


def assign_gics(df):
    tv2gics = {}
    for g, tvs in SECTOR_MAP.items():
        for tv in tvs:
            tv2gics[tv.strip().lower()] = g
    g = df['Sector'].astype(str).str.strip().str.lower().map(tv2gics)
    ind = df['Industry'].astype(str).str.strip()
    for name, target in INDUSTRY_TO_GICS.items():
        g = g.where(ind != name, target)
    if PULL_REITS_OUT_OF_FINANCE:
        is_re = df['Industry'].astype(str).str.contains(REAL_ESTATE_INDUSTRY)
        g = g.where(~is_re, 'Real Estate')
    return g


def screen(df, gics):
    f = sector_filters(gics)
    mask = pd.Series(True, index=df.index)
    reported = pd.Series(0, index=df.index)
    applied, n_active = [], 0
    for key, val in f.items():
        if val is None:
            continue
        col, direction, label = FILTER_SPECS[key]
        if col in MISSING_COLS:
            continue
        series = pd.to_numeric(df[col], errors='coerce')
        ok = (series >= val) if direction == 'min' else (series <= val)
        if key in POSITIVE_ONLY:          # a loss-making P/E or negative equity is not cheap
            ok = ok & (series > 0)
        if not BLANK_FAILS:
            ok = ok | series.isna()
        mask &= ok.fillna(False)
        reported += series.notna().astype(int)
        n_active += 1
        applied.append((label, ('>=' if direction == 'min' else '<='), val))
    if not BLANK_FAILS and n_active and MIN_DATA_COVERAGE > 0:
        mask &= (reported / n_active) >= MIN_DATA_COVERAGE
    return df[mask].copy(), applied, f


# ------------------------------------------------------------------ excel helpers ----
def write_grid(ws, df, groups, freeze_col=3, start_row=1):
    """Two-row banded header + striped data rows. Returns (ncols, {src: col_index})."""
    col, fmts = 1, []
    hr1, hr2 = start_row, start_row + 1
    for gname, gcolor, cols in groups:
        start = col
        for label, src, fmt in cols:
            c = ws.cell(row=hr2, column=col, value=label)
            c.font, c.fill = H2F, PatternFill('solid', fgColor=gcolor)
            c.alignment, c.border = CEN, BORDER
            fmts.append((col, src, fmt))
            col += 1
        g = ws.cell(row=hr1, column=start, value=gname)
        g.font, g.fill = H1F, PatternFill('solid', fgColor=gcolor)
        g.alignment, g.border = CEN, BORDER
        if col - 1 > start:
            ws.merge_cells(start_row=hr1, start_column=start, end_row=hr1, end_column=col - 1)
        for cc in range(start, col):
            ws.cell(row=hr1, column=cc).fill = PatternFill('solid', fgColor=gcolor)
            ws.cell(row=hr1, column=cc).border = BORDER
    ncols = col - 1
    for i, (_, row) in enumerate(df.iterrows()):
        r = hr2 + 1 + i
        band = (i % 2 == 1)
        for cidx, src, fmt in fmts:
            c = ws.cell(row=r, column=cidx)
            if not src.startswith('__'):
                v = row.get(src, None)
                if isinstance(v, float) and pd.isna(v):
                    v = None
                if isinstance(v, str) and v.strip() == '':
                    v = None
                c.value = v
            c.font, c.number_format, c.border = BF, fmt, BORDER
            c.alignment = LEFT if fmt == TXT else RIGHT
            if band:
                c.fill = STRIPE
    ws.freeze_panes = ws.cell(row=hr2 + 1, column=freeze_col)
    ws.auto_filter.ref = f"A{hr2}:{get_column_letter(ncols)}{hr2 + len(df)}"
    ws.row_dimensions[hr1].height = 20
    ws.row_dimensions[hr2].height = 34
    for cidx, src, fmt in fmts:
        hdr = str(ws.cell(row=hr2, column=cidx).value)
        w = max(9, min(30, len(hdr) + 2))
        if fmt == TXT:
            w = max(w, 14)
        if src in ('Stock', 'A_Description', 'Description', 'A_Industry', 'Industry',
                   'Data Basis', 'Match Method', 'Reason Not Identified'):
            w = 34
        ws.column_dimensions[get_column_letter(cidx)].width = w
    ws.sheet_view.showGridLines = False
    return ncols, {src: cidx for cidx, src, _ in fmts}


def fill_cagr(ws, pos, src5, rows):
    """5Y cumulative % -> annualised %, as a live formula."""
    src_col = get_column_letter(pos[src5])
    tgt = pos['__CAGR5__']
    for r in rows:
        ref = f'{src_col}{r}'
        ws.cell(row=r, column=tgt, value=(
            f'=IF(N({ref})=0,IF({ref}="","",0),'
            f'IF({ref}<=-100,"",IFERROR((POWER(1+{ref}/100,1/5)-1)*100,"")))'))


def metric_groups(prefix=''):
    """Financial column groups. prefix='A_' for the merged marquee frame."""
    p = prefix
    return [
        ('Size', '7F6000', [('Market Cap', p + 'Market capitalization', BIG),
                            ('Ccy', p + 'Market capitalization - Currency', TXT)]),
        ('Valuation', GOLD, [('P/E (TTM)', p + 'Price to earnings ratio', X),
                             ('P/E (Forward)', p + 'Price to earnings ratio forward', X)]),
        ('Return Ratios', '1F6B3B', [
            ('ROIC %', p + 'Return on invested capital %, Trailing 12 months', PCT),
            ('ROCE %', p + 'Return on capital employed %, Trailing 12 months', PCT),
            ('ROA %', p + 'Return on assets %, Trailing 12 months', PCT),
            ('ROE %', p + 'Return on equity %, Trailing 12 months', PCT)]),
        ('Profitability Margins', '2F855A', [
            ('Gross Margin %', p + 'Gross margin %, Trailing 12 months', PCT),
            ('Operating Margin %', p + 'Operating margin %, Trailing 12 months', PCT),
            ('Net Margin %', p + 'Net margin %, Trailing 12 months', PCT),
            ('FCF Margin %', p + 'Free cash flow margin %, Trailing 12 months', PCT)]),
        ('Cash Flow', '116466', [('Free Cash Flow', p + 'Free cash flow, Trailing 12 months', BIG),
                                 ('FCF Ccy', p + 'Free cash flow, Trailing 12 months - Currency', TXT)]),
        ('Balance Sheet & Solvency', '7B3F00', [
            ('Current Ratio', p + 'Current ratio, Quarterly', X),
            ('Debt / Equity', p + 'Debt to equity ratio, Quarterly', X),
            ('Interest Coverage', p + 'Interest coverage, Trailing 12 months', X),
            ('EBITDA Int. Coverage', p + 'EBITDA interest coverage, Trailing 12 months', X)]),
        ('Quality & Distress Scores', PLUM, [
            ('Altman Z-Score', p + 'Altman Z-score, Trailing 12 months', Z),
            ('Piotroski F-Score', p + 'Piotroski F-score, Trailing 12 months', F)]),
        ('Price Performance', '7A1F3D', [
            ('1M %', p + 'Performance %, 1 month', PCT),
            ('3M %', p + 'Performance %, 3 months', PCT),
            ('6M %', p + 'Performance %, 6 months', PCT),
            ('YTD %', p + 'Performance %, Year to date', PCT),
            ('1Y %', p + 'Performance %, 1 year', PCT),
            ('5Y Cumulative %', p + 'Performance %, 5 years', PCT),
            ('5Y Annualised % (CAGR)', '__CAGR5__', PCT)]),
        ('Price & Technicals', '3F3F76', [
            ('Price', p + 'Price', NUM), ('Ccy', p + 'Price - Currency', TXT),
            ('EMA 50', p + 'Exponential moving average, 50, 1 day', NUM),
            ('EMA 200', p + 'Exponential moving average, 200, 1 day', NUM),
            ('RSI 14', p + 'Relative strength index, 14, 1 day', '0.0')]),
    ]


# ------------------------------------------------------------------ sheet builders ---
def sheet_readme(wb, n_matched, n_unid, n_all, screen_counts, stamp):
    ws = wb.active
    ws.title = 'Read Me'
    ws.sheet_view.showGridLines = False
    ws['B2'] = 'Marquee Investor Portfolio and GICS Sector Screener'
    ws['B2'].font = Font(name='Arial', size=16, bold=True, color=NAVY)
    ws['B3'] = (f"Updated as of {stamp['data']}  |  marquee holdings list as of {stamp['marquee']}"
                f"  |  workbook built {stamp['built']} by marquee_screener.py")
    ws['B3'].font = Font(name='Arial', size=10, italic=True, color='595959')

    passes = ', '.join(f'{g} {n}' for g, n in screen_counts.items())
    blocks = [
        ('Contents', [
            ('Marquee Portfolio', f'{n_matched} marquee holdings matched to the universe, financials grouped by discipline.'),
            ('Unidentified in Marquee', f'{n_unid} marquee line items with no company-level financial match, each with a reason.'),
            ('Sector Summary', 'Live formula roll-up of the Marquee Portfolio sheet by TradingView sector.'),
            ('All Stocks', f'{n_all} companies - the full source universe, flagged for marquee ownership.'),
            ('Screen Settings', 'The exact threshold set applied to each GICS sector, with overrides flagged.'),
            ('Screen - All Sectors', 'Every row judged against its own GICS sector block and the results combined, so this sheet is the union of the eleven sector sheets. A GICS column on every row shows which rule set applied. Rows whose TradingView sector maps to no GICS sector are judged on the All Sectors block. Set ALL_SECTORS_USES_SECTOR_RULES to False to apply one market-wide rule instead.'),
            ('Screen - <sector> sheets', 'One sheet per GICS sector, split into four labelled blocks: marquee holdings US, marquee holdings non-US, all other stocks US, all other stocks non-US. Each block is sorted by market cap.'),
        ]),
        ('Column groups', [
            ('Identification', 'Ticker, company, sector, industry, country and how the record was matched.'),
            ('Superinvestor Ownership', 'Aggregate portfolio weight, number of investors holding, largest single weight, average hold price and 52-week context.'),
            ('Size', 'Market capitalisation and its reporting currency.'),
            ('Valuation', 'P/E trailing twelve months and forward P/E.'),
            ('Return Ratios', 'ROIC, ROCE, ROA and ROE, ' + period_words(
                'Return on equity %, Trailing 12 months', 'trailing twelve months') + '.'),
            ('Profitability Margins', 'Gross, operating, net and free cash flow margins, '
                + period_words('Gross margin %, Trailing 12 months', 'trailing twelve months') + '.'),
            ('Cash Flow', 'Free cash flow in the reporting currency, ' + period_words(
                'Free cash flow, Trailing 12 months', 'trailing twelve months') + '.'),
            ('Balance Sheet & Solvency', 'Current ratio, debt to equity, interest coverage and EBITDA interest coverage.'),
            ('Quality & Distress Scores', 'Altman Z-score for bankruptcy risk - below 1.8 is distress, above 3.0 is safe, and the bands do not apply to banks or insurers. Piotroski F-score out of 9 for fundamental momentum.'),
            ('Price Performance', 'Year to date, six month and one year performance, plus the five year figure both cumulative and annualised.'),
            ('Price & Technicals', 'Last price, 50 and 200 day exponential moving averages and 14 day RSI.'),
        ]),
        ('How the marquee match was made (in priority order)', [
            ('1. Verified cross-listing alias', 'A hand-checked map for ADRs and alternate share classes, for example TSM to 2330 in Taiwan and BUD to ABI in Belgium.'),
            ('2. Exact ticker', 'Symbol present in both files. Where a symbol appears on several exchanges, the row whose price is within 2% of the marquee quote wins, then US registration, then largest market cap.'),
            ('3. Ticker variant', 'Legacy -OLD suffix removed, or dot and underscore share-class separators swapped.'),
            ('4. Company name match', 'Normalised company name equality after stripping legal suffixes and share-class labels.'),
            ('No fuzzy matching', 'Approximate name similarity was tested and rejected - short legal names such as "AB S.A." score as perfect matches against unrelated companies. Anything unconfirmed is reported as unidentified rather than guessed.'),
        ]),
        ('The sector screens', [
            ('GICS mapping', 'TradingView sectors are folded into the eleven GICS sectors. Real Estate has no TradingView sector, so REITs are pulled out of Finance and other sectors by industry name.'),
            ('Unmapped names', 'About 800 companies sit in the TradingView "Miscellaneous" or "Government" sectors, which map to no GICS sector. They appear on All Stocks and can pass the All Sectors screen, but they reach no individual sector sheet. That is why the sector sheets do not sum to the All Sectors count.'),
            ('Two threshold sets', "THRESHOLD_MODE selects one. 'base' gives every sector the same bars, the median company of the whole market. 'tuned', the default, gives each sector its own bars, the median company of that sector, on only the measures that fit its business; P/E is compared with the sector's quality peers."),
            ('Active set this run', f"THRESHOLD_MODE = '{THRESHOLD_MODE}'. See the Screen Settings sheet for the exact numbers applied to each sector."),
            ('Blank policy', f'BLANK_FAILS is currently {BLANK_FAILS}. ' + ('Missing values are ignored, so a stock is judged only on the metrics it reports.' if not BLANK_FAILS else 'Any stock missing a screened metric is rejected, which is strict given Altman covers about three quarters of the universe.') + f' A stock must also report at least {int(MIN_DATA_COVERAGE*100)}% of the screened metrics to qualify, so thinly covered names cannot pass by default.'),
            ('Passing counts this run', passes),
            ('To change a threshold', "Open marquee_screener.py and edit SECTOR_OVERRIDES (base), or for the tuned set change RULE_TARGETS, SECTOR_TARGETS, SECTOR_RULES or TUNED_OVERRIDES, then run the script again. Sector_Screener.html lets you test values live and copy a ready-made block."),
        ]),
        ('Reading the numbers', [
            ('Ratios are currency neutral', 'P/E, returns, margins and coverage compare directly across rows regardless of listing currency.'),
            ('Absolute figures are not', 'Market capitalisation, free cash flow and price carry their own currency column. Do not sum across currencies without converting.'),
            ('Percentages', 'Stored as the source reported them, so 55.6 means 55.6%.'),
            ('5Y Annualised % (CAGR)', 'A live formula on the cumulative column. Price only, so dividends are excluded, and it assumes a full five years of history. Blank where the source is blank or the loss is 100% or worse.'),
            ('Two as-of dates', 'Ownership columns carry the marquee export date, all financial and price columns the screener export date. Use Price, not Current Price, for anything live.'),
            ('Altman averages are banded', 'Sector Summary averages Altman Z only between -50 and 100 - a few leveraged financials carry scores in the tens of thousands that would swamp the mean. The detail sheets show every value as reported.'),
        ]),
    ]
    r = 5
    for title, items in blocks:
        ws.cell(row=r, column=2, value=title).font = Font(name='Arial', size=12, bold=True, color=GOLD)
        r += 1
        for k, v in items:
            ws.cell(row=r, column=2, value=k).font = Font(name='Arial', size=10, bold=True)
            c = ws.cell(row=r, column=3, value=v)
            c.font = Font(name='Arial', size=10)
            c.alignment = Alignment(wrap_text=True, vertical='top')
            ws.row_dimensions[r].height = 28
            r += 1
        r += 1
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 34
    ws.column_dimensions['C'].width = 118


def sheet_marquee(wb, matched):
    groups = [('Identification', NAVY, [
        ('Ticker', 'Symbol', TXT), ('Company (Marquee list)', 'Stock', TXT),
        ('Matched Ticker', 'Matched Symbol', TXT), ('Matched Company', 'A_Description', TXT),
        ('Sector', 'A_Sector', TXT), ('GICS Sector', 'GICS Sector', TXT),
        ('Industry', 'A_Industry', TXT),
        ('Country', 'A_Country or region of registration', TXT),
        ('Match Method', 'Match Method', TXT), ('Match Confidence', 'Match Confidence', TXT),
        ('Data Basis', 'Data Basis', TXT)]),
        ('Superinvestor Ownership', '2E5C8A', [
            ('Portfolio Wt %', 'Portfolio %', PCT), ('No. of Investors', 'Ownership Count', BIG),
            ('Max Single Wt %', 'Max %', PCT), ('Avg Hold Price', 'Hold Price', USD),
            ('Current Price (US line)', 'Current Price', USD),
            ('52W Low', '52W Low', USD), ('52W High', '52W High', USD),
            ('% Above 52W Low', '% Above 52W Low', PCT)])] + metric_groups('A_')
    ws = wb.create_sheet('Marquee Portfolio')
    _, pos = write_grid(ws, matched, groups)
    fill_cagr(ws, pos, 'A_Performance %, 5 years', range(3, 3 + len(matched)))
    return ws, pos


def sheet_unidentified(wb, unid):
    groups = [('Identification', NAVY, [('Ticker', 'Symbol', TXT),
                                        ('Company (Marquee list)', 'Stock', TXT)]),
              ('Superinvestor Ownership', '2E5C8A', [
                  ('Portfolio Wt %', 'Portfolio %', PCT),
                  ('No. of Investors', 'Ownership Count', BIG),
                  ('Max Single Wt %', 'Max %', PCT), ('Avg Hold Price', 'Hold Price', USD),
                  ('Current Price', 'Current Price', USD), ('52W Low', '52W Low', USD),
                  ('52W High', '52W High', USD), ('% Above 52W Low', '% Above 52W Low', PCT)]),
              ('Diagnosis', '7A1F3D', [('Reason Not Identified', 'Reason Not Identified', TXT)])]
    ws = wb.create_sheet('Unidentified in Marquee')
    write_grid(ws, unid, groups)


def sheet_sector_summary(wb, matched, pos):
    ws = wb.create_sheet('Sector Summary')
    ws.sheet_view.showGridLines = False
    last = 2 + len(matched)
    sc = get_column_letter(pos['A_Sector'])
    srng = f"'Marquee Portfolio'!${sc}$3:${sc}${last}"

    def rng(src):
        L = get_column_letter(pos[src])
        return f"'Marquee Portfolio'!${L}$3:${L}${last}"

    heads = [('Sector', None), ('Holdings', None), ('Aggregate Portfolio Wt %', 'Portfolio %'),
             ('Avg P/E (TTM)', 'A_Price to earnings ratio'),
             ('Avg ROE %', 'A_Return on equity %, Trailing 12 months'),
             ('Avg ROIC %', 'A_Return on invested capital %, Trailing 12 months'),
             ('Avg Operating Margin %', 'A_Operating margin %, Trailing 12 months'),
             ('Avg Net Margin %', 'A_Net margin %, Trailing 12 months'),
             ('Avg 1Y Performance %', 'A_Performance %, 1 year'),
             ('Avg 5Y Annualised %', '__CAGR5__'),
             ('Avg Altman Z (-50 to 100 band)', 'A_Altman Z-score, Trailing 12 months'),
             ('Avg Piotroski F', 'A_Piotroski F-score, Trailing 12 months')]
    ws['A1'] = 'Marquee Portfolio - Sector Roll-up'
    ws['A1'].font = Font(name='Arial', size=13, bold=True, color=NAVY)
    ws['A2'] = ('Every figure is a live formula reading the Marquee Portfolio sheet. Averages are '
                'simple, not weighted, and skip blanks.')
    ws['A2'].font = Font(name='Arial', size=9, italic=True, color='595959')
    for j, (h, _) in enumerate(heads, start=1):
        c = ws.cell(row=4, column=j, value=h)
        c.font, c.fill, c.alignment, c.border = H2F, PatternFill('solid', fgColor=NAVY), CEN, BORDER

    def fmt_for(h):
        return X if 'P/E' in h else (Z if 'Altman' in h else (F if 'Piotroski' in h else PCT))

    sectors = sorted(s for s in matched['A_Sector'].dropna().unique())
    for i, s in enumerate(sectors):
        r = 5 + i
        ws.cell(row=r, column=1, value=s).font = BF
        ws.cell(row=r, column=2, value=f'=COUNTIFS({srng},$A{r})').number_format = BIG
        for j, (h, src) in enumerate(heads[2:], start=3):
            if j == 3:
                f = f'=IFERROR(SUMIFS({rng(src)},{srng},$A{r}),0)'
            elif 'Altman' in h:
                f = (f'=IFERROR(AVERAGEIFS({rng(src)},{srng},$A{r},{rng(src)},">-50",'
                     f'{rng(src)},"<100"),"")')
            else:
                f = f'=IFERROR(AVERAGEIFS({rng(src)},{srng},$A{r}),"")'
            ws.cell(row=r, column=j, value=f).number_format = fmt_for(h)
        for j in range(1, len(heads) + 1):
            c = ws.cell(row=r, column=j)
            c.font, c.border = BF, BORDER
            if i % 2:
                c.fill = STRIPE
    tr = 5 + len(sectors)
    ws.cell(row=tr, column=1, value='Total / Portfolio').font = BOLD
    ws.cell(row=tr, column=2, value=f'=SUM(B5:B{tr - 1})').number_format = BIG
    ws.cell(row=tr, column=3, value=f'=SUM(C5:C{tr - 1})').number_format = PCT
    for j, (h, src) in enumerate(heads[3:], start=4):
        f = (f'=IFERROR(AVERAGEIFS({rng(src)},{rng(src)},">-50",{rng(src)},"<100"),"")'
             if 'Altman' in h else f'=IFERROR(AVERAGE({rng(src)}),"")')
        ws.cell(row=tr, column=j, value=f).number_format = fmt_for(h)
    for j in range(1, len(heads) + 1):
        c = ws.cell(row=tr, column=j)
        c.font, c.border = BOLD, BORDER
        c.fill = PatternFill('solid', fgColor='FFF2CC')
    ws.column_dimensions['A'].width = 30
    for j in range(2, len(heads) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 16
    ws.row_dimensions[4].height = 32
    ws.freeze_panes = 'B5'


def sheet_all_stocks(wb, a, matched):
    keys = set(zip(matched['Matched Symbol'], matched['A_Description']))
    a2 = a.drop(columns=[c for c in a.columns if c.startswith('_')]).copy()
    a2['Marquee Flag'] = ['Marquee holding' if (s, d) in keys else ''
                          for s, d in zip(a['Symbol'], a['Description'])]
    a2['GICS Sector'] = assign_gics(a2)
    a2 = a2.sort_values('Market capitalization', ascending=False).reset_index(drop=True)
    groups = [('Identification', NAVY, [
        ('Ticker', 'Symbol', TXT), ('Company', 'Description', TXT),
        ('Sector', 'Sector', TXT), ('GICS Sector', 'GICS Sector', TXT),
        ('Industry', 'Industry', TXT),
        ('Country', 'Country or region of registration', TXT),
        ('Marquee Flag', 'Marquee Flag', TXT)])] + metric_groups('')
    ws = wb.create_sheet('All Stocks')
    _, pos = write_grid(ws, a2, groups)
    fill_cagr(ws, pos, 'Performance %, 5 years', range(3, 3 + len(a2)))
    return a2


def sheet_screen_settings(wb, screen_counts, stamp):
    ws = wb.create_sheet('Screen Settings')
    ws.sheet_view.showGridLines = False
    ws['A1'] = f"Active screening thresholds by GICS sector   |   updated as of {stamp['data']}"
    ws['A1'].font = Font(name='Arial', size=13, bold=True, color=NAVY)
    src = ('the data: each bar is the median of its own sector, see RULE_TARGETS (sector-tuned set)') if THRESHOLD_MODE == 'tuned' else \
          'SECTOR_OVERRIDES (base set, the whole-market median for every sector)'
    ws['A2'] = (f'THRESHOLD_MODE = {THRESHOLD_MODE!r}, so these come from {src}. Edit that block in '
                'marquee_screener.py and run again. "off" means the filter is not applied, and '
                'values in gold differ from DEFAULT_FILTERS.')
    ws['A2'].font = Font(name='Arial', size=9, italic=True, color='595959')
    keys = list(DEFAULT_FILTERS.keys())
    heads = ['GICS Sector', 'Names passing'] + [FILTER_SPECS[k][2] +
             (' <=' if FILTER_SPECS[k][1] == 'max' else ' >=') for k in keys]
    for j, h in enumerate(heads, start=1):
        c = ws.cell(row=4, column=j, value=h)
        c.font, c.fill, c.alignment, c.border = H2F, PatternFill('solid', fgColor=NAVY), CEN, BORDER
    r = 5
    ws.cell(row=r, column=1, value='DEFAULT (base rules)').font = BOLD
    ws.cell(row=r, column=2, value='-').font = BF
    for j, k in enumerate(keys, start=3):
        v = DEFAULT_FILTERS[k]
        c = ws.cell(row=r, column=j, value='off' if v is None else v)
        c.font = BOLD
    for j in range(1, len(heads) + 1):
        c = ws.cell(row=r, column=j)
        c.border, c.alignment = BORDER, RIGHT if j > 2 else LEFT
        c.fill = PatternFill('solid', fgColor='FFF2CC')
    for gics in screen_order():
        r += 1
        f = sector_filters(gics)
        ov = {k: v for k, v in active_set().get(gics, {}).items() if v != DEFAULT_FILTERS[k]}
        ws.cell(row=r, column=1, value=gics).font = BF
        ws.cell(row=r, column=2, value=screen_counts.get(gics, 0)).number_format = BIG
        for j, k in enumerate(keys, start=3):
            v = f[k]
            c = ws.cell(row=r, column=j, value='off' if v is None else v)
            c.font = Font(name='Arial', size=9, bold=True, color='9C6500') if k in ov else BF
        for j in range(1, len(heads) + 1):
            c = ws.cell(row=r, column=j)
            c.border = BORDER
            c.alignment = RIGHT if j > 2 else LEFT
            if (r - 5) % 2:
                c.fill = STRIPE
    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 14
    for j in range(3, len(heads) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 14
    ws.row_dimensions[4].height = 34
    ws.freeze_panes = 'C5'


def sheet_sector_screen(wb, gics, passed, rule_text, universe_label, stamp):
    """One sheet per sector, split into four labelled blocks:
       marquee holdings US, marquee holdings non-US, other stocks US, other stocks non-US.
    """
    ws = wb.create_sheet(SHEET_NAMES[gics])
    ws.sheet_view.showGridLines = False
    title = (f'{gics} - screened names ({universe_label})' if gics != ALL_LABEL else
             f'All sectors combined - screened names ({universe_label})')
    ws['A1'] = f"{title}   |   updated as of {stamp['data']}"
    ws['A1'].font = Font(name='Arial', size=13, bold=True, color=NAVY)
    ws['A2'] = 'Filters: ' + (rule_text or 'no filters applied')
    ws['A2'].font = Font(name='Arial', size=9, italic=True, color='595959')
    blk = 'TUNED_OVERRIDES' if THRESHOLD_MODE == 'tuned' else 'SECTOR_OVERRIDES'
    ws['A3'] = (f"{len(passed)} names cleared these filters (THRESHOLD_MODE = '{THRESHOLD_MODE}'). "
                f"To change a threshold, edit {blk}['{gics}'] in marquee_screener.py and run again. "
                f"Switch THRESHOLD_MODE between 'base' (same for all) and 'tuned' (per sector) to compare.")
    ws['A3'].font = Font(name='Arial', size=9, italic=True, color='9C6500')

    groups = [('Identification', NAVY, [
        ('Ticker', 'Symbol', TXT), ('Company', 'Description', TXT),
        ('TradingView Sector', 'Sector', TXT), ('Industry', 'Industry', TXT),
        ('Country', 'Country or region of registration', TXT),
        ('Marquee Flag', 'Marquee Flag', TXT),
        ('Investors Holding', 'Investors Holding', BIG),
        ('Marquee Wt %', 'Marquee Wt %', PCT)])] + metric_groups('')
    if gics == ALL_LABEL:      # the combined sheet needs the sector on every row
        groups[0][2].insert(2, ('GICS Sector', 'GICS Sector', TXT))

    if len(passed) == 0:
        ws['A5'] = ('No stock in this sector cleared these filters. Either set '
                    "THRESHOLD_MODE = 'tuned' for the sector-aware thresholds, or open "
                    f"{blk}['{gics}'] and loosen the P/E, ROA or Altman value. "
                    'Sector_Screener.html lets you test values live before committing them.')
        ws['A5'].font = Font(name='Arial', size=11, bold=True, color='C00000')
        ws.column_dimensions['A'].width = 120
        return

    # header only, then the four blocks are written by hand underneath
    ncols, pos = write_grid(ws, passed.iloc[0:0], groups, freeze_col=3, start_row=5)
    flat = [(pos[src], src, fmt) for _, _, cols in groups for (_, src, fmt) in cols]

    is_mq = passed['Marquee Flag'] == 'Marquee holding'
    is_us = passed['Country or region of registration'] == 'United States'
    sections = [
        ('MARQUEE HOLDINGS  -  US LISTED',            passed[is_mq & is_us],   '2E5C8A'),
        ('MARQUEE HOLDINGS  -  NON-US LISTED',        passed[is_mq & ~is_us],  '3F6FA5'),
        ('ALL OTHER STOCKS  -  US LISTED',            passed[~is_mq & is_us],  '7B3F00'),
        ('ALL OTHER STOCKS  -  NON-US LISTED',        passed[~is_mq & ~is_us], '9E6B33'),
    ]

    r = 7
    data_rows = []
    for title, block, colour in sections:
        block = block.sort_values('Market capitalization', ascending=False)
        for j in range(1, ncols + 1):
            c = ws.cell(row=r, column=j)
            c.fill = PatternFill('solid', fgColor=colour)
            c.border = BORDER
        head = ws.cell(row=r, column=1, value=f'{title}   ({len(block)} names)')
        head.font = Font(name='Arial', size=10, bold=True, color='FFFFFF')
        head.alignment = LEFT
        r += 1
        if len(block) == 0:
            c = ws.cell(row=r, column=1, value='none cleared the filters in this group')
            c.font = Font(name='Arial', size=9, italic=True, color='808080')
            for j in range(1, ncols + 1):
                ws.cell(row=r, column=j).border = BORDER
            r += 1
            continue
        for i, (_, row) in enumerate(block.iterrows()):
            for cidx, src, fmt in flat:
                c = ws.cell(row=r, column=cidx)
                if not src.startswith('__'):
                    v = row.get(src, None)
                    if isinstance(v, float) and pd.isna(v):
                        v = None
                    if isinstance(v, str) and v.strip() == '':
                        v = None
                    c.value = v
                c.font, c.number_format, c.border = BF, fmt, BORDER
                c.alignment = LEFT if fmt == TXT else RIGHT
                if i % 2:
                    c.fill = STRIPE
            data_rows.append(r)
            r += 1

    ws.auto_filter.ref = f"A6:{get_column_letter(ncols)}{r - 1}"
    fill_cagr(ws, pos, 'Performance %, 5 years', data_rows)




# ------------------------------------------------------------------ deep app feed -----
# The deep analysis app needs the TradingView numbers, and it cannot get them from the
# HTML screener. So the build publishes two small JSON files next to index.html:
#
#   stocks.json        keyed by ISIN, one entry per name worth analysing. Carries every
#                      TradingView field plus the Dataroma ownership, so the deep page
#                      shows TradingView as its primary source and reaches for Yahoo only
#                      where TradingView has nothing.
#   sector_stats.json  per GICS sector quartiles for the main ratios. This is what turns
#                      "ROE 18.2%" into "ROE 18.2%, sector median 12.8%, 72nd percentile",
#                      and the data was already sitting in the export.
#
# Only screened names and marquee holdings go in, not all 30,000 rows. That keeps the file
# near a megabyte instead of ten, which matters because the app fetches it on load.

DEEP_FEED_FIELDS = [
    ('Symbol', 'ticker'), ('Description', 'name'), ('ISIN', 'isin'),
    ('Sector', 'tv_sector'), ('GICS Sector', 'gics'), ('Industry', 'industry'),
    ('Country or region of registration', 'country'),
    ('Market capitalization', 'mcap'), ('Market capitalization - Currency', 'mcap_ccy'),
    ('Price', 'price'), ('Price - Currency', 'ccy'),
    ('Price to earnings ratio', 'pe'), ('Price to earnings ratio forward', 'fpe'),
    ('Return on invested capital %, Trailing 12 months', 'roic'),
    ('Return on capital employed %, Trailing 12 months', 'roce'),
    ('Return on assets %, Trailing 12 months', 'roa'),
    ('Return on equity %, Trailing 12 months', 'roe'),
    ('Gross margin %, Trailing 12 months', 'gross_margin'),
    ('Operating margin %, Trailing 12 months', 'op_margin'),
    ('Net margin %, Trailing 12 months', 'net_margin'),
    ('Free cash flow margin %, Trailing 12 months', 'fcf_margin'),
    ('Free cash flow, Trailing 12 months', 'fcf'), ('Free cash flow, Trailing 12 months - Currency', 'fcf_ccy'),
    ('Current ratio, Quarterly', 'current_ratio'),
    ('Debt to equity ratio, Quarterly', 'debt_equity'),
    ('Interest coverage, Trailing 12 months', 'interest_cover'),
    ('EBITDA interest coverage, Trailing 12 months', 'ebitda_cover'),
    ('Altman Z-score, Trailing 12 months', 'altman_z'),
    ('Piotroski F-score, Trailing 12 months', 'piotroski_f'),
    ('Performance %, Year to date', 'perf_ytd'),
    ('Performance %, 6 months', 'perf_6m'),
    ('Performance %, 1 year', 'perf_1y'),
    ('Performance %, 5 years', 'perf_5y'),
    ('Exponential moving average, 50, 1 day', 'ema50'),
    ('Exponential moving average, 200, 1 day', 'ema200'),
    ('Relative strength index, 14, 1 day', 'rsi14'),
]

# fields that are labels, never numbers
DEEP_FEED_TEXT = {'ticker', 'name', 'isin', 'tv_sector', 'gics', 'industry', 'country',
                  'mcap_ccy', 'ccy', 'fcf_ccy'}

# the period each TradingView field actually covers, so the deep page can label it and
# nobody compares a trailing-twelve-month return against an annual margin by accident
DEEP_FEED_PERIODS = {
    'pe': 'TTM', 'fpe': 'Forward', 'roic': 'TTM', 'roce': 'TTM', 'roa': 'TTM', 'roe': 'TTM',
    'gross_margin': 'TTM', 'op_margin': 'TTM', 'net_margin': 'TTM',
    'fcf_margin': 'TTM', 'fcf': 'TTM',
    'current_ratio': 'Quarterly', 'debt_equity': 'Quarterly',
    'interest_cover': 'TTM', 'ebitda_cover': 'TTM',
    'altman_z': 'TTM', 'piotroski_f': 'TTM',
    'mcap': 'Current', 'price': 'Current', 'ema50': 'Current', 'ema200': 'Current',
    'rsi14': 'Current',
}

# ratios worth a peer distribution; absolute figures are excluded because a median market
# cap or free cash flow across a sector says nothing useful
PEER_METRICS = ['pe', 'fpe', 'roic', 'roce', 'roa', 'roe', 'gross_margin', 'op_margin',
                'net_margin', 'fcf_margin', 'current_ratio', 'debt_equity',
                'interest_cover', 'altman_z', 'piotroski_f']


def _clean_number(value):
    if value is None:
        return None
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return None
        try:
            value = float(value)
        except ValueError:
            return value
    if isinstance(value, float):
        if pd.isna(value):
            return None
        return round(value, 4)
    return value


DEEP_CARDS = {}               # 'TICKER|CCY' -> compact deep scorecard, for the web page export

def load_deep_scoring():
    """The deep app's scoring module, or None."""
    import importlib.util, types, urllib.request
    local = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'deep_scoring.py')
    try:
        if os.path.exists(local):
            spec = importlib.util.spec_from_file_location('deep_scoring', local)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
        with urllib.request.urlopen(DEEP_SCORING_URL, timeout=20) as fh:
            code = fh.read().decode('utf-8')
        mod = types.ModuleType('deep_scoring')
        exec(compile(code, DEEP_SCORING_URL, 'exec'), mod.__dict__)
        return mod
    except Exception as e:
        print(f'      [WARN] deep scorecard not loaded ({e}); the export will leave it out')
        return None


def _snapshot_momentum(tv):
    """A trend reading from the snapshot averages, in the shape the deep scorecard expects.
    The deep page reads the full price history instead, so it can also date the crossover."""
    px, a, b, rsi = (tv.get(k) for k in ('price', 'ema50', 'ema200', 'rsi14'))
    if not all(isinstance(x, (int, float)) for x in (px, a, b)) or a <= 0:
        return None
    stretch = (px / a - 1) * 100
    if px > a * 1.25 or (isinstance(rsi, (int, float)) and rsi > 75 and px > a):
        return {'score': 50, 'text': f'The price is {stretch:.0f}% above its 50 day average: strong, but stretched.'}
    # same stages as the screener page, so the export never shows two different readings
    if px > a and a > b * 1.10:
        return {'score': 70, 'text': 'Above both averages with the 50 day well over the 200 day: an established uptrend.'}
    if px < a < b:
        return {'score': 20, 'text': 'Below both averages with the 50 day under the 200 day: a downtrend.'}
    return {'score': 45, 'text': 'Between its averages: no clear trend either way.'}


def build_deep_cards(stocks, stats):
    """Score every name in the deep feed with the deep app's scorecard, from snapshot data."""
    DEEP_CARDS.clear()
    mod = load_deep_scoring()
    if mod is None:
        return
    short = {'Quality': 'q', 'Financial strength': 'f', 'Valuation': 'v', 'Momentum': 'm', 'Signals': 'g'}
    for tv in stocks.values():
        try:
            mo = _snapshot_momentum(tv)
            sc = mod.score(tv, stats.get(tv.get('gics')) or {}, {}, {'momentum': mo} if mo else {},
                           [], {}, [])
        except Exception:
            continue
        card = {'s': sc.get('overall'), 'o': sc.get('opinion'), 'c': sc.get('coverage'),
                'l': sc.get('lifts') or [], 'd': sc.get('drags') or [], 'x': sc.get('flags') or []}
        for name, p in (sc.get('pillars') or {}).items():
            if p and name in short:
                card[short[name]] = p.get('score')
        DEEP_CARDS[f"{tv.get('ticker')}|{(tv.get('ccy') or '').upper()}"] = card
    print(f'      deep scorecard -> {len(DEEP_CARDS)} names scored for the export')


def write_deep_feed(all_rows, screened_symbols, out_dir, publish_dir, stamp):
    """Write stocks.json and sector_stats.json for the deep analysis app.

    Both are written on every run. stocks.json is keyed by ISIN where one is known, from
    the export or from the ISIN memory, and by a TV:TICKER:CCY key otherwise, so a missing
    ISIN column can no longer leave the deep app on stale numbers.
    """
    folders = [d for d in (out_dir, publish_dir) if d]

    # peer distributions, computed on the FULL universe rather than the screened subset,
    # because a median taken from names that already passed a quality screen is not a
    # sector median - it is a median of the winners
    stats = {}
    for gics in SECTOR_MAP:
        sub = all_rows[all_rows['GICS Sector'] == gics]
        if len(sub) < 20:
            continue
        block = {'count': int(len(sub))}
        for src, key in DEEP_FEED_FIELDS:
            if key not in PEER_METRICS or src in MISSING_COLS:
                continue
            series = pd.to_numeric(sub[src], errors='coerce').dropna()
            if len(series) < 20:
                continue
            block[key] = {'n': int(len(series)),
                          'p25': round(float(series.quantile(0.25)), 4),
                          'median': round(float(series.median()), 4),
                          'p75': round(float(series.quantile(0.75)), 4)}
        stats[gics] = block
    stats_feed = {'as_of': stamp['data'], 'basis': 'full TradingView universe, not the '
                                                   'screened subset', 'sectors': stats}
    for folder in folders:
        os.makedirs(folder, exist_ok=True)
        with open(os.path.join(folder, 'sector_stats.json'), 'w') as fh:
            json.dump(stats_feed, fh, indent=1)

    rows = all_rows[
        all_rows['Symbol'].isin(screened_symbols)
        | (all_rows['Marquee Flag'] == 'Marquee holding')
    ].copy()
    if len(rows) == 0:
        old_as_of = ''
        for folder in folders:
            try:
                with open(os.path.join(folder, 'stocks.json')) as fh:
                    old_as_of = json.load(fh).get('as_of', '')
                break
            except Exception:
                continue
        print(f'      sector_stats.json refreshed ({len(stats)} sectors)')
        print('      [WARN] no screened or marquee rows, stocks.json left as it was'
              + (f' (as of {old_as_of})' if old_as_of else ''))
        return

    stocks = {}
    for _, r in rows.iterrows():
        entry = {}
        for src, key in DEEP_FEED_FIELDS:
            v = None if src in MISSING_COLS else r.get(src)
            if key in DEEP_FEED_TEXT:
                # kept as text, so ticker 005930 does not turn into the number 5930.0
                entry[key] = None if v is None or (isinstance(v, float) and pd.isna(v)) \
                    else (str(v).strip() or None)
            else:
                entry[key] = _clean_number(v)
        entry['marquee'] = bool(r.get('Marquee Flag') == 'Marquee holding')
        entry['marquee_weight_pct'] = _clean_number(r.get('Marquee Wt %'))
        entry['marquee_investors'] = _clean_number(r.get('Investors Holding'))
        entry['screened'] = bool(r['Symbol'] in screened_symbols)
        # Keyed by ISIN where one is known. A name with no ISIN anywhere still goes in,
        # under a TV: key built from its ticker and currency, so the deep app always has
        # this run's TradingView numbers for it. The app only treats ISIN-shaped keys as
        # ISINs. Last writer wins on a duplicate key, which only happens for dual listings.
        isin = r.get('ISIN')
        if isinstance(isin, str) and isin:
            key = isin
        else:
            key = f"TV:{str(r['Symbol']).strip().upper()}:{str(r.get('Price - Currency') or '').upper()}"
        stocks[key] = entry

    build_deep_cards(stocks, stats)
    feed = {'as_of': stamp['data'], 'marquee_as_of': stamp['marquee'],
            'built': stamp['built'], 'source': 'TradingView screener export',
            'periods': DEEP_FEED_PERIODS, 'count': len(stocks), 'stocks': stocks}
    for folder in folders:
        with open(os.path.join(folder, 'stocks.json'), 'w') as fh:
            json.dump(feed, fh, separators=(',', ':'))
    size = os.path.getsize(os.path.join(out_dir, 'stocks.json')) / 1024
    n_isin = sum(1 for k in stocks if not k.startswith('TV:'))
    print(f'      deep feed -> stocks.json ({len(stocks)} names, {n_isin} by ISIN, '
          f'{len(stocks) - n_isin} by ticker, {size:,.0f} KB) '
          f'and sector_stats.json ({len(stats)} sectors)')


# ------------------------------------------------------------------ html -------------
HTML_COLS = [('Symbol', 'sym'), ('Description', 'name'), ('GICS Sector', 'gics'),
             ('ISIN', 'isin'),
             ('Country or region of registration', 'ctry'), ('Marquee Flag', 'mq'),
             ('Marquee Wt %', 'wt'), ('Market capitalization', 'mcap'),
             ('Price to earnings ratio', 'pe'), ('Price to earnings ratio forward', 'fpe'),
             ('Return on invested capital %, Trailing 12 months', 'roic'),
             ('Return on capital employed %, Trailing 12 months', 'roce'),
             ('Return on assets %, Trailing 12 months', 'roa'),
             ('Return on equity %, Trailing 12 months', 'roe'),
             ('Gross margin %, Trailing 12 months', 'gm'), ('Operating margin %, Trailing 12 months', 'om'),
             ('Net margin %, Trailing 12 months', 'nm'), ('Free cash flow margin %, Trailing 12 months', 'fcfm'),
             ('Current ratio, Quarterly', 'cr'), ('Debt to equity ratio, Quarterly', 'de'),
             ('Interest coverage, Trailing 12 months', 'ic'),
             ('Altman Z-score, Trailing 12 months', 'az'),
             ('Piotroski F-score, Trailing 12 months', 'pf'),
             ('Performance %, Year to date', 'ytd'),
             ('Performance %, 1 year', 'y1'), ('Performance %, 5 years', 'y5'),
             ('Price', 'px'), ('Price - Currency', 'ccy'),
             ('Industry', 'ind'), ('Investors Holding', 'inv'),
             ('Exponential moving average, 50, 1 day', 'e50'),
             ('Exponential moving average, 200, 1 day', 'e200'),
             ('Relative strength index, 14, 1 day', 'rsi'),
             ('Performance %, 6 months', 'm6'),
             # optional: add these to the TradingView export to switch on more momentum signals
             ('Performance %, 1 month', 'm1'), ('Performance %, 3 months', 'm3'),
             ('52 week high', 'h52'), ('Relative volume', 'rv')]

# HTML filter key -> (data field, direction, label)
HTML_FILTERS = [(k, HTML_KEY, FILTER_SPECS[k][1], FILTER_SPECS[k][2]) for k, HTML_KEY in [
    ('pe_max', 'pe'), ('roe_min', 'roe'), ('roic_min', 'roic'), ('roce_min', 'roce'),
    ('roa_min', 'roa'), ('gross_margin_min', 'gm'), ('op_margin_min', 'om'),
    ('net_margin_min', 'nm'), ('fcf_margin_min', 'fcfm'), ('altman_min', 'az'),
    ('piotroski_min', 'pf'), ('current_ratio_min', 'cr'), ('de_max', 'de'),
    ('int_cov_min', 'ic'), ('mcap_min', 'mcap')]]


def build_html(a2, out_path, stamp, split=False):
    """Write the screener. split=True writes a light index.html plus a separate data.json,
    which is what a website wants: the page paints immediately and the data streams in."""
    # industry names repeat thousands of times, so each row carries an index into one list
    industries = sorted({str(x).strip() for x in a2['Industry'].dropna() if str(x).strip()})
    ind_pos = {name: i for i, name in enumerate(industries)}
    # optional momentum columns that this export does not carry are left out of the page
    cols = [(src, k) for src, k in HTML_COLS
            if not (src in OPTIONAL_COLS and pd.to_numeric(a2.get(src), errors='coerce').notna().sum() == 0)]
    rows = []
    for _, r in a2.iterrows():
        rec = []
        for src, _ in cols:
            v = r.get(src, None)
            if src == 'GICS Sector':
                v = None if pd.isna(v) else list(SECTOR_MAP).index(v)
            elif src == 'Industry':
                v = None if pd.isna(v) else ind_pos.get(str(v).strip())
            elif src == 'Marquee Flag':
                v = 1 if v == 'Marquee holding' else 0
            elif isinstance(v, float):
                # four decimals, not two: rounding at the second decimal shifts borderline
                # rows across a threshold and makes the HTML count drift from the workbook
                v = None if pd.isna(v) else round(v, 4)
            elif isinstance(v, str):
                v = v.strip() or None
            elif pd.isna(v):
                v = None
            rec.append(v)
        rows.append(rec)
    payload = {
        'fields': [k for _, k in cols],
        'rows': rows,
        'sectors': list(SECTOR_MAP.keys()),
        'industries': industries,
        'dropdown': [ALL_LABEL] + list(SECTOR_MAP.keys()),
        'allLabel': ALL_LABEL,
        'defaults': DEFAULT_FILTERS,
        'base': SECTOR_OVERRIDES,
        'tuned': COMPUTED_TUNED,
        'dist': DISTRIBUTIONS,
        'tunedPct': 50,
        'fixedBy': fixed_rules(),
        'marketDist': MARKET_DIST,
        'deep': DEEP_CARDS,
        'ruleSets': SECTOR_RULES,
        'ruleShift': {f'{g}|{k}': v for (g, k), v in RULE_TARGET_SHIFT.items()},
        'notUsedWhy': not_used_why(),
        'fixed': {'mcap_min': MCAP_FLOOR},
        'positiveOnly': sorted(POSITIVE_ONLY),
        'mode': THRESHOLD_MODE,
        'perSector': ALL_SECTORS_USES_SECTOR_RULES,
        'filters': [{'key': k, 'field': f, 'dir': d, 'label': l} for k, f, d, l in HTML_FILTERS],
        'blankFails': BLANK_FAILS,
        'minCoverage': MIN_DATA_COVERAGE,
        'stamp': stamp,
        'deepUrl': DEEP_APP_URL.rstrip('/'),
        'xlsxUrl': (f'https://github.com/{GITHUB_REPO}/raw/{GITHUB_BRANCH}/{GITHUB_XLSX}'
                    if (GITHUB_UPLOAD and GITHUB_XLSX) else ''),
    }
    blob = json.dumps(payload, separators=(',', ':'))
    if split:
        data_path = os.path.join(os.path.dirname(out_path) or '.', 'data.json')
        with open(data_path, 'w') as fh:
            fh.write(blob)
        bootstrap = ("""document.getElementById('msg').innerHTML =
  '<div class="note" style="padding:10px 12px">Loading the stock data, one moment...</div>';
fetch('data.json' + (location.search||''), {cache:'no-cache'})
  .then(r=>{ if(!r.ok) throw new Error('HTTP '+r.status); return r.json(); })
  .then(d=>{ document.getElementById('msg').innerHTML=''; boot(d); })
  .catch(e=>{ document.getElementById('msg').innerHTML =
    '<div class="warn">Could not load data.json (' + e.message + '). It must sit in the same '
    + 'folder as this page. Opening this file straight from your desktop will not work - '
    + 'browsers block local file reads. Use the single-file version for that.</div>'; });""")
        html = HTML_TEMPLATE.replace('/*__BOOT__*/', bootstrap)
        with open(out_path, 'w') as fh:
            fh.write(html)
        return os.path.getsize(out_path), os.path.getsize(data_path)
    html = HTML_TEMPLATE.replace('/*__BOOT__*/', 'boot(' + blob + ');')
    with open(out_path, 'w') as fh:
        fh.write(html)
    return os.path.getsize(out_path), 0


HTML_TEMPLATE = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Global equity screener | MICS International</title>
<style>
:root{
 --navy:#1f3864;--navy-deep:#172b4d;--navy-2:#2e4a7d;--gold:#bf8f00;
 --ink:#1a1a1a;--line:#d8dee9;--line-soft:#e9edf3;--canvas:#f4f6fa;--surface:#fff;
 --pale:#eef2f8;--up:#1e6b3b;--down:#a12a2a;--focus:#bf8f00;
 --r:8px;--sans:-apple-system,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;
}
*{box-sizing:border-box}
html,body{margin:0}
body{font:14px/1.45 var(--sans);color:var(--ink);background:var(--canvas);-webkit-font-smoothing:antialiased}
button,input,select{font:inherit;color:inherit}
:focus-visible{outline:2px solid var(--focus);outline-offset:2px}
.num{font-variant-numeric:tabular-nums}

/* ---------- header */
.top{background:var(--navy-deep);color:#fff;border-bottom:3px solid var(--gold)}
.top .in{max-width:1680px;margin:0 auto;padding:16px 24px;display:flex;align-items:flex-end;
 justify-content:space-between;gap:24px;flex-wrap:wrap}
.brand{display:flex;flex-direction:column;gap:2px}
.brand .org{font-size:12.5px;font-weight:600;color:#c9d3e3;letter-spacing:.02em}
.brand h1{margin:0;font-size:22px;font-weight:650;letter-spacing:-.01em}
.dates{display:flex;gap:22px;font-size:12.5px;color:#c9d3e3}
.dates div{display:flex;flex-direction:column;gap:1px}
.dates b{color:#fff;font-weight:600;font-size:13.5px}
.toplink{align-self:center;color:#fff;font-weight:600;font-size:13px;text-decoration:none;border:1px solid #5d7299;
 border-radius:6px;padding:7px 12px;margin-right:6px}
.toplink:hover{background:#24406e}
.toplink[hidden]{display:none}

/* ---------- page grid */
.page{max-width:1680px;margin:0 auto;padding:20px 24px 40px;display:grid;
 grid-template-columns:340px minmax(0,1fr);gap:20px;align-items:start}
@media (max-width:980px){.page{grid-template-columns:1fr;padding:14px}}

/* ---------- sidebar */
.side{position:sticky;top:16px;max-height:calc(100vh - 32px);overflow:auto;
 background:var(--surface);border:1px solid var(--line);border-radius:var(--r)}
@media (max-width:980px){.side{position:static;max-height:none}}
.blk{border-bottom:1px solid var(--line-soft)}
.blk>summary{list-style:none;cursor:pointer;display:flex;align-items:center;justify-content:space-between;
 gap:10px;padding:14px 16px;font-weight:650;color:var(--navy);font-size:14px;user-select:none}
.blk>summary::-webkit-details-marker{display:none}
.blk>summary .sv{margin-left:auto;font-weight:500;font-size:12.5px;color:var(--navy-2);text-align:right;
 overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:180px}
.blk>summary::after{content:"";width:7px;height:7px;border-right:2px solid var(--navy);
 border-bottom:2px solid var(--navy);transform:rotate(45deg);margin:0 2px 3px 4px;flex:0 0 auto;transition:transform .15s}
.blk[open]>summary::after{transform:rotate(-135deg);margin-bottom:-3px}
.bd{padding:0 16px 16px}
.row-acts{display:flex;gap:14px;margin:-4px 0 8px}
.link{background:none;border:0;padding:0;color:var(--navy);font-weight:600;font-size:12.5px;cursor:pointer;
 text-decoration:underline;text-underline-offset:3px;text-decoration-color:#9fb0cc}
.link:hover{text-decoration-color:var(--navy)}

.tick{display:flex;align-items:center;gap:10px;padding:6px 8px;margin:0 -8px;border-radius:6px;cursor:pointer;font-size:13.5px}
.tick:hover{background:var(--pale)}
.box{appearance:none;-webkit-appearance:none;width:17px;height:17px;border:1.5px solid #8ea0bf;border-radius:4px;margin:0;
 flex:0 0 auto;display:grid;place-content:center;background:#fff;cursor:pointer}
.box:checked{background:var(--navy);border-color:var(--navy)}
.box:checked::after{content:"";width:9px;height:5px;border-left:2px solid #fff;border-bottom:2px solid #fff;
 transform:rotate(-45deg) translate(1px,-1px)}
.box:disabled{opacity:.4;cursor:not-allowed}
.tick .t{flex:1 1 auto}
.tick .c{font-size:12px;color:var(--navy-2)}
.tick.dis{cursor:not-allowed}
.tick.dis .t{opacity:.55}
.hint{font-size:12.5px;color:var(--navy-2);line-height:1.5;margin:6px 0 0}
.sep{height:1px;background:var(--line-soft);margin:8px 0}
.search{width:100%;padding:8px 10px;border:1px solid var(--line);border-radius:6px;background:#fff;margin-top:8px}
.search:focus{border-color:var(--navy);outline:none;box-shadow:0 0 0 3px #dfe6f2}
.lab{font-size:12.5px;font-weight:650;color:var(--navy);margin:14px 0 6px}

.seg{display:grid;grid-template-columns:1fr 1fr;border:1px solid var(--line);border-radius:7px;overflow:hidden}
.seg label{position:relative;cursor:pointer}
.seg input{position:absolute;opacity:0;pointer-events:none}
.seg span{display:block;text-align:center;padding:7px 6px;font-size:13px;font-weight:600;color:var(--navy)}
.seg label+label span{border-left:1px solid var(--line)}
.seg input:checked+span{background:var(--navy);color:#fff}
.seg input:focus-visible+span{outline:2px solid var(--focus);outline-offset:-2px}

.slider .top-l{display:flex;justify-content:space-between;align-items:baseline;gap:10px}
.slider output{font-weight:650;color:var(--navy);font-size:13px;white-space:nowrap}
.slider input[type=range]{width:100%;accent-color:var(--navy);margin:8px 0 2px}
.slider .ends{display:flex;justify-content:space-between;font-size:11.5px;color:var(--navy-2)}

/* criteria */
.cg h4{margin:16px 0 4px;font-size:12px;font-weight:650;color:var(--gold)}
.cg h4:first-child{margin-top:2px}
.crit{display:grid;grid-template-columns:17px minmax(0,1fr) auto;gap:2px 10px;align-items:center;
 padding:7px 8px;margin:0 -8px;border-radius:6px}
.crit:hover{background:var(--pale)}
.crit .nm{font-size:13.5px;cursor:pointer}
.crit .val{display:flex;align-items:center;gap:5px;justify-self:end}
.crit .op{color:var(--gold);font-weight:700;width:12px;text-align:center}
.crit input[type=number]{width:72px;padding:5px 7px;border:1px solid var(--line);border-radius:5px;text-align:right;background:#fff}
.crit input[type=number]:focus{border-color:var(--navy);outline:none;box-shadow:0 0 0 3px #dfe6f2}
.crit .u{font-size:12px;color:var(--navy-2);min-width:26px}
.crit .rng{font-size:12.5px;color:var(--navy);text-align:right;white-space:nowrap}
.crit .eff{grid-column:2/4;font-size:11.5px;color:var(--navy-2);display:flex;justify-content:space-between;gap:8px}
.crit .eff .cut{color:var(--down);font-weight:600}
.crit.off .nm,.crit.off .val,.crit.off .eff{opacity:.45}

.side-foot{display:flex;justify-content:space-between;gap:10px;padding:12px 16px;border-top:1px solid var(--line);background:#fafbfd}

/* ---------- main */
.main{min-width:0;display:flex;flex-direction:column;gap:14px}
.card{background:var(--surface);border:1px solid var(--line);border-radius:var(--r)}
.sum{padding:18px 20px;display:flex;flex-wrap:wrap;gap:16px 34px;align-items:flex-end;justify-content:space-between}
.sum .big{display:flex;align-items:baseline;gap:10px}
.sum .big b{font-size:34px;line-height:1;font-weight:700;color:var(--navy);letter-spacing:-.02em}
.sum .big span{font-size:15px;color:var(--navy);font-weight:600}
.sum .rule{font-size:13px;color:var(--navy-2);margin-top:6px;max-width:660px}
.kpis{display:flex;gap:28px;flex-wrap:wrap}
.kpi{display:flex;flex-direction:column;gap:2px}
.kpi b{font-size:18px;font-weight:650;color:var(--ink)}
.kpi span{font-size:12px;color:var(--navy-2)}
.secbar{padding:0 20px 16px;display:flex;flex-wrap:wrap;gap:6px}
.secbar:empty{display:none}
.chip{display:inline-flex;gap:6px;align-items:baseline;padding:4px 10px;border-radius:999px;background:var(--pale);
 font-size:12.5px;color:var(--navy)}
.chip b{font-weight:650}

.bar{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;padding:12px 16px;
 border-bottom:1px solid var(--line)}
.bar .grp{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.bar .lead{font-size:13px;font-weight:650;color:var(--navy);margin-right:2px}
.btn{display:inline-flex;align-items:center;gap:7px;padding:8px 14px;border-radius:6px;border:1px solid var(--navy);
 background:var(--navy);color:#fff;font-weight:600;font-size:13px;cursor:pointer;line-height:1.1;text-decoration:none}
.btn:hover{background:var(--navy-2)}
.btn.ghost{background:#fff;color:var(--navy)}
.btn.ghost:hover{background:var(--pale)}
.btn:disabled{opacity:.5;cursor:not-allowed}
.xstat{font-size:12.5px;color:var(--up);font-weight:600}
.xstat.err{color:var(--down)}
.pg{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--navy)}
.pg select{padding:5px 6px;border:1px solid var(--line);border-radius:5px;background:#fff}

.tw{overflow:auto;max-height:calc(100vh - 300px);min-height:260px}
table{border-collapse:separate;border-spacing:0;width:100%;font-size:13px}
th,td{padding:8px 10px;white-space:nowrap;text-align:right;border-bottom:1px solid var(--line-soft)}
th{position:sticky;top:0;z-index:2;background:#fafbfd;color:var(--navy);font-weight:650;font-size:12px;
 border-bottom:1px solid var(--line);cursor:pointer;user-select:none}
th:hover{background:var(--pale)}
th .ar{display:inline-block;width:10px;color:var(--gold)}
th.l,td.l{text-align:left}
th.co,td.co{position:sticky;left:0;z-index:1;background:#fff;text-align:left;min-width:230px;max-width:290px;
 white-space:normal;border-right:1px solid var(--line-soft)}
th.co{z-index:3;background:#fafbfd}
tbody tr:hover td,tbody tr:hover td.co{background:#fbf6e8}
.co a,.co .tk{color:var(--navy);font-weight:700;text-decoration:none}
.co a:hover{text-decoration:underline;text-underline-offset:3px}
.co .nm{display:block;font-size:12.5px;color:var(--ink);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:265px}
.star{color:var(--gold);margin-left:5px;font-size:12px}
td.neg{color:var(--down)}
td.na{color:#8a96ad}
.empty{padding:56px 24px;text-align:center}
.empty h3{margin:0 0 6px;color:var(--navy);font-size:17px}
.empty p{margin:0 auto;max-width:540px;color:var(--navy-2);font-size:13.5px}
.foot{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;padding:10px 16px;
 border-top:1px solid var(--line);font-size:12.5px;color:var(--navy-2)}
.pager{display:flex;gap:4px;flex-wrap:wrap}
.pager button{min-width:32px;padding:5px 9px;border:1px solid var(--line);border-radius:5px;background:#fff;
 color:var(--navy);font-weight:600;cursor:pointer}
.pager button[aria-current="page"]{background:var(--navy);border-color:var(--navy);color:#fff}
.pager button:disabled{opacity:.4;cursor:default}
.pager .gap{padding:5px 2px;color:var(--navy-2)}
.note{font-size:12px;color:var(--navy-2);padding:0 2px;line-height:1.55;margin:0}
/* rules by sector */
dialog#mx{border:0;border-radius:10px;padding:20px 22px;max-width:min(1200px,94vw);width:100%;
 box-shadow:0 20px 60px rgba(23,43,77,.35)}
dialog#mx::backdrop{background:rgba(23,43,77,.45)}
.mx-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px}
.mx-head h2{margin:0;color:var(--navy);font-size:18px}
.mx-wrap{overflow:auto;max-height:70vh;border:1px solid var(--line);border-radius:8px}
#mx-t th{cursor:default;white-space:normal;min-width:92px;vertical-align:bottom}
#mx-t th:first-child,#mx-t td:first-child{position:sticky;left:0;background:#fafbfd;text-align:left;z-index:3;min-width:200px}
#mx-t td:first-child{background:#fff;font-weight:600;color:var(--navy)}
#mx-t td.off{color:var(--navy-2);font-style:italic;font-size:12px}
#mx-t td.mx-g{background:var(--pale)!important;color:var(--gold);font-weight:650;font-size:12px}
/* momentum chips */
.mo{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px;font-weight:600}
.mo.early{background:#e3f1e8;color:var(--up)}
.mo.established{background:var(--pale);color:var(--navy)}
.mo.extended{background:#fbf0d6;color:#8a6400}
.mo.neutral{background:#f1f3f7;color:var(--navy-2)}
.mo.weak{background:#f7e4e4;color:var(--down)}
.mo.na{background:none;color:#8a96ad;font-weight:500}
.sig-na .t{opacity:.5}
@media (prefers-reduced-motion:reduce){*{transition:none!important}}
</style></head><body>
<header class="top"><div class="in">
 <div class="brand"><span class="org">MICS International</span><h1>Global equity screener</h1></div>
 <div class="dates">
  <a class="toplink" id="idx-link" href="#" target="_blank" rel="noopener" hidden>Index analysis</a>
  <div><span>Company data</span><b id="d-data"></b></div>
  <div><span>Superinvestor holdings</span><b id="d-mq"></b></div>
  <div><span>Page built</span><b id="d-built"></b></div>
 </div>
</div></header>

<div class="page">
<aside class="side" aria-label="Screen settings">

 <details class="blk" open>
  <summary>Sectors <span class="sv" id="sv-sec"></span></summary>
  <div class="bd">
   <div class="row-acts"><button type="button" class="link" id="sec-all">Select all</button>
    <button type="button" class="link" id="sec-none">Clear</button></div>
   <div id="sec-list"></div>
  </div>
 </details>

 <details class="blk" open>
  <summary>Market <span class="sv" id="sv-mkt"></span></summary>
  <div class="bd">
   <label class="tick"><input class="box" type="checkbox" id="m-us" checked><span class="t">Listed in the US</span><span class="c num" id="c-us"></span></label>
   <label class="tick"><input class="box" type="checkbox" id="m-non" checked><span class="t">Listed outside the US</span><span class="c num" id="c-non"></span></label>
   <div class="sep"></div>
   <label class="tick"><input class="box" type="checkbox" id="m-mq"><span class="t">Only companies superinvestors hold</span><span class="c num" id="c-mq"></span></label>
   <input class="search" type="search" id="q" placeholder="Search by company or ticker" aria-label="Search by company or ticker">
  </div>
 </details>

 <details class="blk" open>
  <summary>Rules <span class="sv" id="sv-rules"></span></summary>
  <div class="bd">
   <div class="lab" style="margin-top:0">Threshold set</div>
   <div class="seg" role="radiogroup" aria-label="Threshold set">
    <label><input type="radio" name="preset" value="tuned"><span>Sector-tuned</span></label>
    <label><input type="radio" name="preset" value="base"><span>Same for all</span></label>
   </div>
   <p class="hint" id="preset-hint"></p>
   <div class="sep"></div>
   <label class="tick" id="per-wrap"><input class="box" type="checkbox" id="per"><span class="t">Judge each company on its own sector's rules</span></label>
   <label class="tick"><input class="box" type="checkbox" id="strict"><span class="t">Exclude companies with missing figures</span></label>
   <div class="lab">Strictness</div>
   <div class="slider">
    <div class="top-l"><span class="hint" style="margin:0" id="s-sub">Moves every rule at once</span><output id="s-out" for="s-rng">As set</output></div>
    <input type="range" id="s-rng" min="-25" max="25" step="5" value="0" list="s-ticks" aria-label="Strictness">
    <datalist id="s-ticks"><option value="-25"></option><option value="0"></option><option value="25"></option></datalist>
    <div class="ends"><span>Stricter</span><span id="s-mid">Median</span><span>Looser</span></div>
   </div>
  </div>
 </details>

 <details class="blk" open>
  <summary>Momentum <span class="sv" id="sv-mom"></span></summary>
  <div class="bd">
   <label class="tick"><input class="box" type="checkbox" id="mom-on"><span class="t">Filter by momentum stage</span></label>
   <div id="mom-stages"></div>
   <div class="lab">How early</div>
   <div class="slider">
    <div class="top-l"><span class="hint" style="margin:0" id="e-sub">Sets how far a stock may have run</span><output id="e-out" for="e-rng">Balanced</output></div>
    <input type="range" id="e-rng" min="1" max="5" step="1" value="3" list="e-ticks" aria-label="How early">
    <datalist id="e-ticks"><option value="1"></option><option value="3"></option><option value="5"></option></datalist>
    <div class="ends"><span>Very early</span><span>Already moving</span></div>
   </div>
   <div class="lab">An early stage needs</div>
   <div id="mom-sigs"></div>
   <p class="hint" id="mom-hint"></p>
  </div>
 </details>

 <details class="blk" open>
  <summary>Criteria <span class="sv" id="sv-crit"></span></summary>
  <div class="bd">
   <div class="row-acts"><button type="button" class="link" id="c-all">Tick all</button>
    <button type="button" class="link" id="c-none">Untick all</button>
    <button type="button" class="link" id="c-matrix" style="margin-left:auto">Rules by sector</button></div>
   <div id="crit"></div>
   <p class="hint">The red figure under each rule is how many companies that rule removes on its own, out of those in the ticked sectors and markets.</p>
  </div>
 </details>

 <div class="side-foot">
  <button type="button" class="link" id="reset">Reset everything</button>
  <button type="button" class="link" id="copy">Copy rules for the script</button>
 </div>
</aside>

<main class="main">
 <section class="card" aria-live="polite">
  <div class="sum">
   <div>
    <div class="big"><b class="num" id="n">0</b><span id="n-lab">companies pass</span></div>
    <div class="rule" id="rule-line"></div>
   </div>
   <div class="kpis">
    <div class="kpi"><b class="num" id="k-us">0</b><span>US listed</span></div>
    <div class="kpi"><b class="num" id="k-non">0</b><span>Outside the US</span></div>
    <div class="kpi"><b class="num" id="k-mq">0</b><span>Superinvestor held</span></div>
    <div class="kpi"><b class="num" id="k-cap">-</b><span>Median market cap</span></div>
   </div>
  </div>
  <div class="secbar" id="secbar"></div>
 </section>

 <section class="card">
  <div class="bar">
   <div class="grp">
    <span class="lead">Export</span>
    <button class="btn" id="xl" type="button">Excel workbook</button>
    <button class="btn" id="pp" type="button">PowerPoint</button>
    <button class="btn ghost" id="csv" type="button">CSV</button>
    <a id="xlsx" href="#" style="display:none" class="btn ghost">Full screener workbook</a>
    <span class="xstat" id="xstat" role="status" aria-live="polite"></span>
   </div>
   <div class="pg"><label for="per-page">Rows per page</label>
    <select id="per-page"><option>25</option><option selected>50</option><option>100</option><option>250</option></select></div>
  </div>
  <div class="tw" id="tw"><table><thead><tr id="thead"></tr></thead><tbody id="tbody"></tbody></table>
   <div class="empty" id="empty" hidden><h3 id="empty-h"></h3><p id="empty-p"></p></div></div>
  <div class="foot"><span id="showing"></span><div class="pager" id="pager"></div></div>
 </section>
 <p class="note" id="hint">Click a column heading to sort. Exports always contain every company that passes, not only this page.</p>
</main>
</div>
<dialog id="mx" aria-labelledby="mx-h">
 <div class="mx-head"><h2 id="mx-h">Rules by sector</h2>
  <div class="seg" id="mx-seg" role="radiogroup" aria-label="Threshold set shown" style="min-width:300px;margin-left:auto;margin-right:12px">
   <label><input type="radio" name="mxpreset" value="tuned"><span>Sector-tuned</span></label>
   <label><input type="radio" name="mxpreset" value="base"><span>Same for all</span></label>
  </div>
  <button type="button" class="btn ghost" id="mx-close">Close</button></div>
 <p class="hint" id="mx-sub" style="margin:0 0 12px"></p>
 <div class="mx-wrap"><table id="mx-t"></table></div>
</dialog>
<script>
let DATA, F, ROWS, SEC; const IX = {};
const OTHER = -1;                       // rows whose TradingView sector maps to no GICS sector
const $ = id => document.getElementById(id);
const STORE = 'mics-screener-v4';

// How each criterion is shown. 'short' is used in exports, 'name' in the sidebar.
const CRIT = {
  pe_max:            {g:'Valuation',      name:'P/E, trailing',              short:'P/E',            u:'x'},
  roe_min:           {g:'Returns',        name:'Return on equity',           short:'ROE',            u:'%'},
  roa_min:           {g:'Returns',        name:'Return on assets',           short:'ROA',            u:'%'},
  roic_min:          {g:'Returns',        name:'Return on invested capital', short:'ROIC',           u:'%'},
  roce_min:          {g:'Returns',        name:'Return on capital employed', short:'ROCE',           u:'%'},
  gross_margin_min:  {g:'Margins',        name:'Gross margin',               short:'Gross margin',   u:'%'},
  op_margin_min:     {g:'Margins',        name:'Operating margin',           short:'Op. margin',     u:'%'},
  net_margin_min:    {g:'Margins',        name:'Net margin',                 short:'Net margin',     u:'%'},
  fcf_margin_min:    {g:'Margins',        name:'Free cash flow margin',      short:'FCF margin',     u:'%'},
  current_ratio_min: {g:'Balance sheet',  name:'Current ratio',              short:'Current ratio',  u:'x'},
  de_max:            {g:'Balance sheet',  name:'Debt to equity',             short:'Debt / equity',  u:'x'},
  int_cov_min:       {g:'Balance sheet',  name:'Interest cover',             short:'Interest cover', u:'x'},
  altman_min:        {g:'Quality scores', name:'Altman Z-score',             short:'Altman Z',       u:''},
  piotroski_min:     {g:'Quality scores', name:'Piotroski F-score',          short:'Piotroski F',    u:'of 9'},
  mcap_min:          {g:'Size',           name:'Market cap',                 short:'Market cap',     u:'$bn', scale:1e9},
};
const GROUPS = ['Valuation','Returns','Margins','Balance sheet','Quality scores','Size'];

const S = { sel:new Set(), us:true, non:true, mq:false, q:'', preset:'tuned', per:true, strict:false,
            loosen:0, on:{}, val:{}, page:1, perPage:50, sortKey:'mcap', sortDir:-1,
            momOn:false, early:3, stages:new Set(['early']), sigs:new Set(['trend','fresh','rsi','accel','room','vol']) };
let cur = [], lastRuleKey = null;

function boot(d){
  DATA=d; F=DATA.fields; F.forEach((f,i)=>IX[f]=i); ROWS=DATA.rows; SEC=DATA.sectors; init();
}

// ---------------------------------------------------------------- formatting
const nf0 = new Intl.NumberFormat('en-US',{maximumFractionDigits:0});
const nf1 = new Intl.NumberFormat('en-US',{minimumFractionDigits:1,maximumFractionDigits:1});
const nf2 = new Intl.NumberFormat('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
const isNum = v => v!==null && v!==undefined && !Number.isNaN(v);
function bn(v){ if(!isNum(v)) return '-'; const b=v/1e9;
  return b>=1000? '$'+nf1.format(b/1000)+'tn' : b>=10? '$'+nf0.format(b)+'bn' : b>=1? '$'+nf1.format(b)+'bn'
    : '$'+nf0.format(v/1e6)+'m'; }
function ruleNum(v,u){ // a threshold as people write it
  if(!isNum(v)) return '';
  if(u==='$bn') return '$'+(+(v/1e9).toFixed(2))+'bn';
  const t=+(+v).toFixed(2);
  return u==='%'? t+'%' : u==='x'? t+'x' : u==='of 9'? t+' of 9' : String(t);
}
const esc = s => String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

// ---------------------------------------------------------------- rules
const realSel = () => [...S.sel].filter(i=>i!==OTHER).sort((a,b)=>a-b);
const allSel = () => realSel().length===SEC.length && S.sel.has(OTHER);
function ruleSector(){ const r=realSel(); return (r.length===1 && !S.sel.has(OTHER))? SEC[r[0]] : DATA.allLabel; }
const isMulti = () => ruleSector()===DATA.allLabel;
const perOn = () => S.per && isMulti();
const theSet = () => S.preset==='tuned'? DATA.tuned : DATA.base;
const blockVals = g => Object.assign({}, DATA.defaults, theSet()[g]||{});
const sectorsInPlay = () => realSel().map(i=>SEC[i]).concat(S.sel.has(OTHER)?[DATA.allLabel]:[]);
// A bar moves along its sector's own distribution, so one step of the slider admits or
// removes roughly the same share of companies on every rule, whatever its units.
function whyNot(g,k){ const w=((DATA.notUsedWhy||{})[S.preset==='tuned'? g : DATA.allLabel]||{})[k]; return w||['not used',''];
}
function distOf(g,key){ if(S.preset!=='tuned') return (DATA.marketDist||{})[key]; const d=(DATA.dist||{})[g]; return d && d[key]; }
function pctOf(v,d){
  if(v<=d[0]) return 0; if(v>=d[20]) return 100;
  for(let i=0;i<20;i++){ if(v<=d[i+1]){ const w=d[i+1]-d[i]; return (i+(w?(v-d[i])/w:0))*5; } }
  return 100;
}
function valAt(p,d){ p=Math.max(0,Math.min(100,p)); const x=p/5, i=Math.min(19,Math.floor(x)); return d[i]+(d[i+1]-d[i])*(x-i); }
function adj(v,key,g){
  if(!isNum(v)) return null;
  if(!S.loosen || (DATA.fixed||{})[key]!==undefined || (S.preset==='tuned' && ((DATA.fixedBy||{})[g]||[]).includes(key))) return v;
  const d=distOf(g,key); if(!d) return v;
  const f=DATA.filters.find(x=>x.key===key);
  const p=pctOf(v,d), p2=f.dir==='min'? p-S.loosen : p+S.loosen;
  let out=valAt(Math.max(1,Math.min(99,p2)),d);
  if(key==='piotroski_min') out=Math.round(out);
  return +out.toFixed(4);
}
const ruleSetFor = g => (DATA.ruleSets||{})[g] || (DATA.ruleSets||{})[DATA.allLabel] || DATA.filters.map(f=>f.key);
const usedIn = (g,key) => S.preset!=='tuned' || ruleSetFor(g).includes(key);
// the rules one company is judged on, after ticks and strictness
function rulesFor(g){
  const per=perOn(), base = per? blockVals(g) : S.val;
  const distSector = per? g : ruleSector();
  const out=[];
  for(const f of DATA.filters){
    if(!S.on[f.key]) continue;
    const v=adj(base[f.key], f.key, distSector);
    if(v===null) continue;
    out.push({key:f.key, field:f.field, dir:f.dir, val:v, pos:(DATA.positiveOnly||[]).includes(f.key)});
  }
  return out;
}
function resetRules(){
  S.val = blockVals(ruleSector());
  // per-sector: every rule starts ticked, and a rule no ticked sector uses is shown disabled
  for(const f of DATA.filters) S.on[f.key] = perOn()? true : isNum(S.val[f.key]);
}
function ruleKey(){ return [S.preset, ruleSector(), perOn()].join('|'); }

// ---------------------------------------------------------------- sidebar: sectors
function buildSectors(){
  const counts=new Array(SEC.length).fill(0); let other=0;
  for(const r of ROWS){ const g=r[IX.gics]; if(g===null||g===undefined) other++; else counts[g]++; }
  const box=$('sec-list'); box.innerHTML='';
  const add=(val,name,n)=>{
    const l=document.createElement('label'); l.className='tick';
    l.innerHTML='<input class="box" type="checkbox" value="'+val+'"><span class="t">'+esc(name)+
      '</span><span class="c num">'+nf0.format(n)+'</span>';
    l.querySelector('input').addEventListener('change',e=>{
      e.target.checked? S.sel.add(val) : S.sel.delete(val); changed(true);
    });
    box.appendChild(l);
  };
  SEC.forEach((s,i)=>add(i,s,counts[i]));
  if(other){ const d=document.createElement('div'); d.className='sep'; box.appendChild(d);
    add(OTHER,'Unclassified',other); }
}
function syncSectors(){
  $('sec-list').querySelectorAll('input').forEach(cb=>cb.checked=S.sel.has(parseInt(cb.value,10)));
  const r=realSel();
  $('sv-sec').textContent = !S.sel.size? 'None' : allSel()? 'All' :
    (r.length===1 && !S.sel.has(OTHER))? SEC[r[0]] : S.sel.size+' ticked';
}

// ---------------------------------------------------------------- sidebar: criteria
function buildCriteria(){
  const box=$('crit'); box.innerHTML='';
  const wrap=document.createElement('div'); wrap.className='cg'; box.appendChild(wrap);
  for(const g of GROUPS){
    const fs=DATA.filters.filter(f=>CRIT[f.key] && CRIT[f.key].g===g);
    if(!fs.length) continue;
    const h=document.createElement('h4'); h.textContent=g; wrap.appendChild(h);
    for(const f of fs){
      const m=CRIT[f.key], id='cr-'+f.key;
      const row=document.createElement('div'); row.className='crit'; row.dataset.k=f.key;
      row.innerHTML='<input class="box" type="checkbox" id="'+id+'">'+
        '<label class="nm" for="'+id+'">'+esc(m.name)+'</label>'+
        '<div class="val"></div><div class="eff"><span class="eff-l"></span><span class="cut num"></span></div>';
      row.querySelector('.box').addEventListener('change',e=>{
        S.on[f.key]=e.target.checked;
        // ticking a rule that has no value yet gives it the default bar, so the tick does something
        if(e.target.checked && !perOn() && !isNum(S.val[f.key]) && isNum(DATA.defaults[f.key]))
          S.val[f.key]=DATA.defaults[f.key];
        changed();
      });
      wrap.appendChild(row);
    }
  }
}
let inputTimer=null;
function syncCriteria(fails, poolN){
  const per=perOn(), live=sectorsInPlay();
  let on=0, avail=0;
  document.querySelectorAll('#crit .crit').forEach(row=>{
    const k=row.dataset.k, f=DATA.filters.find(x=>x.key===k), m=CRIT[k];
    const cb=row.querySelector('.box'), val=row.querySelector('.val'), lab=row.querySelector('.eff-l');
    const op='<span class="op">'+(f.dir==='min'?'&ge;':'&le;')+'</span>';
    let usable=true;
    if(per){
      const users=live.filter(g=>isNum(blockVals(g)[k]));
      const vals=users.map(g=>adj(blockVals(g)[k],k,g)).filter(isNum);
      usable=vals.length>0;
      if(!usable){ const w=whyNot(live.length===1? live[0] : ruleSector(),k);
        val.innerHTML='<span class="rng">not used</span>'; lab.textContent=live.length===1? w[1] : 'not used in the ticked sectors'; }
      else{
        const lo=Math.min(...vals), hi=Math.max(...vals);
        const span = lo===hi? ruleNum(lo,m.u)
          : m.u==='of 9'? ruleNum(lo,'')+' to '+ruleNum(hi,m.u) : ruleNum(lo,m.u)+' to '+ruleNum(hi,m.u);
        val.innerHTML=op+'<span class="rng num">'+span+'</span>';
        lab.textContent = live.length>1? 'used in '+users.length+' of '+live.length+' sectors' : '';
      }
    } else {
      const raw=S.val[k], scale=m.scale||1;
      if(!val.querySelector('input')){
        val.innerHTML=op+'<input type="number" step="any" aria-label="'+esc(m.name)+' threshold">'+
          '<span class="u">'+esc(m.u==='of 9'?'/ 9':m.u)+'</span>';
        val.querySelector('input').addEventListener('input',e=>{
          const t=e.target.value.trim();
          S.val[k]= t===''? null : parseFloat(t)*scale;
          if(t!=='') S.on[k]=true;
          clearTimeout(inputTimer); inputTimer=setTimeout(()=>changed(),160);
        });
      }
      const inp=val.querySelector('input');
      if(document.activeElement!==inp) inp.value = isNum(raw)? +(raw/scale).toFixed(4) : '';
      const eff=adj(raw,k,ruleSector());
      const notUsed = S.preset==='tuned' && !isNum(raw) && !ruleSetFor(ruleSector()).includes(k);
      lab.textContent = (S.loosen && isNum(raw) && eff!==raw)? 'applied as '+ruleNum(eff,m.u)
        : notUsed? 'Not used: '+whyNot(ruleSector(),k)[1] : '';
    }
    cb.disabled=!usable;
    cb.checked=usable && !!S.on[k];
    row.classList.toggle('off', !cb.checked);
    const n=fails? fails[k]||0 : 0;
    row.querySelector('.cut').textContent = (cb.checked && poolN && usable)? (n? 'removes '+nf0.format(n) : 'removes none') : '';
    if(usable){ avail++; if(cb.checked) on++; }
  });
  $('sv-crit').textContent = on+' of '+avail+' on';
}

// ---------------------------------------------------------------- rules by sector
function openMatrix(){
  const live=sectorsInPlay(), per=perOn();
  const cols = per? live : [isMulti()? DATA.allLabel : ruleSector()];
  const name = g => g===DATA.allLabel? (per? 'Unclassified' : 'Every ticked sector') : g;
  let html='<thead><tr><th scope="col">Rule</th>'+cols.map(g=>'<th scope="col">'+esc(name(g))+'</th>').join('')+'</tr></thead><tbody>';
  for(const grp of GROUPS){
    const fs=DATA.filters.filter(f=>CRIT[f.key]&&CRIT[f.key].g===grp);
    if(!fs.length) continue;
    html+='<tr><td class="mx-g" colspan="'+(cols.length+1)+'">'+esc(grp)+'</td></tr>';
    for(const f of fs){
      const m=CRIT[f.key];
      html+='<tr><td>'+esc(m.name)+'</td>'+cols.map(g=>{
        const base = per? blockVals(g) : S.val;
        if(!S.on[f.key]) return '<td class="off">unticked</td>';
        if(!isNum(base[f.key])){
          const w = (S.preset!=='tuned' || !ruleSetFor(g).includes(f.key))? whyNot(g,f.key) : ['no bar set',''];
          return '<td class="off" title="'+esc(w[1])+'">'+esc(w[0])+'</td>';
        }
        return '<td class="num">'+(f.dir==='min'?'&ge; ':'&le; ')+ruleNum(adj(base[f.key],f.key,per?g:ruleSector()),m.u)+'</td>';
      }).join('')+'</tr>';
    }
  }
  $('mx-t').innerHTML=html+'</tbody>';
  // say which set is on screen, and let it be switched right here
  $('mx-h').textContent='Rules by sector: '+(S.preset==='tuned'? 'Sector-tuned' : 'Same for all');
  document.querySelectorAll('input[name=mxpreset]').forEach(r=>{ r.checked=r.value===S.preset;
    r.onchange=e=>{ const main=document.querySelector('input[name=preset][value='+e.target.value+']');
      if(main && !main.checked) main.click(); openMatrix(); }; });
  $('mx-sub').textContent = (S.preset==='tuned'
    ? 'Each sector has its own bars: the median company of that sector (P/E: the median of its quality peers). Financials sit at the 70th percentile and Real Estate and Utilities at the 55th, because they are judged on fewer rules. Computed from today\'s data'
    : 'One set of bars for every sector: the median company of the whole market')
    + (S.loosen? ', moved '+Math.abs(S.loosen)+' percentile points '+(S.loosen>0?'looser':'stricter')+' by the strictness slider' : '')
    + '. Where a rule is not used, the cell says why; hover over it for the full reason.';
  if(!$('mx').open) $('mx').showModal();
}
function ordinal(n){ const s=['th','st','nd','rd'], v=n%100; return n+(s[(v-20)%10]||s[v]||s[0]); }

// ---------------------------------------------------------------- momentum
// Computed from the TradingView snapshot, so it describes where the trend stands today.
// The "How early" slider. Level 3 is the original test. Lower levels want a stock that has
// run less: closer to its averages, cooler RSI, further below its high. Higher levels accept
// a move that is already under way.
const EARLY_LEVELS = [
  {name:'Very early',     fresh:0.05, rsiLo:45, rsiHi:60, roomLo:0.10, roomHi:0.35, e50Lo:0.95, e50Hi:1.05},
  {name:'Early',          fresh:0.10, rsiLo:48, rsiHi:65, roomLo:0.08, roomHi:0.30, e50Lo:0.96, e50Hi:1.08},
  {name:'Balanced',       fresh:0.15, rsiLo:50, rsiHi:70, roomLo:0.05, roomHi:0.25, e50Lo:0.97, e50Hi:1.10},
  {name:'Relaxed',        fresh:0.20, rsiLo:50, rsiHi:73, roomLo:0.03, roomHi:0.25, e50Lo:0.97, e50Hi:1.15},
  {name:'Already moving', fresh:0.25, rsiLo:50, rsiHi:75, roomLo:0.00, roomHi:0.25, e50Lo:0.97, e50Hi:1.20},
];
const EL = () => EARLY_LEVELS[Math.min(5,Math.max(1,S.early||3))-1];
const pc = x => Math.round(x*100)+'%';
const SIGS = [
  {k:'trend', name:'Trend turning up', need:['e50','e200','px'], hint:L=>'Price above its 50 day average, and the 50 day between '+pc(L.e50Lo-1)+' and +'+pc(L.e50Hi-1)+' of the 200 day'},
  {k:'fresh', name:'Not yet stretched', need:['e50','px'], hint:L=>'Price no more than '+pc(L.fresh)+' above its 50 day average'},
  {k:'rsi',   name:'RSI in range', need:['rsi'], hint:L=>'RSI between '+L.rsiLo+' and '+L.rsiHi+': strength without being overbought'},
  {k:'accel', name:'Gaining pace', need:['m6','y1'], hint:L=>'Recent return ahead of the longer pace (3 months against 6, or 6 against 12)'},
  {k:'room',  name:'Room below its 52 week high', need:['h52','px'], hint:L=>pc(L.roomLo)+' to '+pc(L.roomHi)+' below the 52 week high'},
  {k:'vol',   name:'Volume confirms', need:['rv'], hint:L=>'Relative volume at least 1.2 times normal'},
];
const STAGES = [
  {k:'early', name:'Early: just turning up'}, {k:'established', name:'Established uptrend'},
  {k:'extended', name:'Extended: far above its average'}, {k:'neutral', name:'No clear trend'},
  {k:'weak', name:'Downtrend'},
];
let SIG_AVAIL = {};
function momSignals(r){
  const g=k=>r[IX[k]], px=g('px'), e50=g('e50'), e200=g('e200'), rsi=g('rsi');
  const m3=IX.m3!==undefined? g('m3') : null, m6=g('m6'), y1=g('y1'), h52=IX.h52!==undefined? g('h52'):null, rv=IX.rv!==undefined? g('rv'):null;
  const s={}, L=EL();
  s.trend = isNum(px)&&isNum(e50)&&isNum(e200)? (px>e50 && e50>=L.e50Lo*e200 && e50<=L.e50Hi*e200) : null;
  s.fresh = isNum(px)&&isNum(e50)? px<=e50*(1+L.fresh) : null;
  s.rsi = isNum(rsi)? (rsi>=L.rsiLo && rsi<=L.rsiHi) : null;
  s.accel = isNum(m3)&&isNum(m6)? (m3>0 && m3>m6/2) : (isNum(m6)&&isNum(y1)? (m6>0 && m6>y1/2) : null);
  s.room = isNum(h52)&&isNum(px)&&h52>0? (px/h52-1<=-L.roomLo && px/h52-1>=-L.roomHi) : null;
  s.vol = isNum(rv)? rv>=1.2 : null;
  return s;
}
function momStage(r){
  const g=k=>r[IX[k]], px=g('px'), e50=g('e50'), e200=g('e200'), rsi=g('rsi');
  if(!isNum(px)||!isNum(e50)||!isNum(e200)) return 'na';
  const sg=r._sig||(r._sig=momSignals(r));
  const need=[...S.sigs].filter(k=>SIG_AVAIL[k]);
  if(need.length && need.every(k=>sg[k]===true)) return 'early';
  if(px>e50*(1.10+EL().fresh) || (isNum(rsi)&&rsi>75&&px>e50)) return 'extended';
  if(px<e50 && e50<e200) return 'weak';
  if(px>e50 && e50>e200*1.10) return 'established';
  return 'neutral';
}
function momScore(r){
  const sg=r._sig||(r._sig=momSignals(r)), w={trend:25,fresh:15,rsi:20,accel:25,room:10,vol:5};
  let got=0, tot=0;
  for(const k in w){ if(sg[k]===null||sg[k]===undefined) continue; tot+=w[k]; if(sg[k]) got+=w[k]; }
  return tot? Math.round(got/tot*100) : null;
}
const MOM_LABEL={early:'Early',established:'Established',extended:'Extended',neutral:'No trend',weak:'Downtrend',na:'n/a'};
function buildMomentum(){
  const has=k=>IX[k]!==undefined && ROWS.some(r=>isNum(r[IX[k]]));
  SIG_AVAIL={}; for(const sg of SIGS) SIG_AVAIL[sg.k]=sg.need.every(has);
  $('mom-sigs').innerHTML=SIGS.map(sg=>'<label class="tick'+(SIG_AVAIL[sg.k]?'':' dis sig-na')+'" data-h="'+sg.k+'" title="'+esc(sg.hint(EL()))+'"><input class="box" type="checkbox" data-s="'+sg.k+'"'+(SIG_AVAIL[sg.k]?'':' disabled')+'><span class="t">'+esc(sg.name)+'</span>'+(SIG_AVAIL[sg.k]?'':'<span class="c">needs data</span>')+'</label>').join('');
  $('mom-stages').innerHTML=STAGES.map(st=>'<label class="tick"><input class="box" type="checkbox" data-st="'+st.k+'"><span class="t">'+esc(st.name)+'</span><span class="c num" data-ct="'+st.k+'"></span></label>').join('');
  $('mom-sigs').querySelectorAll('input').forEach(i=>i.onchange=e=>{ e.target.checked? S.sigs.add(e.target.dataset.s) : S.sigs.delete(e.target.dataset.s); changed(); });
  $('mom-stages').querySelectorAll('input').forEach(i=>i.onchange=e=>{ e.target.checked? S.stages.add(e.target.dataset.st) : S.stages.delete(e.target.dataset.st); if(!S.momOn){ S.momOn=true; $('mom-on').checked=true; } changed(); });
  const missing=SIGS.filter(sg=>!SIG_AVAIL[sg.k]).map(sg=>sg.name.toLowerCase());
  $('mom-hint').textContent='Early means every ticked condition holds. Hover a condition to see its test.'
    + (missing.length? ' Add 1 month, 3 month, 52 week high and relative volume columns to the TradingView export to switch on '+missing.join(' and ')+'.' : '');
}
function syncMomentum(counts){
  $('mom-on').checked=S.momOn;
  const L=EL(); $('e-rng').value=S.early; $('e-out').textContent=L.name;
  $('e-sub').textContent='Up to '+pc(L.fresh)+' above the 50 day, RSI '+L.rsiLo+' to '+L.rsiHi;
  $('mom-sigs').querySelectorAll('[data-h]').forEach(l=>{ const sg=SIGS.find(x=>x.k===l.dataset.h); if(sg) l.title=sg.hint(L); });
  $('mom-stages').querySelectorAll('input').forEach(i=>i.checked=S.stages.has(i.dataset.st));
  $('mom-sigs').querySelectorAll('input').forEach(i=>i.checked=S.sigs.has(i.dataset.s)&&!i.disabled);
  $('mom-stages').querySelectorAll('[data-ct]').forEach(c=>c.textContent=nf0.format(counts[c.dataset.ct]||0));
  $('sv-mom').textContent = S.momOn? ([...S.stages].map(k=>MOM_LABEL[k]).join(', ')||'None')+(S.stages.has('early')? ' ('+L.name.toLowerCase()+')' : '') : 'Off';
}

// ---------------------------------------------------------------- the screen
function changed(rulesMayMove){
  const key=ruleKey();
  if(rulesMayMove && key!==lastRuleKey){ lastRuleKey=key; resetRules(); }
  S.page=1; run(); save();
}
function run(){
  const strict=S.strict, q=S.q.trim().toLowerCase(), per=perOn();
  const common = per? null : rulesFor(null);
  const cache={};
  const fails={}; DATA.filters.forEach(f=>fails[f.key]=0);
  const stageCt={};
  let poolN=0, nUS=0, nNon=0, nMq=0;
  const out=[];
  for(const r of ROWS){
    const g=r[IX.gics], key=(g===null||g===undefined)? OTHER : g;
    if(!S.sel.has(key)) continue;
    const isUS=r[IX.ctry]==='United States';
    if(isUS) nUS++; else nNon++;
    if(r[IX.mq]) nMq++;
    if(isUS? !S.us : !S.non) continue;
    if(S.mq && !r[IX.mq]) continue;
    if(q && !((r[IX.sym]||'').toLowerCase().includes(q)||(r[IX.name]||'').toLowerCase().includes(q))) continue;
    poolN++;
    const sec = key===OTHER? DATA.allLabel : SEC[key];
    const rules = per? (cache[sec]=cache[sec]||rulesFor(sec)) : common;
    let ok=true, reported=0;
    for(const f of rules){
      const v=r[IX[f.field]];
      if(v===null||v===undefined){ if(strict){ ok=false; fails[f.key]++; } continue; }
      reported++;
      if((f.dir==='min'? v<f.val : v>f.val) || (f.pos && v<=0)){ ok=false; fails[f.key]++; }
    }
    // the same data-coverage guard the script applies, so counts match the workbook
    if(ok && !strict && rules.length && DATA.minCoverage>0 && reported/rules.length<DATA.minCoverage) ok=false;
    if(!ok) continue;
    const st=momStage(r); r._stage=st; stageCt[st]=(stageCt[st]||0)+1;
    if(S.momOn && !S.stages.has(st)) continue;
    out.push(r);
  }
  cur=out;
  $('c-us').textContent=nf0.format(nUS); $('c-non').textContent=nf0.format(nNon); $('c-mq').textContent=nf0.format(nMq);
  syncSectors(); syncCriteria(fails, poolN); syncMomentum(stageCt); summary(); renderTable();
}

function summary(){
  const n=cur.length, us=cur.filter(r=>r[IX.ctry]==='United States').length;
  $('n').textContent=nf0.format(n);
  $('n-lab').textContent= n===1? 'company passes' : 'companies pass';
  $('k-us').textContent=nf0.format(us); $('k-non').textContent=nf0.format(n-us);
  $('k-mq').textContent=nf0.format(cur.filter(r=>r[IX.mq]).length);
  const caps=cur.map(r=>r[IX.mcap]).filter(isNum).sort((a,b)=>a-b);
  $('k-cap').textContent= caps.length? bn(caps[Math.floor(caps.length/2)]) : '-';
  // by sector
  const by={}; for(const r of cur){ const s=SEC[r[IX.gics]]||'Unclassified'; by[s]=(by[s]||0)+1; }
  const order=SEC.concat(['Unclassified']).filter(s=>by[s]);
  $('secbar').innerHTML = order.length>1? order.map(s=>'<span class="chip">'+esc(s)+' <b class="num">'+nf0.format(by[s])+'</b></span>').join('') : '';
  // plain description of the rules in force
  const bits=[S.preset==='tuned'? 'Sector-tuned thresholds' : 'The same thresholds for every sector'];
  if(perOn()) bits.push('each company judged on its own sector\'s rules');
  else if(isMulti()) bits.push('one rule set applied to every ticked sector');
  else bits.push(ruleSector()+' rules');
  if(S.loosen) bits.push((S.loosen>0? 'loosened by ' : 'tightened by ')+Math.abs(S.loosen)+' percentile points');
  if(S.momOn) bits.push('momentum: '+([...S.stages].map(k=>MOM_LABEL[k].toLowerCase()).join(', ')||'none')+(S.stages.has('early')? ' ('+EL().name.toLowerCase()+')' : ''));
  const off=DATA.filters.filter(f=>!S.on[f.key]).length;
  const line=bits.join(', ')+'.'+(S.strict? ' Companies missing any screened figure are excluded.' : '');
  $('rule-line').textContent=line;
  $('sv-mkt').textContent=(S.us&&S.non? 'US and outside' : S.us? 'US only' : S.non? 'Outside US only' : 'None')+(S.mq?', held':'');
  $('sv-rules').textContent=(S.preset==='tuned'?'Tuned':'Same for all')+(S.loosen? (S.loosen>0?', looser ':', stricter ')+Math.abs(S.loosen)+' pts' : '');
  // both sets start at a median, so the slider reads as "beat X% of the sector / market"
  const beat=50-S.loosen, grp=S.preset==='tuned'? 'its own sector' : 'the whole market';
  $('s-mid').textContent = 'Median';
  $('s-out').textContent = S.loosen===0? 'Median' : 'Beat '+beat+'%';
  const note = S.preset==='tuned'? ' (banks, property and utilities start higher)' : '';
  $('s-sub').textContent = (S.loosen===0? 'A company must beat the typical company of '+grp+' on every rule'
                                        : 'A company must beat '+beat+'% of '+grp+' on every rule')+note;
  $('preset-hint').textContent = S.preset==='tuned'
    ? 'Each sector gets its own bars: the median company of that sector, on only the measures that fit its business. P/E is compared with the sector\'s quality peers. Financials, Real Estate and Utilities use fewer rules, so each of their bars sits higher (70th, 55th, 55th percentile) to keep passing about as hard as elsewhere. Recomputed on every run.'
    : 'One set of bars for every sector: the median company of the whole market. The like-for-like view, so banks, property and utilities rarely pass.';
  const multi=isMulti();
  $('per').disabled=!multi; $('per-wrap').classList.toggle('dis',!multi);
  $('per-wrap').title = multi? '' : 'Tick more than one sector to use this';
  ['xl','pp','csv'].forEach(id=>$(id).disabled=!n);
  void off;
}

// ---------------------------------------------------------------- table
const COLS = [
  {k:'co',   h:'Company', cls:'co', sort:r=>(r[IX.name]||'').toLowerCase()},
  {k:'gics', h:'Sector', cls:'l', sort:r=>SEC[r[IX.gics]]||'zz', fmt:r=>esc(SEC[r[IX.gics]]||'Unclassified')},
  {k:'ctry', h:'Country', cls:'l', sort:r=>r[IX.ctry]||'', fmt:r=>esc(r[IX.ctry]||'-')},
  {k:'mcap', h:'Market cap', fmt:r=>bn(r[IX.mcap])},
  {k:'pe',   h:'P/E', t:'x'}, {k:'fpe', h:'Fwd P/E', t:'x'},
  {k:'roe',  h:'ROE', t:'p'}, {k:'roa', h:'ROA', t:'p'}, {k:'roic', h:'ROIC', t:'p'},
  {k:'om',   h:'Op. margin', t:'p'}, {k:'nm', h:'Net margin', t:'p'}, {k:'fcfm', h:'FCF margin', t:'p'},
  {k:'de',   h:'Debt / equity', t:'n2'}, {k:'ic', h:'Int. cover', t:'x'}, {k:'cr', h:'Current ratio', t:'n2'},
  {k:'az',   h:'Altman Z', t:'n2'}, {k:'pf', h:'Piotroski', t:'pf'},
  {k:'m1',   h:'1 month', t:'p', signed:true}, {k:'m3', h:'3 months', t:'p', signed:true},
  {k:'m6',   h:'6 months', t:'p', signed:true}, {k:'ytd', h:'YTD', t:'p', signed:true},
  {k:'y1',   h:'1 year', t:'p', signed:true}, {k:'cagr', h:'5 yr a year', t:'p', signed:true,
    get:r=>cagr(r[IX.y5])},
  {k:'mom',  h:'Momentum', cls:'l', sort:r=>momScore(r), fmt:r=>{ const st=r._stage||momStage(r);
    const sc=momScore(r); return '<span class="mo '+st+'" title="Momentum score '+(sc===null?'n/a':sc)+' of 100">'+MOM_LABEL[st]+'</span>'; }},
  {k:'px',   h:'Price', fmt:r=>isNum(r[IX.px])? nf2.format(r[IX.px])+' <span class="u">'+esc(r[IX.ccy]||'')+'</span>' : '-',
    sort:r=>r[IX.px]},
];
function cagr(y5){ if(!isNum(y5)||y5<=-100) return null; return (Math.pow(1+y5/100,0.2)-1)*100; }
const getv = (c,r) => c.get? c.get(r) : r[IX[c.k]];
function cell(c,r){
  if(c.k==='co'){
    const sym=r[IX.sym]||'', star=r[IX.mq]? '<span class="star" title="Held by at least one superinvestor">&#9733;</span>':'';
    let t='<span class="tk">'+esc(sym)+'</span>';
    if(DATA.deepUrl){
      let q='?ticker='+encodeURIComponent(sym);
      if(r[IX.isin]) q+='&isin='+encodeURIComponent(r[IX.isin]);
      if(r[IX.ccy]) q+='&ccy='+encodeURIComponent(r[IX.ccy]);
      t='<a href="'+DATA.deepUrl+'/'+q+'" target="_blank" rel="noopener" title="Open the full analysis of '+esc(sym)+'">'+esc(sym)+'</a>';
    }
    return '<td class="co">'+t+star+'<span class="nm" title="'+esc(r[IX.name])+'">'+esc(r[IX.name]||'')+'</span></td>';
  }
  if(c.fmt) return '<td class="'+(c.cls||'num')+'">'+c.fmt(r)+'</td>';
  const v=getv(c,r);
  if(!isNum(v)) return '<td class="na">-</td>';
  let s;
  if(c.t==='x') s=nf1.format(v)+'x';
  else if(c.t==='p') s=(c.signed&&v>0?'+':'')+nf1.format(v)+'%';
  else if(c.t==='pf') s=nf0.format(v);
  else s=nf2.format(v);
  const neg=(c.t==='p' && v<0) || (c.k==='pe' && v<0);
  return '<td class="num'+(neg?' neg':'')+'">'+s+'</td>';
}
function renderTable(){
  const th=$('thead');
  th.innerHTML=COLS.map((c,i)=>'<th class="'+(c.cls||'')+'" data-i="'+i+'" scope="col" aria-sort="'+
    (S.sortKey===c.k?(S.sortDir<0?'descending':'ascending'):'none')+'">'+esc(c.h)+
    '<span class="ar">'+(S.sortKey===c.k?(S.sortDir<0?' \u25be':' \u25b4'):'')+'</span></th>').join('');
  th.querySelectorAll('th').forEach(e=>e.onclick=()=>{
    const c=COLS[+e.dataset.i];
    if(S.sortKey===c.k) S.sortDir=-S.sortDir; else { S.sortKey=c.k; S.sortDir=(c.cls==='l'||c.k==='co')?1:-1; }
    S.page=1; renderTable(); save();
  });
  const c=COLS.find(x=>x.k===S.sortKey)||COLS[3];
  const key=c.sort||(r=>getv(c,r));
  const rows=[...cur].sort((x,y)=>{
    const a=key(x), b=key(y);
    const na=(a===null||a===undefined||Number.isNaN(a)), nb=(b===null||b===undefined||Number.isNaN(b));
    if(na&&nb) return 0; if(na) return 1; if(nb) return -1;
    if(typeof a==='string') return a.localeCompare(b)*S.sortDir;
    return (a-b)*S.sortDir;
  });
  const pages=Math.max(1,Math.ceil(rows.length/S.perPage));
  S.page=Math.min(S.page,pages);
  const from=(S.page-1)*S.perPage, slice=rows.slice(from, from+S.perPage);
  $('tbody').innerHTML=slice.map(r=>'<tr>'+COLS.map(c=>cell(c,r)).join('')+'</tr>').join('');
  const empty=!rows.length;
  $('empty').hidden=!empty; document.querySelector('#tw table').style.display=empty?'none':'';
  if(empty){
    if(!S.sel.size){ $('empty-h').textContent='No sector is ticked';
      $('empty-p').textContent='Tick at least one sector on the left to start the screen.'; }
    else if(!S.us && !S.non){ $('empty-h').textContent='No market is ticked';
      $('empty-p').textContent='Tick US, outside the US, or both, under Market.'; }
    else { $('empty-h').textContent='No company passes these rules';
      $('empty-p').textContent='Move the strictness slider towards looser, or untick the rule that removes the most companies. Its count is shown in red under each criterion.'; }
  }
  $('showing').textContent= empty? '' : 'Showing '+nf0.format(from+1)+' to '+nf0.format(from+slice.length)+' of '+nf0.format(rows.length);
  // pager with a window of page numbers
  const pg=$('pager'); pg.innerHTML='';
  if(pages>1){
    const btn=(label,p,cur_,dis)=>{ const b=document.createElement('button'); b.type='button'; b.textContent=label;
      if(cur_) b.setAttribute('aria-current','page'); b.disabled=!!dis;
      b.onclick=()=>{ S.page=p; renderTable(); $('tw').scrollTop=0; }; pg.appendChild(b); };
    btn('Previous',S.page-1,false,S.page===1);
    const show=new Set([1,pages,S.page-1,S.page,S.page+1].filter(p=>p>=1&&p<=pages));
    let last=0;
    [...show].sort((a,b)=>a-b).forEach(p=>{ if(p-last>1){ const s=document.createElement('span'); s.className='gap'; s.textContent='...'; pg.appendChild(s);} btn(String(p),p,p===S.page); last=p; });
    btn('Next',S.page+1,false,S.page===pages);
  }
}

// ---------------------------------------------------------------- exports
const LIBS = {
  ExcelJS:   ['https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js',
              'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js'],
  PptxGenJS: ['https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js',
              'https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundle.js'],
};
function loadLib(name){
  if(window[name]) return Promise.resolve(window[name]);
  return LIBS[name].reduce((p,url)=>p.catch(()=>new Promise((res,rej)=>{
    const s=document.createElement('script'); s.src=url; s.async=true;
    s.onload=()=> window[name]? res(window[name]) : rej(new Error('empty'));
    s.onerror=()=>{ s.remove(); rej(new Error('blocked')); };
    document.head.appendChild(s);
  })), Promise.reject()).catch(()=>{ throw new Error('Could not load the export library. The first export needs an internet connection; check that cdn.jsdelivr.net is not blocked.'); });
}
function selectionLabel(){
  const r=realSel();
  if(allSel()) return 'All sectors';
  if(r.length===SEC.length) return 'All eleven GICS sectors';
  const names=r.map(i=>SEC[i]).concat(S.sel.has(OTHER)?['Unclassified']:[]);
  return names.length<=1? names.join('') : names.slice(0,-1).join(', ')+' and '+names[names.length-1];
}
// ---------------------------------------------------------------- why each name passed
// Reads the rules the company was actually judged on, and where each of its figures sits in
// the same spread the bars came from (its own sector for the tuned set, the whole market for
// the same-for-all set).
function passReasons(r){
  const key=r[IX.gics], g=(key===null||key===undefined)? DATA.allLabel : SEC[key];
  const rules=perOn()? rulesFor(g) : rulesFor(null);
  const ds=perOn()? g : ruleSector(), grp=S.preset==='tuned'? 'sector' : 'market';
  const items=[]; let missing=0;
  for(const f of rules){
    if(f.key==='mcap_min') continue;
    const v=r[IX[f.field]]; if(!isNum(v)){ missing++; continue; }
    const d=distOf(ds,f.key); let rank=null;
    if(d){ const p=pctOf(v,d); rank=f.dir==='min'? p : 100-p; }
    const m=CRIT[f.key];
    items.push({rank, txt:m.short+' '+ruleNum(v,m.u), bar:(f.dir==='min'?'bar \u2265 ':'bar \u2264 ')+ruleNum(f.val,m.u)});
  }
  const ranked=items.filter(x=>x.rank!==null).sort((a,b)=>b.rank-a.rank);
  const where=x=> x.rank>=90? 'top '+Math.max(1,Math.round(100-x.rank))+'% of '+grp
               : x.rank>=50? (grp==='sector'? 'better than ' : 'beats ')+Math.round(x.rank)+'% of '+grp
               : 'only '+Math.round(x.rank)+'% of '+grp;
  const best=ranked.slice(0,3).map(x=>x.txt+', '+where(x));
  const list=ranked.slice(0,3).map(x=>({txt:x.txt, where:where(x)}));
  const worst=ranked.length? ranked[ranked.length-1] : null;
  return {
    passed: 'Cleared '+items.length+' rules'+(missing? ', '+missing+' not reported' : ''),
    why: best.join('; '), list,
    close: worst? worst.txt+' ('+worst.bar+'), '+where(worst) : ''
  };
}
// the sector's median company on the measures a profile slide compares against
function sectorMedians(r){
  const key=r[IX.gics], g=(key===null||key===undefined)? DATA.allLabel : SEC[key], d=(DATA.dist||{})[g]||{};
  const at=k=>d[k]? d[k][10] : null;
  return {pe:at('pe_max'), roic:at('roic_min'), roe:at('roe_min'), gm:at('gross_margin_min'),
          om:at('op_margin_min'), fcfm:at('fcf_margin_min'), de:at('de_max'), az:at('altman_min'),
          roa:at('roa_min'), nm:at('net_margin_min')};
}
// figures that are not believable for an operating business: usually a holding company's
// investment gains or a one-off, so the screen and the scorecard both read them too kindly
function dataChecks(r){
  const out=[], g=k=>r[IX[k]];
  const key=r[IX.gics], used=ruleSetFor((key===null||key===undefined)? DATA.allLabel : SEC[key]);
  // only figures the company is actually judged on: a bank's free cash flow is not used, so
  // an odd value there does not matter
  const big=(k,rule,lab)=>{ const v=g(k); if(used.includes(rule)&&isNum(v)&&Math.abs(v)>100) out.push(lab+' of '+v.toFixed(0)+'% is not plausible for an operating business'); };
  big('om','op_margin_min','Operating margin'); big('nm','net_margin_min','Net margin'); big('fcfm','fcf_margin_min','Free cash flow margin');
  const pe=g('pe'); if(isNum(pe)&&pe>0&&pe<3) out.push('A P/E of '+pe.toFixed(1)+' usually means one-off or investment gains, not repeatable earnings');
  return out;
}
function deepLink(r){
  if(!DATA.deepUrl) return '';
  let q='?ticker='+encodeURIComponent(r[IX.sym]||'');
  if(r[IX.isin]) q+='&isin='+encodeURIComponent(r[IX.isin]);
  if(r[IX.ccy]) q+='&ccy='+encodeURIComponent(r[IX.ccy]);
  return DATA.deepUrl+'/'+q;
}
function deepCard(r){ return (DATA.deep||{})[(r[IX.sym]||'')+'|'+String(r[IX.ccy]||'').toUpperCase()] || null; }
function exportData(){
  const txt=(k,f,v)=>(f.dir==='min'?'\u2265 ':'\u2264 ')+ruleNum(v,CRIT[k].u);
  const asRules=list=>Object.fromEntries(list.map(x=>{ const f=DATA.filters.find(y=>y.key===x.key);
    return [CRIT[x.key].short, txt(x.key,f,x.val)]; }));
  const rules = perOn()
    ? sectorsInPlay().map(g=>({sector: g===DATA.allLabel? 'Unclassified' : g, rules: asRules(rulesFor(g))}))
    : [{sector: isMulti()? 'Every ticked sector' : ruleSector(), rules: asRules(rulesFor(null))}];
  const strictTxt = (S.loosen? (S.loosen>0? ', loosened by ' : ', tightened by ')+Math.abs(S.loosen)+' percentile points' : '')
    + (S.momOn? '; momentum stage: '+([...S.stages].map(k=>MOM_LABEL[k].toLowerCase()).join(', ')||'none')+(S.stages.has('early')? ' ('+EL().name.toLowerCase()+')' : '') : '');
  const rows=cur.map(r=>({
    sym:r[IX.sym], name:r[IX.name], sector:SEC[r[IX.gics]]||'Unclassified',
    ind:(DATA.industries||[])[r[IX.ind]]||'', ctry:r[IX.ctry], mq:r[IX.mq]?1:0,
    inv:r[IX.inv], wt:r[IX.wt], mcap:r[IX.mcap], pe:r[IX.pe], fpe:r[IX.fpe], roe:r[IX.roe],
    roa:r[IX.roa], roic:r[IX.roic], roce:r[IX.roce], gm:r[IX.gm], om:r[IX.om], nm:r[IX.nm],
    fcfm:r[IX.fcfm], cr:r[IX.cr], de:r[IX.de], ic:r[IX.ic], az:r[IX.az], pf:r[IX.pf],
    m1:r[IX.m1], m3:r[IX.m3], m6:r[IX.m6], ytd:r[IX.ytd], y1:r[IX.y1], y5:r[IX.y5], cagr:cagr(r[IX.y5]), px:r[IX.px], ccy:r[IX.ccy],
    mom:MOM_LABEL[r._stage||momStage(r)], momScore:momScore(r),
    ...(()=>{ const pr=passReasons(r), dc=deepCard(r);
      const join=a=>(a||[]).map(x=>String(x).replace(/\.$/,'')).join('; ');
      return { passed:pr.passed, why:pr.why, close:pr.close,
        dScore:dc? dc.s : null, dOp:dc? dc.o : 'Not scored', dQ:dc? dc.q : null, dF:dc? dc.f : null,
        dV:dc? dc.v : null, dM:dc? dc.m : null, dG:dc? dc.g : null, dCov:dc? dc.c : null,
        dLifts:dc? (dc.l||[]) : [], dDrags:dc? (dc.d||[]) : [], dFlags:dc? (dc.x||[]) : [],
        whyList:pr.list, closeTxt:pr.close, med:sectorMedians(r), link:deepLink(r), checks:dataChecks(r),
        e50:r[IX.e50], e200:r[IX.e200], rsi:r[IX.rsi],
        useKeys:ruleSetFor((r[IX.gics]===null||r[IX.gics]===undefined)? DATA.allLabel : SEC[r[IX.gics]]),
        dLift:dc? join(dc.l) : '', dDrag:join(dataChecks(r).map(x=>'Check the data: '+x).concat(dc? (dc.x||[]).map(x=>'Red flag: '+x).concat(dc.d||[]) : [])) }; })()}));
  return {deepNote: 'Deep analysis scorecard from the snapshot figures: each figure scored against its sector, five weighted pillars, Buy from 70, Hold from 45. The full deep analysis page adds price history, cash quality and forensic checks, so its verdict can differ. Names outside the deep feed are marked Not scored.', rows, meta:{
    asOf:DATA.stamp.data, marqueeAsOf:DATA.stamp.marquee, sectorOrder:SEC.concat(['Unclassified']),
    sectorsLabel:selectionLabel(),
    region:S.us&&S.non? 'US and non-US' : S.us? 'US listed only' : 'Listed outside the US only',
    universe:S.mq? 'Superinvestor holdings only' : 'All stocks',
    mode:(S.preset==='tuned'? 'Sector-tuned: every bar is the median of its own sector (Financials 70th, Real Estate and Utilities 55th percentile)' : 'The same thresholds for every sector: the whole-market median')+strictTxt,
    ruleBasis: perOn()? 'Each company judged on its own sector\'s rules'+strictTxt
             : isMulti()? 'One rule set, as set on screen, applied to every ticked sector'+strictTxt
             : 'The '+ruleSector()+' rules, as set on screen'+strictTxt,
    blankPolicy: S.strict? 'A company missing any screened figure is excluded'
      : 'Blanks ignored, but at least '+Math.round(DATA.minCoverage*100)+'% of the screened figures must be reported',
    search:S.q.trim(), rules, ruleLabels: DATA.filters.map(f=>CRIT[f.key].short)}};
}
function fileStem(){
  const r=realSel();
  const tag=(allSel()||r.length>3)? 'sectors' : r.map(i=>SEC[i].replace(/\s+/g,'-')).join('_')||'screen';
  return 'MICS_screen_'+tag+'_'+DATA.stamp.iso;
}
function status(t,err){ const s=$('xstat'); s.textContent=t; s.className='xstat'+(err?' err':''); }
async function doExport(btn,kind){
  const label=btn.textContent; btn.disabled=true; btn.textContent='Preparing...'; status('');
  try{
    const D=exportData();
    if(kind==='xl'){
      const buf=await XP.workbook(await loadLib('ExcelJS'), D);
      const a=document.createElement('a');
      a.href=URL.createObjectURL(new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      a.download=fileStem()+'.xlsx'; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(()=>URL.revokeObjectURL(a.href),4000);
      status('Excel workbook saved, '+nf0.format(D.rows.length)+' companies.');
    } else {
      status('Building '+nf0.format(D.rows.length)+' company pages...');
      const pres=XP.deck(await loadLib('PptxGenJS'), D);
      const raw=await pres.write({outputType:'arraybuffer'});
      const type='application/vnd.openxmlformats-officedocument.presentationml.presentation';
      // the library stores its files uncompressed, which makes a large deck several times
      // bigger than it needs to be, so the file is re-zipped before saving
      let blob;
      try{ const z=await window.JSZip.loadAsync(raw);
        blob=await z.generateAsync({type:'blob', mimeType:type, compression:'DEFLATE', compressionOptions:{level:6}}); }
      catch(e){ blob=new Blob([raw],{type}); }
      const a=document.createElement('a');
      a.href=URL.createObjectURL(blob); a.download=fileStem()+'.pptx'; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(()=>URL.revokeObjectURL(a.href),4000);
      status('PowerPoint saved, '+nf0.format(D.rows.length)+' company pages ('+(blob.size/1048576).toFixed(1)+' MB).');
    }
  } catch(e){ status(e.message||String(e), true); }
  finally{ btn.disabled=!cur.length; btn.textContent=label; }
}
function downloadCsv(){
  const head=['Ticker','Company','Sector','Country','Market cap (USD)','P/E','Fwd P/E','ROE %','ROA %','ROIC %',
    'Op margin %','Net margin %','FCF margin %','Debt/equity','Interest cover','Current ratio','Altman Z','Piotroski',
    '1 month %','3 months %','6 months %','YTD %','1 year %','5 year a year %','Momentum','Momentum score','Price','Currency','Superinvestor held','ISIN',
    'Why it passed','Closest call','Deep score','Deep opinion','Deep quality','Deep financial strength','Deep valuation','Deep momentum','Lifting the score','Holding it back'];
  const q=v=>{ if(v===null||v===undefined||Number.isNaN(v)) return ''; const s=String(v);
    return /[",\n]/.test(s)? '"'+s.replace(/"/g,'""')+'"' : s; };
  const body=cur.map(r=>[r[IX.sym],r[IX.name],SEC[r[IX.gics]]||'Unclassified',r[IX.ctry],r[IX.mcap],r[IX.pe],r[IX.fpe],
    r[IX.roe],r[IX.roa],r[IX.roic],r[IX.om],r[IX.nm],r[IX.fcfm],r[IX.de],r[IX.ic],r[IX.cr],r[IX.az],r[IX.pf],
    r[IX.m1],r[IX.m3],r[IX.m6],r[IX.ytd],r[IX.y1],cagr(r[IX.y5]),MOM_LABEL[r._stage||momStage(r)],momScore(r),r[IX.px],r[IX.ccy],r[IX.mq]?'Yes':'',r[IX.isin]]
    .concat((()=>{ const pr=passReasons(r), dc=deepCard(r), j=a=>(a||[]).join('; ');
      return [pr.why, pr.close, dc&&dc.s, dc? dc.o : 'Not scored', dc&&dc.q, dc&&dc.f, dc&&dc.v, dc&&dc.m,
              dc? j(dc.l) : '', dc? j((dc.x||[]).map(x=>'Red flag: '+x).concat(dc.d||[])) : '']; })()).map(q).join(','));
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob(['\ufeff'+head.join(',')+'\n'+body.join('\n')],{type:'text/csv'}));
  a.download=fileStem()+'.csv'; a.click();
  status('CSV saved, '+nf0.format(cur.length)+' companies.');
}
function copyRules(){
  const tuned=S.preset==='tuned', where=tuned? 'TUNED_OVERRIDES' : 'SECTOR_OVERRIDES';
  const block=(g, base, distSector)=>{
    const lines=DATA.filters.map(f=>{
      const v = S.on[f.key]? adj(base[f.key], f.key, distSector) : null;
      return "        '"+f.key+"'"+' '.repeat(Math.max(1,20-f.key.length))+': '+(v===null?'None':v)+',';
    });
    return "    '"+g+"': {\n"+lines.join('\n')+"\n    },";
  };
  const txt = perOn()? sectorsInPlay().map(g=>block(g, blockVals(g), g)).join('\n\n')
                     : block(ruleSector(), S.val, ruleSector());
  const done=()=>status('Rules copied. Paste them inside '+where+' in marquee_screener.py.');
  navigator.clipboard? navigator.clipboard.writeText(txt).then(done,()=>prompt('Copy these rules:',txt)) : prompt('Copy these rules:',txt);
}

// ---------------------------------------------------------------- remember settings
function rulesSig(){
  // a short fingerprint of today's rule sets, so settings saved against older rules are not reused
  const t=JSON.stringify([DATA.tuned,DATA.base,DATA.ruleSets]); let h=0;
  for(let i=0;i<t.length;i++) h=(h*31+t.charCodeAt(i))|0;
  return String(h);
}
function save(){
  try{ localStorage.setItem(STORE, JSON.stringify({sel:[...S.sel], us:S.us, non:S.non, mq:S.mq, preset:S.preset,
    per:S.per, strict:S.strict, loosen:S.loosen, on:S.on, val:S.val, perPage:S.perPage, sortKey:S.sortKey,
    sortDir:S.sortDir, momOn:S.momOn, early:S.early, sig:rulesSig(), stages:[...S.stages], sigs:[...S.sigs], nsec:SEC.length})); }catch(e){}
}
function restore(){
  try{
    const o=JSON.parse(localStorage.getItem(STORE)||'null');
    if(!o || o.nsec!==SEC.length) return false;
    Object.assign(S,{sel:new Set(o.sel), us:o.us, non:o.non, mq:o.mq, preset:o.preset, per:o.per, strict:o.strict,
      loosen:Math.max(-25,Math.min(25,o.loosen||0)), on:o.on||{}, val:o.val||{}, perPage:o.perPage||50, sortKey:o.sortKey||'mcap', sortDir:o.sortDir||-1,
      momOn:!!o.momOn, early:o.early||3, stages:new Set(o.stages||['early']), sigs:new Set(o.sigs||SIGS.map(s=>s.k))});
    if(o.sig!==rulesSig()){ S.loosen=0; resetRules(); }   // the rules were rebuilt since this was saved
    return true;
  }catch(e){ return false; }
}
function defaults(){
  Object.assign(S,{sel:new Set(SEC.map((_,i)=>i).concat([OTHER])), us:true, non:true, mq:false, q:'',
    preset:DATA.mode, per:DATA.perSector, strict:DATA.blankFails, loosen:0, page:1, perPage:50, sortKey:'mcap', sortDir:-1,
    momOn:false, early:3, stages:new Set(['early']), sigs:new Set(SIGS.map(s=>s.k))});
  resetRules();
  ROWS.forEach(r=>{ delete r._sig; });
}
function syncControls(){
  $('m-us').checked=S.us; $('m-non').checked=S.non; $('m-mq').checked=S.mq; $('q').value=S.q;
  document.querySelectorAll('input[name=preset]').forEach(r=>r.checked=r.value===S.preset);
  $('per').checked=S.per; $('strict').checked=S.strict; $('s-rng').value=S.loosen;
  $('per-page').value=String(S.perPage);
}

// ---------------------------------------------------------------- wiring
function init(){
  $('d-data').textContent=DATA.stamp.data; $('d-mq').textContent=DATA.stamp.marquee;
  $('d-built').textContent=DATA.stamp.builtAt||DATA.stamp.built||'';
  if(DATA.xlsxUrl){ const a=$('xlsx'); a.href=DATA.xlsxUrl; a.style.display=''; }
  if(DATA.deepUrl){ const l=$('idx-link'); l.href=DATA.deepUrl+'/?view=indexes'; l.hidden=false; }
  if(DATA.deepUrl) $('hint').textContent='Click a ticker to open its full analysis. Click a column heading to sort. Exports always contain every company that passes, not only this page.';
  buildSectors(); buildCriteria(); buildMomentum();
  if(!restore()) defaults();
  lastRuleKey=ruleKey();
  syncControls(); run();

  $('sec-all').onclick=()=>{ S.sel=new Set(SEC.map((_,i)=>i).concat([OTHER])); changed(true); };
  $('sec-none').onclick=()=>{ S.sel=new Set(); changed(true); };
  $('m-us').onchange=e=>{ S.us=e.target.checked; changed(); };
  $('m-non').onchange=e=>{ S.non=e.target.checked; changed(); };
  $('m-mq').onchange=e=>{ S.mq=e.target.checked; changed(); };
  let qt=null; $('q').oninput=e=>{ S.q=e.target.value; clearTimeout(qt); qt=setTimeout(()=>changed(),150); };
  document.querySelectorAll('input[name=preset]').forEach(r=>r.onchange=e=>{ S.preset=e.target.value; lastRuleKey=null; changed(true); });
  $('per').onchange=e=>{ S.per=e.target.checked; changed(true); };
  $('strict').onchange=e=>{ S.strict=e.target.checked; changed(); };
  $('s-rng').oninput=e=>{ S.loosen=parseInt(e.target.value,10); clearTimeout(inputTimer); inputTimer=setTimeout(()=>changed(),60); };
  $('e-rng').oninput=e=>{ S.early=parseInt(e.target.value,10); ROWS.forEach(r=>{ delete r._sig; });
    // moving it means the user wants early names, so switch the stage filter on
    if(!S.momOn){ S.momOn=true; S.stages=new Set(['early']); }
    clearTimeout(inputTimer); inputTimer=setTimeout(()=>changed(),60); };
  $('c-all').onclick=()=>{ DATA.filters.forEach(f=>S.on[f.key]=true);
    if(!perOn()) DATA.filters.forEach(f=>{ if(!isNum(S.val[f.key])) S.val[f.key]=isNum(DATA.defaults[f.key])? DATA.defaults[f.key] : null; });
    changed(); };
  $('c-none').onclick=()=>{ DATA.filters.forEach(f=>S.on[f.key]=false); changed(); };
  $('reset').onclick=()=>{ try{ localStorage.removeItem(STORE); }catch(e){} defaults(); lastRuleKey=ruleKey(); syncControls(); S.page=1; run(); status('Settings reset.'); };
  $('copy').onclick=copyRules;
  $('mom-on').onchange=e=>{ S.momOn=e.target.checked; changed(); };
  $('c-matrix').onclick=openMatrix;
  $('mx-close').onclick=()=>$('mx').close();
  $('per-page').onchange=e=>{ S.perPage=parseInt(e.target.value,10); S.page=1; renderTable(); save(); };
  $('xl').onclick=e=>doExport(e.currentTarget,'xl');
  $('pp').onclick=e=>doExport(e.currentTarget,'pp');
  $('csv').onclick=downloadCsv;
}

/*EXPORT_START*/
// ---------------------------------------------------------------------------------
// Export builders. Pure functions: they take the library and a plain data object and
// return the file, so the same code runs in the browser and in a test harness.
//   D = { rows:[{sym,name,sector,ind,ctry,mq,inv,wt,mcap,pe,fpe,roe,roa,roic,roce,
//                gm,om,nm,fcfm,cr,de,ic,az,pf,ytd,y1,y5,cagr,px,ccy}],
//         meta:{asOf, marqueeAsOf, sectors:[names], sectorsLabel, region, universe,
//               mode, ruleBasis, blankPolicy, search, rules:[{sector, rules:{label:txt}}],
//               ruleLabels:[...]} }
// ---------------------------------------------------------------------------------
const XP = (() => {
  const NAVY = '1F3864', NAVY2 = '2E4A7D', GOLD = 'BF8F00', BAND = 'F2F5FA',
        LINE = 'D8DEE9', INK = '1A1A1A', WHITE = 'FFFFFF', PALE = 'E9EEF6';

  const num = v => (v === null || v === undefined || Number.isNaN(v)) ? null : v;
  function median(arr) {
    const a = arr.filter(v => v !== null && v !== undefined && !Number.isNaN(v)).sort((x, y) => x - y);
    if (!a.length) return null;
    const m = Math.floor(a.length / 2);
    return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
  }
  const fmt = {
    x: v => num(v) === null ? '-' : v.toFixed(1) + 'x',
    pct: v => num(v) === null ? '-' : v.toFixed(1) + '%',
    bn: v => num(v) === null ? '-' : (v / 1e9 >= 100 ? Math.round(v / 1e9).toLocaleString('en-US')
                                                     : (v / 1e9).toFixed(1)),
    int: v => num(v) === null ? '-' : Math.round(v).toLocaleString('en-US'),
    z: v => num(v) === null ? '-' : v.toFixed(2),
  };

  function bySector(D) {
    const order = D.meta.sectorOrder || [];
    const groups = new Map();
    for (const r of D.rows) {
      const k = r.sector || 'Unclassified';
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }
    const keys = [...groups.keys()].sort((a, b) => {
      const ia = order.indexOf(a), ib = order.indexOf(b);
      return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
    });
    return keys.map(k => ({ sector: k,
      rows: groups.get(k).slice().sort((a, b) => (b.mcap || 0) - (a.mcap || 0)) }));
  }

  function sectorStats(g) {
    const rows = g.rows;
    return {
      sector: g.sector, n: rows.length,
      us: rows.filter(r => r.ctry === 'United States').length,
      mq: rows.filter(r => r.mq).length,
      pe: median(rows.map(r => r.pe)), roe: median(rows.map(r => r.roe)),
      om: median(rows.map(r => r.om)),
      mcap: rows.reduce((s, r) => s + (r.mcap || 0), 0),
    };
  }

  // ================================================================ EXCEL ==========
  async function workbook(ExcelJS, D) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'MICS International';
    wb.title = 'Sector screen results';
    wb.created = new Date();
    const A = c => 'FF' + c;
    const font = (o = {}) => Object.assign({ name: 'Arial', size: 10, color: { argb: A(INK) } }, o);
    const fill = c => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: A(c) } });
    const hair = { style: 'thin', color: { argb: A(LINE) } };
    const box = { top: hair, left: hair, bottom: hair, right: hair };
    const groups = bySector(D);
    const m = D.meta;

    // ---- Results sheet first in the build (Summary formulas point at it), shown second
    const cols = [
      ['Company', [['sym', 'Ticker', 10, '@'], ['name', 'Company', 36, '@'],
                   ['sector', 'GICS sector', 24, '@'], ['ind', 'Industry', 30, '@'],
                   ['ctry', 'Country', 18, '@']]],
      ['Ownership', [['mqTxt', 'Marquee held', 10, '@'], ['inv', 'Investors holding', 11, '0'],
                     ['wt', 'Marquee weight %', 11, '0.00"%"']]],
      ['Size and value', [['mcapBn', 'Market cap ($bn)', 12, '#,##0.0'],
                          ['pe', 'P/E', 9, '0.0"x"'], ['fpe', 'Forward P/E', 9, '0.0"x"']]],
      ['Returns, trailing twelve months', [['roe', 'ROE %', 9, '0.0"%";-0.0"%"'],
                          ['roa', 'ROA %', 9, '0.0"%";-0.0"%"'],
                          ['roic', 'ROIC %', 9, '0.0"%";-0.0"%"'],
                          ['roce', 'ROCE %', 9, '0.0"%";-0.0"%"']]],
      ['Margins', [['gm', 'Gross %', 9, '0.0"%";-0.0"%"'], ['om', 'Operating %', 10, '0.0"%";-0.0"%"'],
                   ['nm', 'Net %', 9, '0.0"%";-0.0"%"'], ['fcfm', 'FCF %', 9, '0.0"%";-0.0"%"']]],
      ['Balance sheet', [['cr', 'Current ratio', 9, '0.00'], ['de', 'Debt / equity', 9, '0.00'],
                         ['ic', 'Interest cover', 9, '0.0"x"']]],
      ['Quality scores', [['az', 'Altman Z', 9, '0.00'], ['pf', 'Piotroski F', 9, '0']]],
      ['Price performance', [['m1', '1 month %', 9, '0.0"%";-0.0"%"'], ['m3', '3 months %', 9, '0.0"%";-0.0"%"'],
                             ['m6', '6 months %', 9, '0.0"%";-0.0"%"'], ['ytd', 'YTD %', 9, '0.0"%";-0.0"%"'], ['y1', '1 year %', 9, '0.0"%";-0.0"%"'],
                             ['y5', '5 year %', 9, '0.0"%";-0.0"%"'],
                             ['cagr', '5 year a year %', 10, '0.0"%";-0.0"%"'],
                             ['px', 'Price', 11, '#,##0.00'], ['ccy', 'Ccy', 7, '@']]],
      ['Timing', [['mom', 'Momentum stage', 13, '@'], ['momScore', 'Momentum score', 10, '0']]],
    ];
    const flat = cols.flatMap(([g, cs]) => cs.map(c => ({ group: g, key: c[0], label: c[1], w: c[2], f: c[3] })));

    // Summary is created first so it is the first tab; it is filled in further down.
    const sum = wb.addWorksheet('Summary', { views: [{ showGridLines: false }],
                                             properties: { tabColor: { argb: A(GOLD) } } });
    const res = wb.addWorksheet('Results', { views: [{ state: 'frozen', xSplit: 2, ySplit: 5, showGridLines: false }],
                                             properties: { tabColor: { argb: A(NAVY) } } });
    res.columns = flat.map(c => ({ width: c.w }));
    res.getCell('A1').value = 'Screen results';
    res.getCell('A1').font = font({ size: 16, bold: true, color: { argb: A(NAVY) } });
    res.getCell('A2').value = `${D.rows.length.toLocaleString('en-US')} names, ${m.sectorsLabel}. ` +
      `Data as at ${m.asOf}, marquee holdings as at ${m.marqueeAsOf}. Sorted by sector, then market cap.`;
    res.getCell('A2').font = font({ size: 10, color: { argb: A(NAVY) } });

    // group band row 4, column headers row 5
    let c0 = 1;
    for (const [g, cs] of cols) {
      const c1 = c0 + cs.length - 1;
      if (c1 > c0) res.mergeCells(4, c0, 4, c1);
      const cell = res.getCell(4, c0);
      cell.value = g;
      for (let c = c0; c <= c1; c++) {
        const x = res.getCell(4, c);
        x.fill = fill(NAVY); x.border = { left: { style: 'thin', color: { argb: A(WHITE) } } };
      }
      cell.font = font({ bold: true, color: { argb: A(WHITE) } });
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      c0 = c1 + 1;
    }
    flat.forEach((c, i) => {
      const cell = res.getCell(5, i + 1);
      cell.value = c.label;
      cell.fill = fill(NAVY2);
      cell.font = font({ bold: true, size: 9, color: { argb: A(WHITE) } });
      cell.alignment = { horizontal: c.f === '@' ? 'left' : 'right', vertical: 'middle', wrapText: true };
      cell.border = { bottom: { style: 'medium', color: { argb: A(GOLD) } } };
    });
    res.getRow(4).height = 20;
    res.getRow(5).height = 30;

    let r = 6, first = 6, band = 0;
    const sectorRanges = [];
    for (const g of groups) {
      const start = r;
      g.rows.forEach((row, i) => {
        const vals = Object.assign({}, row, { mqTxt: row.mq ? 'Yes' : '',
          mcapBn: num(row.mcap) === null ? null : row.mcap / 1e9 });
        const line = res.getRow(r);
        flat.forEach((c, j) => {
          const cell = line.getCell(j + 1);
          const v = vals[c.key];
          cell.value = (v === undefined || v === null || v === '' || Number.isNaN(v)) ? null : v;
          cell.numFmt = c.f;
          cell.font = font({ size: 9 });
          cell.alignment = { horizontal: c.f === '@' ? 'left' : 'right', vertical: 'middle' };
          if (band % 2) cell.fill = fill(BAND);
          cell.border = { bottom: hair };
          if (i === 0 && r !== first) cell.border = { top: { style: 'thin', color: { argb: A(NAVY) } }, bottom: hair };
        });
        line.getCell(1).font = font({ size: 9, bold: true, color: { argb: A(NAVY) } });
        if (row.mq) line.getCell(6).font = font({ size: 9, bold: true, color: { argb: A(GOLD) } });
        band++; r++;
      });
      sectorRanges.push({ sector: g.sector, start, end: r - 1 });
    }
    const last = Math.max(first, r - 1);
    res.autoFilter = { from: { row: 5, column: 1 }, to: { row: last, column: flat.length } };

    // ---- Summary sheet
    sum.columns = [{ width: 3 }, { width: 28 }, { width: 12 }, { width: 10 }, { width: 10 },
                   { width: 13 }, { width: 13 }, { width: 13 }, { width: 15 }, { width: 17 }];
    sum.getCell('B2').value = 'Sector screen';
    sum.getCell('B2').font = font({ size: 22, bold: true, color: { argb: A(NAVY) } });
    sum.getCell('B3').value = `MICS International   |   Data as at ${m.asOf}   |   Marquee holdings as at ${m.marqueeAsOf}`;
    sum.getCell('B3').font = font({ size: 10, color: { argb: A(NAVY) } });
    for (let c = 2; c <= 10; c++) sum.getCell(4, c).border = { bottom: { style: 'medium', color: { argb: A(GOLD) } } };

    const label = (row, text) => {
      const cell = sum.getCell(row, 2);
      cell.value = text;
      cell.font = font({ size: 12, bold: true, color: { argb: A(NAVY) } });
    };
    label(6, 'How the screen was set');
    const settings = [
      ['Sectors', m.sectorsLabel], ['Region', m.region], ['Universe', m.universe],
      ['Threshold set', m.mode], ['Rules applied', m.ruleBasis], ['Missing values', m.blankPolicy],
    ];
    if (m.search) settings.push(['Search filter', m.search]);
    let sr = 7;
    for (const [k, v] of settings) {
      const a = sum.getCell(sr, 2), b = sum.getCell(sr, 3);
      sum.mergeCells(sr, 3, sr, 10);
      a.value = k; b.value = v;
      a.font = font({ bold: true }); b.font = font();
      a.fill = fill(PALE);
      for (let c = 2; c <= 10; c++) sum.getCell(sr, c).border = { bottom: hair };
      b.alignment = { wrapText: true, vertical: 'top' }; a.alignment = { vertical: 'top' };
      if (String(v).length > 95) sum.getRow(sr).height = 28;
      sr++;
    }

    sr += 1;
    label(sr, 'What came through, by sector'); sr++;
    const heads = ['Sector', 'Names', 'US', 'Non-US', 'Marquee held', 'Median P/E',
                   'Median ROE %', 'Median op margin %', 'Market cap ($bn)'];
    heads.forEach((h, i) => {
      const cell = sum.getCell(sr, i + 2);
      cell.value = h;
      cell.fill = fill(NAVY);
      cell.font = font({ bold: true, size: 9, color: { argb: A(WHITE) } });
      cell.alignment = { horizontal: i ? 'right' : 'left', vertical: 'middle', wrapText: true };
    });
    sum.getRow(sr).height = 28;
    const hdr = sr; sr++;
    const rngs = { sec: `Results!$C$6:$C$${last}`, ctry: `Results!$E$6:$E$${last}`,
                   mq: `Results!$F$6:$F$${last}`, mc: `Results!$I$6:$I$${last}` };
    const tableStart = sr;
    groups.forEach((g, i) => {
      const s = sectorStats(g);
      const vals = [
        s.sector,
        { formula: `COUNTIF(${rngs.sec},B${sr})`, result: s.n },
        { formula: `COUNTIFS(${rngs.sec},B${sr},${rngs.ctry},"United States")`, result: s.us },
        { formula: `C${sr}-D${sr}`, result: s.n - s.us },
        { formula: `COUNTIFS(${rngs.sec},B${sr},${rngs.mq},"Yes")`, result: s.mq },
        s.pe, s.roe, s.om,
        { formula: `SUMIF(${rngs.sec},B${sr},${rngs.mc})`, result: s.mcap / 1e9 },
      ];
      const fmts = ['@', '#,##0', '#,##0', '#,##0', '#,##0', '0.0"x"', '0.0"%"', '0.0"%"', '#,##0.0'];
      vals.forEach((v, j) => {
        const cell = sum.getCell(sr, j + 2);
        cell.value = v === null ? null : v;
        cell.numFmt = fmts[j];
        cell.font = font(j === 0 ? { bold: true, color: { argb: A(NAVY) } } : {});
        cell.alignment = { horizontal: j ? 'right' : 'left' };
        cell.border = { bottom: hair };
        if (i % 2) cell.fill = fill(BAND);
      });
      sr++;
    });
    const tableEnd = sr - 1;
    const all = sectorStats({ sector: 'All', rows: D.rows });
    const tot = [
      'Total',
      { formula: `SUM(C${tableStart}:C${tableEnd})`, result: all.n },
      { formula: `SUM(D${tableStart}:D${tableEnd})`, result: all.us },
      { formula: `SUM(E${tableStart}:E${tableEnd})`, result: all.n - all.us },
      { formula: `SUM(F${tableStart}:F${tableEnd})`, result: all.mq },
      all.pe, all.roe, all.om,
      { formula: `SUM(J${tableStart}:J${tableEnd})`, result: all.mcap / 1e9 },
    ];
    const tf = ['@', '#,##0', '#,##0', '#,##0', '#,##0', '0.0"x"', '0.0"%"', '0.0"%"', '#,##0.0'];
    tot.forEach((v, j) => {
      const cell = sum.getCell(sr, j + 2);
      cell.value = v; cell.numFmt = tf[j];
      cell.font = font({ bold: true, color: { argb: A(NAVY) } });
      cell.alignment = { horizontal: j ? 'right' : 'left' };
      cell.border = { top: { style: 'medium', color: { argb: A(GOLD) } }, bottom: { style: 'thin', color: { argb: A(NAVY) } } };
    });
    sr += 2;
    const note = sum.getCell(sr, 2);
    sum.mergeCells(sr, 2, sr + 1, 10);
    note.value = 'Counts and market cap are live formulas on the Results sheet, so deleting rows there ' +
      'updates this table. Medians were computed at export. Market cap is in US dollars; price is in the ' +
      'listing currency shown beside it. Percentages are stored as reported, so 12.5 means 12.5%.';
    note.font = font({ size: 9, color: { argb: A(NAVY) } });
    note.alignment = { wrapText: true, vertical: 'top' };
    sum.getRow(sr).height = 18; sum.getRow(sr + 1).height = 18;

    // ---- Why they passed: the screen's reasons next to the deep analysis scorecard
    const why = wb.addWorksheet('Why they passed', { views: [{ state: 'frozen', xSplit: 2, ySplit: 5, showGridLines: false }],
                                                    properties: { tabColor: { argb: A(GOLD) } } });
    const wcols = [
      ['sym', 'Ticker', 10], ['name', 'Company', 30], ['sector', 'GICS sector', 20],
      ['passed', 'Screen', 14], ['why', 'Why it passed: its strongest figures', 52],
      ['close', 'Closest call: its weakest figure against the bar', 40],
      ['dScore', 'Deep score (of 100)', 10], ['dOp', 'Deep opinion', 12],
      ['dQ', 'Quality', 8], ['dF', 'Financial strength', 9], ['dV', 'Valuation', 9], ['dM', 'Momentum', 9],
      ['dLift', 'Lifting the score', 52], ['dDrag', 'Holding it back', 52],
      ['mom', 'Momentum stage', 13], ['inv', 'Superinvestors holding', 11], ['link', 'Deep analysis', 12],
    ];
    why.columns = wcols.map(c => ({ width: c[2] }));
    why.getCell('A1').value = 'Why these companies passed';
    why.getCell('A1').font = font({ size: 16, bold: true, color: { argb: A(NAVY) } });
    why.getCell('A2').value = 'Screen columns show where the figures of each company sit in the same spread the bars came from. ' + (D.deepNote || '');
    why.getCell('A2').font = font({ size: 9, color: { argb: A(NAVY) } });
    why.mergeCells(2, 1, 2, 10); why.getRow(2).height = 42;
    why.getCell('A2').alignment = { wrapText: true, vertical: 'top' };
    const wgroups = [['Company', 3], ['The screen', 3], ['Deep analysis scorecard', 8], ['Context', 3]];
    let wc = 1;
    for (const [g, n] of wgroups) {
      if (n > 1) why.mergeCells(4, wc, 4, wc + n - 1);
      const cell = why.getCell(4, wc); cell.value = g;
      for (let c = wc; c < wc + n; c++) why.getCell(4, c).fill = fill(g === 'Deep analysis scorecard' ? GOLD : NAVY);
      cell.font = font({ bold: true, color: { argb: A(WHITE) } }); cell.alignment = { horizontal: 'center' };
      wc += n;
    }
    wcols.forEach((c, i) => {
      const cell = why.getCell(5, i + 1); cell.value = c[1]; cell.fill = fill(NAVY2);
      cell.font = font({ bold: true, size: 9, color: { argb: A(WHITE) } });
      cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    });
    why.getRow(5).height = 32;
    const opColor = { Buy: '1E7B34', Hold: 'BF8F00', Sell: 'B42318' };
    let wr = 6;
    for (const g of groups) for (const row of g.rows) {
      wcols.forEach((c, j) => {
        const cell = why.getCell(wr, j + 1);
        const v = row[c[0]];
        cell.value = (v === undefined || (typeof v === 'number' && Number.isNaN(v))) ? null : v;
        if (c[0] === 'link') cell.value = v ? { text: 'Open', hyperlink: v } : null;
        cell.font = font({ size: 9 });
        cell.alignment = { wrapText: true, vertical: 'top', horizontal: typeof v === 'number' ? 'right' : 'left' };
        cell.border = { bottom: hair };
        if (typeof v === 'number') cell.numFmt = '0';
        if (c[0] === 'dOp' && opColor[v]) cell.font = font({ size: 9, bold: true, color: { argb: A(opColor[v]) } });
        if (c[0] === 'link' && v) cell.font = font({ size: 9, underline: true, color: { argb: A(NAVY2) } });
      });
      if ((wr - 6) % 2) for (let j = 1; j <= wcols.length; j++) why.getCell(wr, j).fill = fill(BAND);
      wr++;
    }

    // ---- Thresholds sheet
    const th = wb.addWorksheet('Thresholds', { views: [{ showGridLines: false, state: 'frozen', ySplit: 4 }],
                                               properties: { tabColor: { argb: A(NAVY2) } } });
    th.getCell('A1').value = 'Thresholds applied';
    th.getCell('A1').font = font({ size: 16, bold: true, color: { argb: A(NAVY) } });
    th.getCell('A2').value = m.ruleBasis;
    th.getCell('A2').font = font({ size: 10, color: { argb: A(NAVY) } });
    const tl = m.ruleLabels;
    th.columns = [{ width: 26 }].concat(tl.map(() => ({ width: 13 })));
    ['Rule set'].concat(tl).forEach((h, i) => {
      const cell = th.getCell(4, i + 1);
      cell.value = h; cell.fill = fill(NAVY);
      cell.font = font({ bold: true, size: 9, color: { argb: A(WHITE) } });
      cell.alignment = { horizontal: i ? 'center' : 'left', vertical: 'middle', wrapText: true };
    });
    th.getRow(4).height = 30;
    m.rules.forEach((rs, i) => {
      const row = 5 + i;
      th.getCell(row, 1).value = rs.sector;
      th.getCell(row, 1).font = font({ bold: true, color: { argb: A(NAVY) } });
      tl.forEach((lab, j) => {
        const cell = th.getCell(row, j + 2);
        cell.value = rs.rules[lab] || 'off';
        cell.font = font(rs.rules[lab] ? {} : { color: { argb: A(NAVY2) }, italic: true });
        cell.alignment = { horizontal: 'center' };
      });
      for (let c = 1; c <= tl.length + 1; c++) {
        const cell = th.getCell(row, c);
        cell.border = { bottom: hair };
        if (i % 2) cell.fill = fill(BAND);
      }
    });

    // print setup: landscape, one page wide, header rows repeated on the results pages
    for (const ws of [sum, res, why, th]) {
      ws.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1,
                       fitToHeight: 0, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5,
                                                  header: 0.2, footer: 0.2 } };
      ws.headerFooter = { oddFooter: '&L&8MICS International, sector screen&R&8Page &P of &N' };
    }
    res.pageSetup.printTitlesRow = '4:5';
    return wb.xlsx.writeBuffer();
  }

  // ================================================================ POWERPOINT =====
  function deck(PptxGenJS, D) {
    const pres = new PptxGenJS();
    pres.layout = 'LAYOUT_WIDE';            // 13.33 x 7.5 in
    pres.author = 'MICS International';
    pres.company = 'MICS International';
    pres.title = 'Sector screen';
    const HEAD = 'Georgia', BODY = 'Calibri';
    pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };
    const m = D.meta;
    const W = 13.333, X0 = 0.6, CW = W - 1.2;
    const foot = `MICS International   |   Sector screen   |   Data as at ${m.asOf}`;

    pres.defineSlideMaster({
      title: 'COVER',
      background: { color: NAVY },
      objects: [
        { line: { x: X0, y: 4.05, w: 1.4, h: 0, line: { color: GOLD, width: 2 } } },
        { placeholder: { options: { name: 'title', type: 'title', x: X0, y: 1.9, w: 11.5, h: 1.5,
            fontFace: HEAD, fontSize: 44, color: WHITE, bold: false, valign: 'bottom', align: 'left', margin: 0 },
            text: 'Title' } },
        { placeholder: { options: { name: 'body', type: 'body', x: X0, y: 4.3, w: 11.5, h: 1.6,
            fontFace: BODY, fontSize: 16, color: 'C9D3E3', valign: 'top', align: 'left', margin: 0 }, text: '' } },
      ],
    });
    pres.defineSlideMaster({
      title: 'CONTENT',
      background: { color: WHITE },
      margin: [0.5, 0.6, 0.75, 0.6],
      objects: [
        { placeholder: { options: { name: 'title', type: 'title', x: X0, y: 0.38, w: CW, h: 0.75,
            fontFace: HEAD, fontSize: 24, color: NAVY, valign: 'bottom', align: 'left', margin: 0 }, text: 'Title' } },
        { line: { x: X0, y: 1.22, w: CW, h: 0, line: { color: LINE, width: 0.75 } } },
        { text: { text: foot, options: { x: X0, y: 7.0, w: 9, h: 0.3, fontFace: BODY, fontSize: 9,
            color: NAVY, margin: 0 } } },
      ],
      slideNumber: { x: 12.1, y: 7.0, w: 0.63, h: 0.3, fontFace: BODY, fontSize: 9, color: NAVY, align: 'right' },
    });

    const T = (text, o = {}) => ({ text: String(text), options: o });
    const head = (cols) => cols.map((c, i) => T(c.h, { bold: true, color: WHITE, fill: { color: NAVY },
      align: c.a || (i ? 'right' : 'left'), valign: 'middle' }));
    const tableOpts = (colW, extra = {}) => Object.assign({
      x: X0, colW, fontFace: BODY, fontSize: 10.5, color: INK, valign: 'middle',
      border: { type: 'solid', pt: 0.5, color: LINE }, margin: [0.03, 0.08, 0.03, 0.08],
      rowH: 0.3, autoPage: false,
    }, extra);
    const zebra = (cells, i) => cells.map(c => { if (i % 2) c.options.fill = { color: BAND }; return c; });

    // 1. cover
    const groups = bySector(D);
    const s1 = pres.addSlide({ masterName: 'COVER' });
    s1.addText('Sector screen', { placeholder: 'title' });
    s1.addText([
      T(groups.length === 1
        ? `${D.rows.length.toLocaleString('en-US')} companies cleared the ${groups[0].sector} screen`
        : `${D.rows.length.toLocaleString('en-US')} companies cleared the screen across ${groups.length} sectors`,
        { breakLine: true, color: WHITE }),
      T(m.sectorsLabel, { breakLine: true, fontSize: 13 }),
      T(`Data as at ${m.asOf}   |   Prepared by MICS International`, { fontSize: 12 }),
    ], { placeholder: 'body' });

    // 2. summary
    const stats = groups.map(sectorStats);
    const all = sectorStats({ sector: 'Total', rows: D.rows });
    const s2 = pres.addSlide({ masterName: 'CONTENT' });
    const lead = stats.slice().sort((a, b) => b.n - a.n)[0];
    s2.addText(stats.length > 1
      ? `${D.rows.length.toLocaleString('en-US')} names passed, ${lead.sector} contributed the most`
      : `${D.rows.length.toLocaleString('en-US')} names passed the ${lead ? lead.sector : ''} screen`,
      { placeholder: 'title' });
    const sc = [{ h: 'Sector', a: 'left' }, { h: 'Names' }, { h: 'US' }, { h: 'Non-US' },
                { h: 'Marquee held' }, { h: 'Median P/E' }, { h: 'Median ROE' },
                { h: 'Median op margin' }, { h: 'Market cap $bn' }];
    const rows2 = [head(sc)];
    stats.forEach((s, i) => rows2.push(zebra([
      T(s.sector, { bold: true, color: NAVY, align: 'left' }), T(fmt.int(s.n), { align: 'right' }),
      T(fmt.int(s.us), { align: 'right' }), T(fmt.int(s.n - s.us), { align: 'right' }),
      T(fmt.int(s.mq), { align: 'right' }), T(fmt.x(s.pe), { align: 'right' }),
      T(fmt.pct(s.roe), { align: 'right' }), T(fmt.pct(s.om), { align: 'right' }),
      T(fmt.bn(s.mcap), { align: 'right' })], i)));
    const totStyle = { bold: true, color: NAVY, border: [{ type: 'solid', pt: 1.5, color: GOLD },
      { type: 'none' }, { type: 'solid', pt: 0.5, color: LINE }, { type: 'none' }] };
    if (stats.length > 1) rows2.push([
      T('Total', Object.assign({ align: 'left' }, totStyle)),
      T(fmt.int(all.n), Object.assign({ align: 'right' }, totStyle)),
      T(fmt.int(all.us), Object.assign({ align: 'right' }, totStyle)),
      T(fmt.int(all.n - all.us), Object.assign({ align: 'right' }, totStyle)),
      T(fmt.int(all.mq), Object.assign({ align: 'right' }, totStyle)),
      T(fmt.x(all.pe), Object.assign({ align: 'right' }, totStyle)),
      T(fmt.pct(all.roe), Object.assign({ align: 'right' }, totStyle)),
      T(fmt.pct(all.om), Object.assign({ align: 'right' }, totStyle)),
      T(fmt.bn(all.mcap), Object.assign({ align: 'right' }, totStyle))]);
    const fs2 = rows2.length > 11 ? 10 : 11;
    s2.addTable(rows2, tableOpts([2.7, 0.95, 0.85, 0.95, 1.2, 1.15, 1.15, 1.45, 1.53],
      { y: 1.5, fontSize: fs2, rowH: rows2.length > 11 ? 0.33 : 0.38 }));
    const tableBottom = 1.5 + rows2.length * (rows2.length > 11 ? 0.33 : 0.38);
    const held = D.rows.filter(r => r.mq).sort((a, b) => (b.inv || 0) - (a.inv || 0) || (b.mcap || 0) - (a.mcap || 0));
    const room = Math.floor((6.75 - tableBottom - 0.75) / 0.3) - 1;
    if (held.length && room >= 3) {
      const yH = tableBottom + 0.35;
      s2.addText('Most widely held by the superinvestors tracked on Dataroma',
        { x: X0, y: yH, w: CW, h: 0.32, fontFace: HEAD, fontSize: 14, color: NAVY, margin: 0, isTextBox: true });
      const hc = [{ h: 'Ticker', a: 'left' }, { h: 'Company', a: 'left' }, { h: 'Sector', a: 'left' },
                  { h: 'Investors holding' }, { h: 'Mkt cap $bn' }, { h: 'P/E' }, { h: 'ROE' }];
      const hr = [head(hc)];
      held.slice(0, Math.min(5, room)).forEach((r, i) => hr.push(zebra([
        T(r.sym, { bold: true, color: NAVY, align: 'left' }), T(r.name || '', { align: 'left' }),
        T(r.sector || '', { align: 'left' }), T(fmt.int(r.inv), { align: 'right' }),
        T(fmt.bn(r.mcap), { align: 'right' }), T(fmt.x(r.pe), { align: 'right' }),
        T(fmt.pct(r.roe), { align: 'right' })], i)));
      s2.addTable(hr, tableOpts([1.1, 3.9, 2.6, 1.45, 1.2, 0.9, 0.98], { y: yH + 0.42, fontSize: 10, rowH: 0.3 }));
    }
    const tally = { Buy: 0, Hold: 0, Sell: 0 };
    D.rows.forEach(r => { if (tally[r.dOp] !== undefined) tally[r.dOp]++; });
    const scored = tally.Buy + tally.Hold + tally.Sell;
    if (scored) s2.addText(`On the deep analysis scorecard: ${tally.Buy} Buy, ${tally.Hold} Hold, ${tally.Sell} Sell`
      + (scored < D.rows.length ? ` (${D.rows.length - scored} not scored)` : ''),
      { x: X0, y: 1.25, w: CW, h: 0.25, fontFace: BODY, fontSize: 11, color: NAVY, margin: 0, isTextBox: true });
    s2.addNotes(`Settings: ${m.region}; ${m.universe}; ${m.mode}; ${m.ruleBasis}`);

    // 2b. the strongest names on the deep scorecard
    // the strongest names, leaving out those whose figures look implausible or whose score
    // rests on less than 80% of the expected figures; they stay in the sector tables
    const ranked = D.rows.filter(r => num(r.dScore) !== null && !(r.checks || []).length && (r.dCov || 0) >= 80)
      .sort((a, b) => b.dScore - a.dScore || (b.mcap || 0) - (a.mcap || 0));
    const doubtful = D.rows.filter(r => (r.checks || []).length).length;
    const OPC = { Buy: '1E7B34', Hold: GOLD, Sell: 'B42318' };
    if (ranked.length) {
      const top = ranked.slice(0, 10);
      const sT = pres.addSlide({ masterName: 'CONTENT' });
      sT.addText(`The ${top.length} strongest names on the deep analysis scorecard`, { placeholder: 'title' });
      sT.addText('Each figure is scored against its own sector and rolled up into weighted pillars. Pillar scores are out of 100; 50 is the sector median.'
        + (doubtful ? ` ${doubtful} name${doubtful === 1 ? '' : 's'} with implausible figures left out of this ranking.` : ''),
        { x: X0, y: 1.3, w: CW, h: 0.3, fontFace: BODY, fontSize: 10.5, color: NAVY, margin: 0, isTextBox: true });
      const tcT = [{ h: 'Ticker', a: 'left' }, { h: 'Company', a: 'left' }, { h: 'Sector', a: 'left' },
        { h: 'Score' }, { h: 'Opinion', a: 'center' }, { h: 'Quality' }, { h: 'Strength' }, { h: 'Value' },
        { h: 'Momentum' }, { h: 'Its strongest figure', a: 'left' }];
      const rT = [head(tcT)];
      top.forEach((r, i) => rT.push(zebra([
        T(r.sym, { bold: true, color: NAVY, align: 'left', hyperlink: r.link ? { url: r.link } : undefined }),
        T(r.name || '', { align: 'left' }), T(r.sector || '', { align: 'left' }),
        T(fmt.int(r.dScore), { align: 'right', bold: true }),
        T(r.dOp || '-', { align: 'center', bold: true, color: OPC[r.dOp] || NAVY2 }),
        T(fmt.int(r.dQ), { align: 'right' }), T(fmt.int(r.dF), { align: 'right' }),
        T(fmt.int(r.dV), { align: 'right' }), T(fmt.int(r.dM), { align: 'right' }),
        T(r.whyList && r.whyList[0] ? r.whyList[0].txt + ', ' + r.whyList[0].where : '-', { align: 'left' })], i)));
      sT.addTable(rT, tableOpts([0.85, 2.45, 1.75, 0.6, 0.8, 0.7, 0.75, 0.6, 0.85, 2.78],
        { y: 1.75, fontSize: 9.5, rowH: 0.4 }));
    }

    // 3. how names were selected
    const s3 = pres.addSlide({ masterName: 'CONTENT' });
    s3.addText('How the names were selected', { placeholder: 'title' });
    const setRows = [
      ['Sectors', m.sectorsLabel], ['Region', m.region], ['Universe', m.universe],
      ['Threshold set', m.mode], ['Rules applied', m.ruleBasis], ['Missing values', m.blankPolicy],
    ];
    if (m.search) setRows.push(['Search filter', m.search]);
    s3.addTable(setRows.map((r, i) => zebra([T(r[0], { bold: true, color: NAVY }), T(r[1])], i)),
      tableOpts([2.2, CW - 2.2], { y: 1.45, fontSize: 10.5, rowH: 0.32 }));
    // thresholds grid, active metrics only
    const used = m.ruleLabels.filter(l => m.rules.some(r => r.rules[l]));
    if (used.length) {
      const y3 = 1.45 + setRows.length * 0.34 + 0.35;
      const nRules = m.rules.length;
      const fsz = nRules > 8 ? 8.5 : 9.5;
      const firstW = 2.2, other = (CW - firstW) / used.length;
      const grid = [[T('Rule set', { bold: true, color: WHITE, fill: { color: NAVY } })]
        .concat(used.map(l => T(l, { bold: true, color: WHITE, fill: { color: NAVY }, align: 'center' })))];
      m.rules.forEach((r, i) => grid.push(zebra([T(r.sector, { bold: true, color: NAVY })]
        .concat(used.map(l => T(r.rules[l] || 'off', { align: 'center', italic: !r.rules[l],
          color: r.rules[l] ? INK : NAVY2 }))), i)));
      s3.addTable(grid, tableOpts([firstW].concat(used.map(() => other)),
        { y: y3, fontSize: fsz, rowH: Math.min(0.3, (6.85 - y3) / (nRules + 1)) }));
    }

    // 4+. sector tables. One sector: up to three pages. Several: the largest 12 of each.
    const per = 12;
    const pages = groups.length === 1 ? 3 : 1;
    const tc = [{ h: 'Ticker', a: 'left' }, { h: 'Company', a: 'left' }, { h: 'Country', a: 'left' },
      { h: 'Mkt cap $bn' }, { h: 'P/E' }, { h: 'ROE' }, { h: 'Op margin' }, { h: 'Deep score' },
      { h: 'Opinion', a: 'center' }, { h: 'Momentum', a: 'left' }, { h: 'Marquee', a: 'center' }];
    for (const g of groups) {
      for (let p = 0; p < pages; p++) {
        const chunk = g.rows.slice(p * per, (p + 1) * per);
        if (!chunk.length) break;
        const s = pres.addSlide({ masterName: 'CONTENT' });
        const shown = Math.min(g.rows.length, (p + 1) * per);
        const title = p === 0
          ? `${g.sector}: ${g.rows.length.toLocaleString('en-US')} name${g.rows.length === 1 ? '' : 's'} cleared`
          : `${g.sector}, continued`;
        s.addText(title, { placeholder: 'title' });
        const sub = g.rows.length <= per && p === 0 ? 'All names shown, largest first by market cap'
          : `Names ${p * per + 1} to ${shown} of ${g.rows.length.toLocaleString('en-US')}, largest first by market cap`;
        s.addText(sub, { x: X0, y: 1.3, w: CW, h: 0.3, fontFace: BODY, fontSize: 11, color: NAVY,
          margin: 0, isTextBox: true });
        const rows = [head(tc)];
        chunk.forEach((r, i) => rows.push(zebra([
          T(r.sym, { bold: true, color: NAVY, align: 'left' }),
          T(r.name || '', { align: 'left' }), T(r.ctry || '', { align: 'left' }),
          T(fmt.bn(r.mcap), { align: 'right' }), T(fmt.x(r.pe), { align: 'right' }),
          T(fmt.pct(r.roe), { align: 'right' }), T(fmt.pct(r.om), { align: 'right' }),
          T(fmt.int(r.dScore), { align: 'right', bold: true }),
          T(r.dOp && r.dOp !== 'Not scored' ? r.dOp : '-', { align: 'center', bold: true, color: OPC[r.dOp] || NAVY2 }),
          T(r.mom || '-', { align: 'left' }),
          T(r.mq ? 'Held' : '', { align: 'center', bold: true, color: GOLD })], i)));
        s.addTable(rows, tableOpts([0.95, 2.75, 1.4, 0.95, 0.7, 0.75, 0.9, 0.85, 0.85, 1.05, 0.98],
          { y: 1.7, fontSize: 10, rowH: 0.36 }));
        const mqN = g.rows.filter(r => r.mq).length;
        s.addText(`Market cap in US dollars. Returns and margins are trailing twelve months. ` +
          `${mqN} of the ${g.rows.length.toLocaleString('en-US')} names are held by at least one investor tracked by Dataroma.`,
          { x: X0, y: 6.55, w: CW, h: 0.3, fontFace: BODY, fontSize: 9.5, color: NAVY, margin: 0, isTextBox: true });
      }
    }

    // company profiles: one page for every company in the export, by sector, strongest
    // deep score first within each sector
    const prof = [];
    for (const g of groups) prof.push(...g.rows.slice().sort((a, b) =>
      (num(b.dScore) === null ? -1 : b.dScore) - (num(a.dScore) === null ? -1 : a.dScore) || (b.mcap || 0) - (a.mcap || 0)));
    if (prof.length) {
      const sd = pres.addSlide({ masterName: 'COVER' });
      sd.addText('Company profiles', { placeholder: 'title' });
      sd.addText([
        T(`One page for each of the ${prof.length.toLocaleString('en-US')} companies, by sector, strongest deep score first`, { breakLine: true, color: WHITE }),
        T('Why each one passed the screen, what lifts its score, what holds it back, how it compares with its sector median, and where the price stands.', { fontSize: 13 }),
      ], { placeholder: 'body' });
    }
    const RECT = (pres.shapes && pres.shapes.RECTANGLE) || (pres.ShapeType && pres.ShapeType.rect) || 'rect';
    const sign = v => num(v) === null ? '-' : (v > 0 ? '+' : '') + v.toFixed(1) + '%';
    const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const bullets = (list, color) => list.length
      ? list.map((t, i) => T(String(t).replace(/\.$/, '') + '.', { bullet: { indent: 11 }, breakLine: i < list.length - 1,
          color: color || INK, paraSpaceAfter: 3 }))
      : [T('Nothing stands out.', { color: NAVY2, italic: true })];
    const heading = (sl, txt, x, y) => {
      sl.addText(txt, { x, y, w: 4, h: 0.28, fontFace: HEAD, fontSize: 12.5, color: NAVY, margin: 0, isTextBox: true });
      sl.addShape(RECT, { x, y: y + 0.3, w: 0.55, h: 0.03, fill: { color: GOLD }, line: { color: GOLD } });
    };
    prof.forEach(r => {
      const sl = pres.addSlide({ masterName: 'CONTENT' });
      sl.addText(`${r.sym}   ${r.name || ''}`, { placeholder: 'title' });
      if (r.link) sl.addText('Open the full deep analysis', { x: 9.9, y: 0.45, w: 2.83, h: 0.3, fontFace: BODY,
        fontSize: 10, color: NAVY2, align: 'right', margin: 0, isTextBox: true, hyperlink: { url: r.link } });
      const facts = [r.sector, r.ind, r.ctry, num(r.mcap) === null ? null : 'Market cap $' + fmt.bn(r.mcap) + 'bn',
        num(r.px) === null ? null : 'Price ' + r.px.toLocaleString('en-US', { maximumFractionDigits: 2 }) + ' ' + (r.ccy || ''),
        r.inv ? 'Held by ' + fmt.int(r.inv) + ' superinvestor' + (r.inv === 1 ? '' : 's') : null].filter(Boolean);
      sl.addText(facts.join('   |   '), { x: X0, y: 1.28, w: CW, h: 0.26, fontFace: BODY, fontSize: 10,
        color: NAVY, margin: 0, isTextBox: true });

      // top left: verdict and pillars
      const xL = X0, wL = 3.55, opc = OPC[r.dOp] || NAVY2;
      sl.addShape(RECT, { x: xL, y: 1.7, w: wL, h: 0.78, fill: { color: opc }, line: { color: opc } });
      sl.addText([T(r.dOp || 'Not scored', { fontSize: 20, bold: true, color: WHITE, breakLine: true }),
                  T(num(r.dScore) === null ? 'Not in the deep feed' : `Deep score ${fmt.int(r.dScore)} of 100`, { fontSize: 10.5, color: WHITE })],
        { x: xL + 0.15, y: 1.7, w: wL - 0.3, h: 0.78, fontFace: BODY, valign: 'middle', margin: 0 });
      const pill = [['Quality', r.dQ], ['Financial strength', r.dF], ['Valuation', r.dV], ['Momentum', r.dM], ['Signals', r.dG]]
        .filter(([lab, v]) => lab !== 'Signals' || num(v) !== null);
      pill.forEach(([lab, v], i) => {
        const y = 2.6 + i * 0.33;
        sl.addText(lab, { x: xL, y, w: 1.45, h: 0.27, fontFace: BODY, fontSize: 9.5, color: NAVY, margin: 0, valign: 'middle' });
        sl.addShape(RECT, { x: xL + 1.5, y: y + 0.06, w: 1.6, h: 0.15, fill: { color: PALE }, line: { color: PALE } });
        if (num(v) !== null) sl.addShape(RECT, { x: xL + 1.5, y: y + 0.06, w: Math.max(0.02, 1.6 * v / 100), h: 0.15,
          fill: { color: v >= 65 ? NAVY : v >= 45 ? GOLD : 'B42318' }, line: { type: 'none' } });
        sl.addText(num(v) === null ? 'n/a' : String(Math.round(v)), { x: xL + 3.15, y, w: 0.4, h: 0.27,
          fontFace: BODY, fontSize: 9.5, bold: true, color: NAVY, align: 'right', margin: 0, valign: 'middle' });
      });

      // top middle: why it passed
      const xM = 4.45, wM = 4.1;
      heading(sl, 'Why it passed the screen', xM, 1.7);
      sl.addText(bullets((r.whyList || []).map(x => x.txt + ', ' + x.where)),
        { x: xM, y: 2.1, w: wM, h: 1.05, fontFace: BODY, fontSize: 9.5, color: INK, valign: 'top', margin: 0 });
      heading(sl, 'Closest call', xM, 3.2);
      sl.addText((r.closeTxt ? r.closeTxt + '. The figure with the least room above its bar. ' : '') + (r.passed ? r.passed + '.' : ''),
        { x: xM, y: 3.6, w: wM, h: 0.75, fontFace: BODY, fontSize: 9.5, color: INK, valign: 'top', margin: 0 });

      // top right: what the deep scorecard says
      const xR = 8.8, wR = 3.93;
      heading(sl, 'Lifting the score', xR, 1.7);
      sl.addText(bullets((r.dLifts || []).slice(0, 3)), { x: xR, y: 2.1, w: wR, h: 1.05, fontFace: BODY, fontSize: 9.5,
        color: INK, valign: 'top', margin: 0 });
      heading(sl, 'Holding it back', xR, 3.2);
      const hb = (r.checks || []).map(x => 'Check the data: ' + x).concat((r.dFlags || []).map(x => 'Red flag: ' + x), r.dDrags || []).slice(0, 3);
      sl.addText(bullets(hb), { x: xR, y: 3.58, w: wR, h: 0.9, fontFace: BODY, fontSize: 9, color: INK, valign: 'top', margin: 0 });

      // bottom left: against the sector median
      const md = r.med || {};
      const mc = [['P/E', r.pe, md.pe, fmt.x], ['ROIC', r.roic, md.roic, fmt.pct], ['ROE', r.roe, md.roe, fmt.pct],
        ['Gross margin', r.gm, md.gm, fmt.pct], ['Op margin', r.om, md.om, fmt.pct], ['FCF margin', r.fcfm, md.fcfm, fmt.pct],
        ['Debt / equity', r.de, md.de, fmt.z], ['Altman Z', r.az, md.az, fmt.z]];
      const KEYOF = { 'P/E': 'pe_max', ROIC: 'roic_min', ROE: 'roe_min', 'Gross margin': 'gross_margin_min',
        'Op margin': 'op_margin_min', 'FCF margin': 'fcf_margin_min', 'Debt / equity': 'de_max', 'Altman Z': 'altman_min' };
      const uk = r.useKeys || [];
      let use = mc.filter(c => uk.includes(KEYOF[c[0]]));
      if (uk.includes('roa_min') && !uk.includes('roic_min')) use.splice(2, 0, ['ROA', r.roa, md.roa, fmt.pct]);
      if (uk.includes('net_margin_min') && !uk.includes('fcf_margin_min')) use.push(['Net margin', r.nm, md.nm, fmt.pct]);
      if (use.length < 3) use = mc;
      const cw = Math.min(1.0, 5.75 / use.length);
      const hdr = [T('Against its sector', { bold: true, color: WHITE, fill: { color: NAVY } })]
        .concat(use.map(c => T(c[0], { bold: true, color: WHITE, fill: { color: NAVY }, align: 'right' })));
      const rowC = [T(r.sym, { bold: true, color: NAVY })].concat(use.map(c => T(c[3](num(c[1])), { align: 'right', bold: true })));
      const rowM = [T('Sector median', { color: NAVY })].concat(use.map(c => T(c[3](num(c[2])), { align: 'right', color: NAVY })));
      sl.addTable([hdr, rowC, zebra(rowM, 1)], tableOpts([1.35].concat(use.map(() => cw)),
        { x: X0, y: 4.7, fontSize: 9, rowH: 0.3 }));
      sl.addText('Sector median among companies above $300M. For P/E, the median of the sector\'s quality peers (ROIC or ROE of 10% or more).',
        { x: X0, y: 5.68, w: 7.1, h: 0.4, fontFace: BODY, fontSize: 8.5, color: NAVY, margin: 0, isTextBox: true, valign: 'top' });
      // bottom right: recent price performance and where the price sits against its averages
      const xC = 7.95, wC = 4.78;
      heading(sl, 'Price and trend', xC, 4.6);
      const pr = [['1 month', r.m1], ['3 months', r.m3], ['6 months', r.m6], ['YTD', r.ytd], ['1 year', r.y1]];
      sl.addTable([pr.map(c => T(c[0], { bold: true, color: WHITE, fill: { color: NAVY }, align: 'right' })),
        pr.map(c => T(sign(num(c[1])), { align: 'right', bold: true,
          color: num(c[1]) === null ? NAVY2 : c[1] >= 0 ? '1E7B34' : 'B42318' }))],
        tableOpts(pr.map(() => wC / 5), { x: xC, y: 5.0, fontSize: 9.5, rowH: 0.3 }));
      const vs = (a, b) => num(a) === null || num(b) === null || !b ? null : (a / b - 1) * 100;
      const p50 = vs(r.px, r.e50), p200 = vs(r.px, r.e200);
      const lines = [];
      if (p50 !== null) lines.push([`Price vs 50 day EMA`, sign(p50), p50]);
      if (p200 !== null) lines.push([`Price vs 200 day EMA`, sign(p200), p200]);
      if (num(r.e50) !== null && num(r.e200) !== null)
        lines.push(['50 day vs 200 day', r.e50 > r.e200 ? 'above' : 'below', r.e50 > r.e200 ? 1 : -1]);
      if (num(r.rsi) !== null) lines.push(['RSI (14 day)', r.rsi.toFixed(0), r.rsi > 70 ? -1 : r.rsi < 30 ? -1 : 0]);
      lines.push(['Momentum stage', r.mom || 'n/a', 0]);
      sl.addTable(lines.map((l, i) => zebra([T(l[0], { color: NAVY }),
          T(l[1], { align: 'right', bold: true, color: l[2] > 0 ? '1E7B34' : l[2] < 0 ? 'B42318' : INK })], i)),
        tableOpts([wC - 1.3, 1.3], { x: xC, y: 5.72, fontSize: 9, rowH: 0.21 }));
      sl.addNotes(`Deep analysis scorecard from the snapshot figures${num(r.dCov) === null ? '' : ' (' + r.dCov + '% of expected figures available)'}. `
        + 'The full deep analysis page adds price history, cash quality and forensic checks, so its verdict can differ. '
        + (r.link ? 'Full analysis: ' + r.link : ''));
    });

    // last. notes
    const sN = pres.addSlide({ masterName: 'CONTENT' });
    sN.addText('Notes on the data', { placeholder: 'title' });
    const notes = [
      ['Source', `TradingView screener export as at ${m.asOf}. Superinvestor holdings from Dataroma as at ${m.marqueeAsOf}.`],
      ['Sector mapping', 'TradingView sectors are folded into the eleven GICS sectors. REITs are moved from Finance to Real Estate by industry name.'],
      ['Figures', 'P/E, returns and margins are trailing twelve months unless stated. Current ratio and debt to equity are from the latest quarter. Market cap is in US dollars.'],
      ['Marquee held', 'Held by at least one superinvestor portfolio tracked by Dataroma. Ownership data lags the filing date.'],
      ['Deep scorecard', 'Each figure is scored 0 to 100 against its own sector; five weighted pillars (quality 30, financial strength 20, valuation 25, momentum 15, signals 10) give the score. Buy from 70, Hold from 45. Red flags cap the opinion at Hold. Built from the snapshot figures; the full deep analysis page adds price history, cash quality and forensic checks.'],
      ['Use', 'A screen narrows the field for research. It is not a recommendation to buy or sell any security, and every name needs its own work before any decision.'],
    ];
    sN.addTable(notes.map((r, i) => zebra([T(r[0], { bold: true, color: NAVY, valign: 'top' }),
      T(r[1], { valign: 'top' })], i)), tableOpts([2.2, CW - 2.2], { y: 1.5, fontSize: 11.5, rowH: 0.6 }));
    return pres;
  }

  return { workbook, deck, bySector, sectorStats };
})();
/*EXPORT_END*/
/*__BOOT__*/
</script></body></html>
"""




def _gh(method, url, token, payload=None):
    """One GitHub API call. Returns the parsed JSON body."""
    import urllib.error, urllib.request
    data = json.dumps(payload).encode() if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        'Authorization': f'Bearer {token}',
        'Accept': 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28',
        'Content-Type': 'application/json',
        'User-Agent': 'marquee-screener',
    })
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read() or b'{}')
    except urllib.error.HTTPError as e:
        body = e.read().decode(errors='replace')[:400]
        hint = {
            401: 'the token is wrong, expired, or was never set',
            403: 'the token lacks Contents: Read and write on this repository',
            404: 'repo, branch or path not found, or the token cannot see this repo',
            409: 'the branch moved while uploading - just run again',
            422: 'GitHub rejected the content, often a stale sha - just run again',
        }.get(e.code, 'see the message above')
        sys.exit(f'\n[ERROR] GitHub said {e.code} on {method} {url}\n'
                 f'        Likely cause: {hint}\n        Response: {body}\n')
    except urllib.error.URLError as e:
        sys.exit(f'\n[ERROR] Could not reach GitHub: {e.reason}\n'
                 f'        Check the internet connection or a corporate proxy.\n')


def publish_to_github(files, stamp):
    """Commit several files to a GitHub repo in one commit, using the Git Data API.

    files is a list of (local path, path inside the repo).
    blob -> tree -> commit -> ref is how git itself works, so a 6 MB file is no problem.
    The simpler Contents API is unreliable above about 1 MB.
    """
    import base64
    token = (GITHUB_TOKEN or os.environ.get('GITHUB_TOKEN', '')).strip()
    if not token or 'PASTE_YOUR' in token:
        print('      [WARN] no token found, upload skipped. Fill GITHUB_TOKEN or set the '
              'GITHUB_TOKEN environment variable.')
        return
    files = [(lp, rp) for lp, rp in files if rp and os.path.exists(lp)]
    if not files:
        print('      [WARN] nothing to upload')
        return

    api = f'https://api.github.com/repos/{GITHUB_REPO}'
    total = sum(os.path.getsize(lp) for lp, _ in files) / 1024 / 1024
    names = ', '.join(rp for _, rp in files)
    print(f'      uploading {total:,.1f} MB to {GITHUB_REPO} ({GITHUB_BRANCH}): {names}')

    ref = _gh('GET', f'{api}/git/ref/heads/{GITHUB_BRANCH}', token)
    parent = ref['object']['sha']
    base_tree = _gh('GET', f'{api}/git/commits/{parent}', token)['tree']['sha']

    tree_items = []
    for local, repo_path in files:
        with open(local, 'rb') as fh:
            content = base64.b64encode(fh.read()).decode()
        sha = _gh('POST', f'{api}/git/blobs', token,
                  {'content': content, 'encoding': 'base64'})['sha']
        tree_items.append({'path': repo_path, 'mode': '100644', 'type': 'blob', 'sha': sha})
        print(f'        uploaded {repo_path}')

    tree = _gh('POST', f'{api}/git/trees', token,
               {'base_tree': base_tree, 'tree': tree_items})['sha']
    commit = _gh('POST', f'{api}/git/commits', token, {
        'message': f"Screener data as of {stamp['data']}", 'tree': tree,
        'parents': [parent]})['sha']
    _gh('PATCH', f'{api}/git/refs/heads/{GITHUB_BRANCH}', token, {'sha': commit})

    owner, repo = GITHUB_REPO.split('/')
    print(f'      committed {commit[:7]} - live in a minute or two at '
          f'https://{owner}.github.io/{repo}/')
    if GITHUB_XLSX:
        print(f'      workbook -> https://github.com/{GITHUB_REPO}/raw/{GITHUB_BRANCH}/{GITHUB_XLSX}')
    print('      if the page looks unchanged, hard refresh with Ctrl+Shift+R')


# ------------------------------------------------------------------ main -------------
def main():
    ap = argparse.ArgumentParser(description='Marquee portfolio and GICS sector screener builder')
    ap.add_argument('--data-dir', default=DATA_DIR, help='folder with the two CSVs')
    ap.add_argument('--out-dir', default=OUT_DIR, help='where to write the workbook and HTML')
    ap.add_argument('--no-html', action='store_true', help='skip the interactive HTML')
    ap.add_argument('--publish-dir', default=None,
                    help='also write index.html here, e.g. a checked-out website repo')
    ap.add_argument('--scrape', choices=['never', 'auto', 'always'], default=None,
                    help='override SCRAPE_DATAROMA')
    ap.add_argument('--no-upload', action='store_true',
                    help='skip the GitHub API upload, e.g. when a CI job commits instead')
    ap.add_argument('--stable-names', action='store_true',
                    help='drop the date from the output filenames')
    ap.add_argument('--web', action='store_true',
                    help='write a light index.html plus data.json instead of one big file')

    # Jupyter, Spyder and VS Code notebooks inject their own arguments such as
    # -f kernel-xxxx.json, which argparse would reject. Inside a notebook the command
    # line is ignored entirely and the DATA_DIR / OUT_DIR values above are used.
    in_notebook = ('ipykernel' in sys.modules or 'spyder_kernels' in sys.modules
                   or any(a.endswith('.json') for a in sys.argv[1:]))
    args = ap.parse_args([]) if in_notebook else ap.parse_known_args()[0]

    # command line beats the constants at the top, which is how the CI workflow drives
    # this script without editing the file
    global SCRAPE_DATAROMA, PUBLISH_DIR, GITHUB_UPLOAD, DATE_IN_FILENAME
    if args.scrape:
        SCRAPE_DATAROMA = args.scrape
    if args.publish_dir is not None:
        PUBLISH_DIR = args.publish_dir
    if args.stable_names:
        DATE_IN_FILENAME = False
    if args.no_upload or os.environ.get('GITHUB_ACTIONS') == 'true':
        # inside GitHub Actions the workflow commits the files itself, so the API
        # uploader is never needed and a personal token is never required
        GITHUB_UPLOAD = False
    if in_notebook:
        print('[note] notebook detected, command line ignored - using DATA_DIR and OUT_DIR '
              'from the top of the script')

    os.makedirs(args.out_dir, exist_ok=True)
    print(f'[1/6] data folder : {os.path.abspath(args.data_dir)}')
    if not os.path.isdir(args.data_dir):
        sys.exit(f'\n[ERROR] The folder does not exist:\n        '
                 f'{os.path.abspath(args.data_dir)}\n        Fix: correct DATA_DIR.\n')
    do_scrape, why = dataroma_needs_scrape(args.data_dir)
    if do_scrape:
        print(f'      scraping dataroma.com ({why})')
        have_old = any(classify(x) == 'marquee'
                       for x in glob.glob(os.path.join(args.data_dir, '*.csv')))
        try:
            scrape_dataroma(args.data_dir)
        except SystemExit as e:
            if not have_old:
                raise
            print(f'      [WARN] scrape failed, carrying on with the existing CSV')
            print(f'             reason: {str(e).strip().splitlines()[-1] if str(e) else e}')
    elif why:
        print(f'      dataroma scrape skipped - {why}')
    p_all, p_dr = find_files(args.data_dir)
    prune_old_marquee(args.data_dir, p_dr)
    d_all, d_dr = file_date(p_all), file_date(p_dr)
    stamp = {'data': pretty_date(d_all), 'marquee': pretty_date(d_dr),
             'iso': d_all.isoformat(),
             'built': pretty_date(datetime.date.today()),
             'builtAt': datetime.datetime.now(datetime.timezone.utc).strftime('%d %b %Y, %H:%M UTC')}
    print(f'      data as of  : {stamp["data"]}   (marquee list {stamp["marquee"]})')
    print(f'      universe    : {os.path.basename(p_all)}')
    print(f'      marquee     : {os.path.basename(p_dr)}')
    print(f'      output to   : {os.path.abspath(args.out_dir)}')

    a = load_universe(p_all, feed_dirs=[PUBLISH_DIR, args.out_dir])
    d = load_marquee(p_dr)
    print(f'[2/6] loaded {len(a)} universe rows, {len(d)} unique marquee rows')

    merged = match_marquee(a, d)
    matched = merged[merged['Matched Symbol'].notna()].copy()
    unid = merged[merged['Matched Symbol'].isna()].copy()
    matched['GICS Sector'] = assign_gics(pd.DataFrame({
        'Sector': matched['A_Sector'], 'Industry': matched['A_Industry']}))
    matched = matched.sort_values('Portfolio %', ascending=False).reset_index(drop=True)
    unid = unid.sort_values('Portfolio %', ascending=False).reset_index(drop=True)
    print(f'[3/6] matched {len(matched)}, unidentified {len(unid)}')

    wb = Workbook()
    sheet_readme(wb, len(matched), len(unid), len(a), {}, stamp)
    _, pos = sheet_marquee(wb, matched)
    sheet_unidentified(wb, unid)
    sheet_sector_summary(wb, matched, pos)
    a2 = sheet_all_stocks(wb, a, matched)

    # marquee ownership carried onto the screener rows
    wt = matched.set_index('Matched Symbol')['Portfolio %'].to_dict()
    cnt = matched.set_index('Matched Symbol')['Ownership Count'].to_dict()
    a2['Marquee Wt %'] = a2['Symbol'].map(wt)
    a2['Investors Holding'] = a2['Symbol'].map(cnt)

    compute_tuned(a2)
    pool = a2[a2['Marquee Flag'] == 'Marquee holding'] if SCREEN_UNIVERSE == 'marquee' else a2
    label = 'marquee holdings only' if SCREEN_UNIVERSE == 'marquee' else 'full universe'
    unmapped = int(pool['GICS Sector'].isna().sum())
    print(f"[4/6] screening {len(pool)} rows ({label}), THRESHOLD_MODE = '{THRESHOLD_MODE}'")
    if unmapped:
        print(f'      note: {unmapped} rows carry a TradingView sector with no GICS equivalent '
              f'(Miscellaneous / Government). They can pass All Sectors but reach no sector sheet.')

    def rule_line(applied):
        return '   |   '.join(f'{lbl} {op} {val}' for lbl, op, val in applied)

    counts = {}
    screened_all = set()
    for gics in screen_order():
        if gics == ALL_LABEL and ALL_SECTORS_USES_SECTOR_RULES:
            # each row judged on its own sector's block, then combined
            parts = []
            for g2 in SECTOR_MAP:
                p2, _, _ = screen(pool[pool['GICS Sector'] == g2], g2)
                parts.append(p2)
            unmapped = pool[pool['GICS Sector'].isna()]
            p3, applied, _ = screen(unmapped, ALL_LABEL)
            parts.append(p3)
            passed = pd.concat(parts) if parts else pool.iloc[0:0]
            sub = pool
            rule_text = ('each row judged against its own GICS sector block, then combined '
                         '(so this sheet is the union of the eleven sector sheets). Rows with no '
                         "GICS sector use the 'All Sectors' block: " + rule_line(applied))
        else:
            sub = pool if gics == ALL_LABEL else pool[pool['GICS Sector'] == gics]
            passed, applied, _ = screen(sub, gics)
            rule_text = rule_line(applied)
        counts[gics] = len(passed)
        screened_all.update(passed['Symbol'])
        sheet_sector_screen(wb, gics, passed, rule_text, label, stamp)
        us = int((passed['Country or region of registration'] == 'United States').sum())
        mq = int((passed['Marquee Flag'] == 'Marquee holding').sum())
        flag = '   <-- nothing passed, loosen the thresholds' if len(passed) == 0 else ''
        note = ''
        if gics == ALL_LABEL and ALL_SECTORS_USES_SECTOR_RULES:
            note = '   = union of the 11 sector screens + unmapped rows'
        print(f'      {gics:<24} pool {len(sub):>5}   pass {len(passed):>4} '
              f'(US {us}, non-US {len(passed) - us} | marquee {mq}, other {len(passed) - mq})'
              f'{flag}{note}')
    sheet_screen_settings(wb, counts, stamp)
    write_deep_feed(a2, screened_all, args.out_dir, PUBLISH_DIR, stamp)

    # refresh the Read Me pass-count line now that screens have run
    ws = wb['Read Me']
    for row in ws.iter_rows(min_col=2, max_col=3):
        if row[0].value == 'Passing counts this run':
            row[1].value = ', '.join(f'{g} {n}' for g, n in counts.items())

    wb.properties.title = f"Marquee Investor Portfolio and GICS Sector Screener - {stamp['data']}"
    wb.properties.creator = 'marquee_screener.py'
    suffix = f"_{stamp['iso']}" if DATE_IN_FILENAME else ''
    xlsx = os.path.join(args.out_dir, f'Marquee_Investor_Portfolio_Consolidated{suffix}.xlsx')
    wb.save(xlsx)
    print(f'[5/6] workbook saved -> {xlsx}')

    if not args.no_html:
        web = WEB_SPLIT or args.web
        html = os.path.join(args.out_dir,
                            'index.html' if web else f'Sector_Screener{suffix}.html')
        n_html, n_data = build_html(a2, html, stamp, split=web)
        print(f'[6/6] screener saved -> {html}  ({n_html/1024:,.0f} KB)')
        if n_data:
            print(f'      data file      -> {os.path.join(args.out_dir, "data.json")}  '
                  f'({n_data/1024:,.0f} KB)   keep both files together')
        if PUBLISH_DIR:
            if not os.path.isdir(PUBLISH_DIR):
                print(f'      [WARN] PUBLISH_DIR does not exist, skipped: {PUBLISH_DIR}')
            else:
                import shutil
                dest = os.path.join(PUBLISH_DIR, 'index.html')
                shutil.copyfile(html, dest)
                print(f'      published copy -> {dest}')
                if n_data:
                    shutil.copyfile(os.path.join(args.out_dir, 'data.json'),
                                    os.path.join(PUBLISH_DIR, 'data.json'))
                    print(f'      published copy -> {os.path.join(PUBLISH_DIR, "data.json")}')
                print('      next: cd into that folder, then')
                print(f'        git add index.html && git commit -m "screener {stamp["iso"]}" && git push')
        if GITHUB_UPLOAD:
            payload = [(html, GITHUB_PATH), (xlsx, GITHUB_XLSX)]
            payload += [(os.path.join(args.out_dir, f), f) for f in GITHUB_FEED]
            publish_to_github(payload, stamp)
    else:
        print('[6/6] HTML skipped')
    print('\nDone. The "Screen Settings" sheet lists the final thresholds for every sector.')


if __name__ == '__main__':
    main()
