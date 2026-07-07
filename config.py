# Tuneable constants — change here, no need to touch source files.

CACHE_TTL_SECONDS           = 300   # how long loaded Excel data stays cached
SCRAPE_PAGE_TIMEOUT_SECONDS = 35    # per-country Selenium wait for page render
SCRAPE_MAX_RETRIES          = 3     # attempts per country before marking as failed
SCRAPE_RETRY_DELAY_SECONDS  = 3     # pause between retry attempts
