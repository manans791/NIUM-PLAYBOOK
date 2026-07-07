"""Standalone scrape runner — use instead of the Streamlit UI for testing."""
import sys
sys.path.insert(0, r"C:\Automations\Playbook APp")

from scraper import scrape_dataset, transform_raw_to_wide, save_scraped_data, log_scrape_failures

PREVIOUSLY_MISSING = [
    "Bonaire, Sint Eustatius and Saba", "Burundi", "Cayman Islands", "Chad",
    "Cocos (Keeling) Islands", "Croatia (Hrvatska)", "Equatorial Guinea",
    "Eritrea", "Falkland Islands (Malvinas)", "Kazakhstan",
    "Micronesia, Federated States of", "New Zealand", "Nicaragua",
    "Saint Martin (French part)", "Sint Maarten (Dutch part)", "Virgin Islands (U.S.)",
]

class MockProgress:
    def progress(self, pct, text=""): print(f"\r  {text}", end="", flush=True)

class MockStatus:
    def caption(self, text): pass

all_results = {}
for ds_type in ["FI", "Non-FI"]:
    print(f"\n{'='*60}")
    print(f"Scraping {ds_type} dataset...")
    raw_rows, failed = scrape_dataset(ds_type, MockProgress(), MockStatus())
    print(f"\n  Raw rows: {len(raw_rows):,}  |  Failed: {len(failed)}")
    if failed:
        print(f"  Failed countries: {sorted(failed)}")

    df = transform_raw_to_wide(raw_rows)
    if not df.empty:
        app_path, archive = save_scraped_data(df, ds_type)
        print(f"  Saved: {app_path}")
        print(f"  Countries in data: {df['Country'].nunique()}")
    all_results[ds_type] = {"df": df, "failed": failed}

print(f"\n{'='*60}")
print("PREVIOUSLY MISSING — now present?")
for ds_type, info in all_results.items():
    df = info["df"]
    failed = info["failed"]
    print(f"\n  [{ds_type}]")
    for c in PREVIOUSLY_MISSING:
        if df.empty:
            status = "NO DATA"
        elif c in df["Country"].values:
            status = "RECOVERED"
        elif c in failed:
            status = "still FAILED"
        else:
            status = "missing (no rows)"
        print(f"    {status:20s}  {c}")

log_scrape_failures({ds: info["failed"] for ds, info in all_results.items()})
print("\nDone.")
