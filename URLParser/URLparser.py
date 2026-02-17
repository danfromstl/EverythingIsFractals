from __future__ import annotations

import csv
import re
from pathlib import Path

CSV_CANDIDATES = (
    "DarknetDiaries-CVS_Export.csv",  # Original filename in this repo.
    "DarknetDiaries-CSV_Export.csv",  # Common corrected variant.
)

URL_PATTERN = re.compile(
    r"(?i)\b(?:https?://)?(?:www\.)?[a-z0-9.-]+\.(?:com|org|tv)\b[^\s,;\"')]*"
)
TWITTER_PATTERN = re.compile(
    r"(?i)\b(?:https?://)?(?:www\.)?(?:twitter\.com|x\.com)/[a-z0-9_]{1,15}\b"
)


def find_csv_path() -> Path:
    base_dir = Path(__file__).resolve().parent
    for name in CSV_CANDIDATES:
        candidate = base_dir / name
        if candidate.exists():
            return candidate
    raise FileNotFoundError(
        f"Could not find an input CSV in {base_dir}. Tried: {', '.join(CSV_CANDIDATES)}"
    )


def main() -> None:
    csv_path = find_csv_path()

    episode_count = 0
    global_url_count = 0
    twitter_handle_count = 0

    with csv_path.open(encoding="latin1", newline="") as csvfile:
        episode_reader = csv.DictReader(csvfile)

        for row in episode_reader:
            ep_title = str(row.get("Title", "")).strip()
            ep_desc = str(row.get("Description", ""))

            current_urls = URL_PATTERN.findall(ep_desc)
            url_count = len(current_urls)

            if url_count > 0:
                print(f"Found {url_count} urls in {ep_title}")
                for ep_url in current_urls:
                    if TWITTER_PATTERN.search(ep_url):
                        print(ep_url)
                        twitter_handle_count += 1
                print("")

            episode_count += 1
            global_url_count += url_count

    print("\n------\nEpisode list - first pass complete")
    print(f"Read {episode_count} episodes from {csv_path.name}")
    print(f"Found {global_url_count} URLs")
    print(f"Found {twitter_handle_count} Twitter handles")


if __name__ == "__main__":
    main()
