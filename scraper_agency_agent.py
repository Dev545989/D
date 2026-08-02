import os
import random
import time

import pandas as pd
from camoufox.sync_api import Camoufox

BASE_URLS = {
    "agency": "https://uae.dubizzle.com/property-agencies/{slug}/",
    "agent": "https://uae.dubizzle.com/property-agents/{slug}/",
}

BUTTON_SELECTORS = [
    '[data-testid="profile-call-button"]',
    'button:has-text("Call")',
    '[data-testid="call-cta-button"]',
    'button:has-text("Show Phone Number")',
    'button:has-text("Show Number")',
    '[data-testid*="phone" i]',
    '[data-testid*="call" i]',
]

CHALLENGE_MARKERS = [
    "Pardon Our Interruption",
    "Additional security check is required",
    "I am human",
    "hCaptcha",
    "reeseSkipExpirationCheck",
]


def _is_challenge_page(html: str) -> bool:
    return any(marker in html for marker in CHALLENGE_MARKERS)


def _safe_content(page, retries=3, delay=1500):
    for attempt in range(retries):
        try:
            return page.content()
        except Exception:
            if attempt == retries - 1:
                raise
            page.wait_for_timeout(delay)
    return ""


PHONE_KEY_HINTS = ["phone", "didnumber"]


def _find_phone_recursive(obj):
    """
    Generic recursive search for any key whose name looks phone-related
    (contains 'phone', or is 'didNumber' -- the agent profile endpoint
    returns the real number under didNumber and leaves phoneNumber null)
    and whose value is a non-empty string/int.
    """
    if isinstance(obj, dict):
        for key, value in obj.items():
            key_lower = key.lower()
            if any(hint in key_lower for hint in PHONE_KEY_HINTS) and isinstance(value, (str, int)) and value:
                return str(value)
        for value in obj.values():
            found = _find_phone_recursive(value)
            if found:
                return found
    elif isinstance(obj, list):
        for item in obj:
            found = _find_phone_recursive(item)
            if found:
                return found
    return None


def _reveal_phone(page, timeout_ms=10000, debug=False):
    captured = {"data": None}

    def handle_response(response):
        if "graphql" not in response.url:
            return
        try:
            post_data = response.request.post_data or ""
        except Exception:
            post_data = ""
        if "phone" in post_data.lower() and response.status == 200:
            try:
                captured["data"] = response.json()
            except Exception:
                pass

    page.on("response", handle_response)

    button = None
    for selector in BUTTON_SELECTORS:
        loc = page.locator(selector).first
        try:
            if loc.is_visible(timeout=3500):
                button = loc
                break
        except Exception:
            continue

    if button is None:
        page.remove_listener("response", handle_response)
        return None, "button_not_found"

    try:
        button.scroll_into_view_if_needed()
        page.wait_for_timeout(300)
        try:
            button.click(timeout=6000)
        except Exception:
            button.click(force=True)

        waited = 0
        while captured["data"] is None and waited < timeout_ms:
            page.wait_for_timeout(400)
            waited += 400
    except Exception as e:
        page.remove_listener("response", handle_response)
        return None, f"click_error: {e}"
    finally:
        page.remove_listener("response", handle_response)

    if captured["data"] is None:
        return None, "no_response_captured"

    phone = _find_phone_recursive(captured["data"])

    if phone is None:
        if debug:
            # Dump the raw response once so you can see the real key path
            # and hardcode it back in if you want, instead of the generic
            # recursive search.
            import json
            print("    [DEBUG] captured data but no 'phone' key found. Raw response:")
            print(json.dumps(captured["data"], ensure_ascii=False, indent=2)[:2000])
        return None, "phone_field_not_found"

    return phone, "ok"


def enrich_profiles_with_phone(
    df: pd.DataFrame,
    slug_column: str = "slug",
    profile_type: str = "agency",  # "agency" or "agent"
    headless: bool = True,
    min_delay: float = 10,
    max_delay: float = 20,
    save_every: int = 25,
    checkpoint_path: str = "profiles_with_phone_checkpoint.xlsx",
    resume: bool = True,
    max_new: int = None,
    debug: bool = False,
) -> pd.DataFrame:
    df = df.copy()
    base_url = BASE_URLS[profile_type]

    already_done = {}  # slug -> (phone, status)
    if resume and os.path.exists(checkpoint_path):
        try:
            prev_df = pd.read_excel(checkpoint_path)
            for _, prow in prev_df.iterrows():
                slug = prow.get(slug_column)
                status = prow.get("_scrape_status")
                if slug and pd.notna(status) and status not in ("imperva_challenge", "button_not_found", "phone_field_not_found"):
                    already_done[slug] = (prow.get("contact_phone_number"), status)
            print(f"Resuming: found {len(already_done)} already-processed {profile_type} profiles in checkpoint.")
        except Exception as e:
            print(f"Could not read checkpoint for resume ({e}), starting fresh.")

    phones = [None] * len(df)
    statuses = [None] * len(df)

    for pos, (idx, row) in enumerate(df.iterrows()):
        slug = row.get(slug_column)
        if slug in already_done:
            phones[pos], statuses[pos] = already_done[slug]

    def save_checkpoint():
        partial_df = df.copy()
        partial_df["contact_phone_number"] = phones
        partial_df["_scrape_status"] = statuses
        partial_df.to_excel(checkpoint_path, index=False)

    new_processed_count = 0
    consecutive_soft_fails = 0
    SOFT_FAIL_THRESHOLD = 3

    with Camoufox(
        headless=headless,
        humanize=True,
        geoip=True,
        block_images=False
    ) as browser:
        page = browser.new_page()

        for pos, (idx, row) in enumerate(df.iterrows()):
            slug = row.get(slug_column)

            if slug in already_done:
                continue

            if max_new is not None and new_processed_count >= max_new:
                print(f"Reached max_new limit ({max_new}), stopping this run.")
                break

            if not slug:
                statuses[pos] = "no_slug"
                print(f"[{pos + 1}/{len(df)}] Skipped - no slug")
                continue

            url = base_url.format(slug=slug)
            print(f"[{pos + 1}/{len(df)}] {url}")

            attempt_result = None
            for attempt_num in range(2):
                try:
                    page.goto(url, wait_until="domcontentloaded", timeout=45000)
                    page.wait_for_timeout(random.uniform(10000, 15000))

                    html = _safe_content(page)
                    if _is_challenge_page(html):
                        print("Challenge detected")

                        page.screenshot(path="imperva.png", full_page=True)
                        with open("imperva.html", "w", encoding="utf-8") as f:
                            f.write(html)

                        print(page.url)
                        print(page.title())

                        input("Press Enter to close...")

                        statuses[pos] = "imperva_challenge"
                        save_checkpoint()
                        attempt_result = "imperva_break"
                        break

                    phone, status = _reveal_phone(page, debug=debug)
                    phones[pos] = phone
                    statuses[pos] = status
                    print(f"  -> phone: {phone} (status: {status})")

                    if status in ("button_not_found", "phone_field_not_found"):
                        consecutive_soft_fails += 1
                        try:
                            safe_slug = slug.replace("/", "_")
                            page.screenshot(
                                path=f"debug_no_phone_{safe_slug}.png", full_page=True
                            )
                            page_text = page.locator("body").inner_text()[:500]
                            print(f"     [DEBUG] Page title: {page.title()}")
                            print(f"     [DEBUG] Body text snippet: {page_text[:200]!r}")
                        except Exception as e:
                            print(f"     [DEBUG] Screenshot failed: {e}")
                    else:
                        consecutive_soft_fails = 0

                    attempt_result = "done"
                    break

                except Exception as e:
                    if attempt_num == 0:
                        print(f"  -> Transient error, retrying once: {e}")
                        page.wait_for_timeout(2000)
                        continue
                    statuses[pos] = f"error: {e}"
                    print(f"  -> FAILED after retry: {e}")
                    attempt_result = "failed"

            if attempt_result == "imperva_break":
                break

            if consecutive_soft_fails >= SOFT_FAIL_THRESHOLD:
                print(
                    f"\n\u26a0\ufe0f  {consecutive_soft_fails} consecutive soft-fail results - "
                    "this looks like a soft rate-limit/reputation warning from Imperva, "
                    "not real missing buttons. Stopping this run early to avoid a full block."
                )
                save_checkpoint()
                break

            new_processed_count += 1

            if new_processed_count % save_every == 0:
                save_checkpoint()
                print(f"  [Checkpoint saved - {new_processed_count} new rows this run]")

            if pos < len(df) - 1:
                time.sleep(random.uniform(min_delay, max_delay))

        page.close()

    save_checkpoint()
    print(f"\nThis run processed {new_processed_count} new {profile_type} profiles.")

    df["contact_phone_number"] = phones
    df["_scrape_status"] = statuses
    return df


if __name__ == "__main__":
    import argparse
    from pathlib import Path

    parser = argparse.ArgumentParser()

    parser.add_argument("--start", type=int, default=0)
    parser.add_argument("--end", type=int, default=2)
    parser.add_argument("--input", type=str, default="all_property_agents_0_100.csv")
    parser.add_argument("--profile-type", type=str, default="agent", choices=["agency", "agent"])
    parser.add_argument("--debug", action="store_true", help="Dump raw response JSON when no phone key is found")

    args = parser.parse_args()

    ext = Path(args.input).suffix.lower()

    if ext == ".xlsx":
        profiles_df = pd.read_excel(args.input)
    else:
        profiles_df = pd.read_csv(args.input)

    result_df = enrich_profiles_with_phone(
        profiles_df,
        profile_type=args.profile_type,
        max_new=100,
        checkpoint_path=f"{args.profile_type}_checkpoint.xlsx",
        debug=args.debug,
    )

    result_df.to_csv(
        f"{args.profile_type}_with_phone_{args.start}_{args.end}.csv",
        index=False,
        encoding="utf-8-sig",
    )

    total = len(result_df)
    with_phone = result_df["contact_phone_number"].notna().sum()
    attempted = result_df["_scrape_status"].notna().sum()
    print(f"\nProgress so far: {attempted}/{total} attempted, {with_phone} have a phone number")