import pandas as pd
import re
from collections import defaultdict
from copy import deepcopy
import os
import json
import logging
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ===============================
# 🔧 CONFIGURATION SECTION
# ===============================

LOG_DIR = "output"
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

def setup_logging():
    """Configures timestamped logging."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = os.path.join(LOG_DIR, f"merger_{timestamp}.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(log_filename),
            logging.StreamHandler(),
        ],
    )
    return timestamp

# Call setup_logging() at the start
TIMESTAMP = setup_logging()

# Normalized group name map (Arabic → English)
GROUP_MAP = {
    "* myContacts": "* myContacts",
    "lab ::: * myContacts": "🧪 Lab ::: * myContacts",
    "شخصي ::: * myContacts": "🏠 Personal ::: * myContacts",
    "* family ::: * myContacts": "👨‍👩‍👧‍👦 Family ::: * myContacts",
    "شركات ومندوبين ::: * myContacts": "🏢 Companies & Agents ::: * myContacts",
    "اطباء ::: * myContacts": "🧑‍⚕️ Doctors ::: * myContacts",
    "وظائف ::: * myContacts": "💼 Jobs ::: * myContacts",
}

# Default group for new MSSQL contacts
DEFAULT_NEW_GROUP = "🧪 Lab ::: * myContacts"

# Default country code
DEFAULT_COUNTRY_CODE = "+20"

# Google API Settings
SCOPES = ["https://www.googleapis.com/auth/contacts.readonly"]
CLIENT_SECRET_FILE = "client_secret.json"
TOKEN_FILE = "token.json"


# ===============================
# 🧠 UTILITIES
# ===============================


def safe_read_csv(path):
    """Read CSV with automatic encoding detection."""
    encodings_to_try = ["utf-8", "utf-8-sig", "utf-16", "cp1252", "latin1"]
    for enc in encodings_to_try:
        try:
            df = pd.read_csv(
                path, dtype=str, encoding=enc, on_bad_lines="skip", engine="python"
            ).fillna("")
            df.columns = [
                c.strip().replace("\ufeff", "").replace("ÿþ", "") for c in df.columns
            ]
            logging.info(f"Successfully read CSV '{path}' with encoding '{enc}'.")
            return df
        except Exception as e:
            logging.warning(
                f"Could not read CSV '{path}' with encoding '{enc}': {e}"
            )
            continue
    raise ValueError(f"Cannot read file {path} with tried encodings.")


def clean_phone(phone: str) -> str:
    return re.sub(r"[^\d+]", "", str(phone)).strip()


def normalize_phone(num: str) -> str | None:
    if num is None:
        return None
    s = str(num).strip()
    if not s or s.lower() == "null":
        return None
    num = clean_phone(s)
    if num.startswith("+"):
        return num
    if num.startswith("00"):
        return "+" + num[2:]
    if re.match(r"^01[0-9]{8,}$", num):
        return "+2" + num
    if re.match(r"^05[0-9]{8,}$", num):
        return "+966" + num[1:]
    if not num.startswith("+"):
        digits_only = re.sub(r"\D", "", num).lstrip("0")
        if not digits_only:
            return None
        cc = (
            DEFAULT_COUNTRY_CODE
            if DEFAULT_COUNTRY_CODE.startswith("+")
            else "+" + DEFAULT_COUNTRY_CODE
        )
        return cc + digits_only
    return num


def normalize_group_name(group_value: str) -> str:
    group_value = (group_value or "").strip().replace("::: * starred", "").strip()
    for old, new in GROUP_MAP.items():
        if old in group_value:
            group_value = group_value.replace(old, new)
    return group_value


def strip_lab_token(name: str) -> str:
    s = re.sub(r"(?i)\blab\b", "", str(name or ""))
    return re.sub(r"\s+", " ", s).strip()


def normalize_display_name(
    raw: str, append_lab: bool = True, preserve_lab: bool = False
) -> str:
    if not raw:
        return " Lab" if append_lab else ""
    s = str(raw)
    if not preserve_lab:
        s = re.sub(r"(?i)\blab\b", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    if append_lab and not s.endswith(" Lab"):
        s = (s + " Lab").strip()
    return s


def expand_normalize_numbers(seq):
    out = []
    for v in seq or []:
        if v is None:
            continue
        s = str(v)
        parts = s.split(":::") if ":::" in s else [s]
        for p in parts:
            n = normalize_phone(p.strip())
            if n and n not in out:
                out.append(n)
    return out


# ===============================
# 📥 LOAD CONTACTS
# ===============================


def _process_google_df(df):
    """
    Processes a DataFrame of Google contacts (from either CSV or API)
    and returns the standard contacts dictionary. This is the single source of truth.
    """
    contacts = {}
    for _, row in df.iterrows():
        first = (row.get("First Name") or "").strip()
        middle = (row.get("Middle Name") or "").strip()
        last = (row.get("Last Name") or "").strip()

        raw_name = (row.get("Name") or "").strip()
        if not raw_name and (first or middle or last):
            raw_name = " ".join([p for p in [first, middle, last] if p]).strip()
        if not raw_name:
            continue

        groups_raw = row.get("Labels") or row.get("Group Membership") or "* myContacts"

        def _is_lab_group(g: str) -> bool:
            tokens = [t.strip().lower() for t in (g or "").split(":::")]
            for t in tokens:
                t_clean = re.sub(r"[^\w\s]", "", t)
                if "lab" in t_clean:
                    return True
            return False

        is_lab = _is_lab_group(groups_raw)
        groups = normalize_group_name(groups_raw)
        name = normalize_display_name(raw_name, append_lab=is_lab, preserve_lab=True)
        cmp_name = strip_lab_token(name).lower()

        phone_cols = [f"Phone {i} - Value" for i in range(1, 10)]
        existing_cells = [
            row.get(c, "") for c in phone_cols if c in row and pd.notna(row.get(c))
        ]
        numbers_list = expand_normalize_numbers(existing_cells)

        contacts[name] = {
            "numbers": set(numbers_list),
            "groups": {groups},
            "protected": not is_lab,
            "sources": {"Google"},
            "_cmp_name": cmp_name,
            "first_name": first,
            "middle_name": middle,
            "last_name": last,
            "_raw_row": {c: (row.get(c) or "") for c in df.columns},
        }
    return contacts


def load_google_contacts(path):
    """Loads Google contacts from a CSV file path and processes them."""
    logging.info(f"Loading Google contacts from CSV: '{path}'")
    df = safe_read_csv(path)
    contacts = _process_google_df(df)
    logging.info(f"Loaded {len(contacts)} Google contacts from CSV.")
    return contacts


def load_google_contacts_from_api():
    """
    Loads Google Contacts via API, converts to a DataFrame, and processes it.
    """
    logging.info("Loading Google contacts from API.")
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            logging.info("Refreshing expired Google API token.")
            creds.refresh(Request())
        else:
            if not os.path.exists(CLIENT_SECRET_FILE):
                logging.error(f"'{CLIENT_SECRET_FILE}' not found.")
                raise FileNotFoundError(f"'{CLIENT_SECRET_FILE}' not found.")
            logging.info("Initiating new Google OAuth2 flow.")
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as token:
            token.write(creds.to_json())
            logging.info(f"Saved new Google API token to '{TOKEN_FILE}'.")

    service = build("people", "v1", credentials=creds)

    group_map = {}
    try:
        group_results = service.contactGroups().list().execute()
        for group in group_results.get("contactGroups", []):
            group_map[group.get("resourceName")] = group.get("formattedName")
        logging.info(f"Fetched {len(group_map)} contact groups.")
    except Exception as e:
        logging.warning(f"Could not fetch contact group names: {e}.")

    connections = []
    next_page_token = None
    while True:
        try:
            results = (
                service.people()
                .connections()
                .list(
                    resourceName="people/me",
                    pageSize=1000,
                    personFields="names,phoneNumbers,memberships,emailAddresses,organizations,addresses,biographies,birthdays,urls",
                    pageToken=next_page_token,
                )
                .execute()
            )
            connections.extend(results.get("connections", []))
            next_page_token = results.get("nextPageToken")
            logging.info(f"Fetched {len(connections)} connections so far...")
            if not next_page_token:
                break
        except Exception as e:
            logging.error(f"An error occurred while fetching connections: {e}")
            break

    rows = []
    for person in connections:
        row = {}
        names = person.get("names", [])
        if not names:
            continue

        primary_name = names[0]
        row["First Name"] = primary_name.get("givenName", "")
        row["Middle Name"] = primary_name.get("middleName", "")
        row["Last Name"] = primary_name.get("familyName", "")
        row["Name"] = primary_name.get("displayName", "")

        for i, phone in enumerate(person.get("phoneNumbers", [])[:5]):
            row[f"Phone {i+1} - Value"] = phone.get("value", "")
            row[f"Phone {i+1} - Type"] = phone.get("type", "Mobile")

        group_names = [
            group_map[m["contactGroupMembership"]["contactGroupResourceName"]]
            for m in person.get("memberships", [])
            if m.get("contactGroupMembership", {}).get("contactGroupResourceName")
            in group_map
        ]

        STAR_MAP = {
            "My Contacts": "* myContacts",
            "Family": "* family",
            "Starred": "* starred",
        }
        normalized_group_names = [STAR_MAP.get(name, name) for name in group_names]

        final_parts = []
        has_my_contacts = "* myContacts" in normalized_group_names
        has_starred = "* starred" in normalized_group_names
        other_groups = [
            n
            for n in normalized_group_names
            if n and n not in ["* myContacts", "* starred"]
        ]
        final_parts.extend(sorted(other_groups))
        if has_my_contacts:
            final_parts.append("* myContacts")
        if has_starred:
            final_parts.append("* starred")

        row["Labels"] = " ::: ".join(final_parts) if final_parts else "* myContacts"

        for i, email in enumerate(person.get("emailAddresses", [])[:4]):
            row[f"E-mail {i+1} - Value"] = email.get("value", "")
            row[f"E-mail {i+1} - Type"] = email.get("type", "Other")

        if orgs := person.get("organizations"):
            row["Organization 1 - Name"] = orgs[0].get("name", "")
            row["Organization 1 - Title"] = orgs[0].get("title", "")

        if addresses := person.get("addresses"):
            row["Address 1 - Formatted"] = addresses[0].get("formattedValue", "")
            row["Address 1 - Type"] = addresses[0].get("type", "Home")

        if bios := person.get("biographies"):
            row["Notes"] = bios[0].get("value", "")

        if birthdays := person.get("birthdays"):
            if bday := birthdays[0].get("date"):
                if year := bday.get("year"):
                    row["Birthday"] = (
                        f'{year:04d}-{bday.get("month", 1):02d}-{bday.get("day", 1):02d}'
                    )

        for i, site in enumerate(person.get("urls", [])[:4]):
            row[f"Website {i+1} - Value"] = site.get("value", "")
            row[f"Website {i+1} - Type"] = site.get("type", "Home Page")

        rows.append(row)

    if not rows:
        logging.warning("No contacts found via Google API.")
        return {}

    df = pd.DataFrame(rows)
    contacts = _process_google_df(df)
    logging.info(f"Loaded and processed {len(contacts)} contacts from Google API.")
    return contacts


def load_mssql_contacts(path):
    logging.info(f"Loading MSSQL contacts from CSV: '{path}'")
    df = safe_read_csv(path)
    contacts = {}
    for _, row in df.iterrows():
        full = str(row.iloc[0]).strip()
        nums_list = expand_normalize_numbers([v for v in list(row.iloc[1:]) if v])
        if not nums_list:
            continue

        parts = [p for p in full.split() if p]
        first = parts[0] if parts else ""
        middle = " ".join(parts[1:]) if len(parts) > 1 else ""
        last = "Lab"
        raw_display = (first + (" " + middle if middle else "") + (" " + last)).strip()
        display_name = normalize_display_name(raw_display, append_lab=True)
        cmp_name = strip_lab_token(display_name).lower()

        contacts[display_name] = {
            "numbers": set(nums_list),
            "sources": {"MSSQL"},
            "first_name": first,
            "middle_name": middle,
            "last_name": last,
            "original_name": full,
            "_cmp_name": cmp_name,
        }
    logging.info(f"Loaded {len(contacts)} MSSQL contacts from CSV.")
    return contacts


def load_mssql_from_db(
    server: str, database: str, user: str, password: str, query: str | None = None
):
    logging.info(f"Loading MSSQL contacts from database: {server}/{database}")
    try:
        import pyodbc
    except Exception as e:
        logging.error("pyodbc is required to load from MSSQL.")
        raise RuntimeError("pyodbc is required to load from MSSQL.") from e

    if not query:
        query = "select DISTINCT patientnamear, patientaddress, patientphone, patienttel, PFax from patients_1..patientinfo where patienttel != 'NULL' or patientaddress != 'NULL' or PFax != 'NULL' or patientphone != 'NULL' order by patientnamear asc"

    drivers = [
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 13 for SQL Server",
        "SQL Server Native Client 11.0",
        "SQL Server",
    ]
    conn = None
    last_err: str | Exception = (
        "No suitable ODBC driver was found or all connection attempts failed."
    )
    for drv in drivers:
        try:
            conn_str = f"DRIVER={{{drv}}};SERVER={server};DATABASE={database};UID={user};PWD={password};TrustServerCertificate=YES"
            conn = pyodbc.connect(conn_str, timeout=10)
            logging.info(f"Connected to MSSQL database with driver '{drv}'.")
            break
        except pyodbc.Error as e:
            last_err = e
    if conn is None:
        logging.error(f"Cannot connect to database. Last error: {last_err}")
        raise RuntimeError(f"Cannot connect to database. Last error: {last_err}")

    contacts = {}
    cur = conn.cursor()
    try:
        cur.execute(query)
        for row in cur.fetchall():
            full = str(row[0] or "").strip()
            nums_list = expand_normalize_numbers([str(v) for v in row[1:] if v])
            if not nums_list:
                continue

            parts = [p for p in full.split() if p]
            first = parts[0] if parts else ""
            middle = " ".join(parts[1:]) if len(parts) > 1 else ""
            last = "Lab"
            raw_display = (
                first + (" " + middle if middle else "") + (" " + last)
            ).strip()
            display_name = normalize_display_name(raw_display, append_lab=True)
            cmp_name = strip_lab_token(display_name).lower()

            contacts[display_name] = {
                "numbers": set(nums_list),
                "sources": {"MSSQL"},
                "first_name": first,
                "middle_name": middle,
                "last_name": last,
                "original_name": full,
                "_cmp_name": cmp_name,
            }
    except Exception as e:
        logging.error(f"Failed to fetch or process data from MSSQL: {e}")
    finally:
        cur.close()
        conn.close()
        logging.info("Closed MSSQL database connection.")

    logging.info(f"Loaded {len(contacts)} contacts from MSSQL database.")
    return contacts


# ===============================
# 🔄 MERGING LOGIC
# ===============================


def _merge_entry_into(merged, src_name, dst_name, phone_to_names):
    if src_name == dst_name or src_name not in merged or dst_name not in merged:
        return
    src, dst = merged[src_name], merged[dst_name]
    logging.debug(f"Merging '{src_name}' into '{dst_name}'.")

    for n in list(src.get("numbers") or []):
        dst.setdefault("numbers", set()).add(n)
        phone_to_names[n].add(dst_name)
        if src_name in phone_to_names[n]:
            try:
                phone_to_names[n].remove(src_name)
            except Exception:
                pass

    dst.setdefault("groups", set()).update(set(src.get("groups") or []))
    dst.setdefault("sources", set()).update(set(src.get("sources") or []))
    dst.setdefault("duplicates", set()).update(set(src.get("duplicates") or []))
    dst["duplicates"].add(src_name)

    if (not dst.get("_raw_row")) and src.get("_raw_row"):
        dst["_raw_row"] = deepcopy(src.get("_raw_row"))

    dst["protected"] = bool(dst.get("protected", False) or src.get("protected", False))
    try:
        del merged[src_name]
    except Exception:
        pass


def merge_contacts(google, mssql):
    phone_to_names = defaultdict(set)
    merged = {}
    per_row_logs = []
    logging.info("Starting contact merge process.")
    logging.info(f"Initial Google contacts: {len(google)}, MSSQL contacts: {len(mssql)}")

    google_cmp_index = {}
    for g_name, g_data in google.items():
        cmp_key = (g_data.get("_cmp_name") or strip_lab_token(g_name)).lower()
        if cmp_key and cmp_key not in google_cmp_index:
            google_cmp_index[cmp_key] = g_name
        for num in set(g_data.get("numbers") or []):
            phone_to_names[num].add(g_name)

        entry = merged.setdefault(
            g_name,
            {
                "numbers": set(),
                "groups": set(),
                "sources": set(),
                "duplicates": set(),
                "protected": False,
                "first_name": g_data.get("first_name", ""),
                "last_name": g_data.get("last_name", ""),
            },
        )
        entry["numbers"].update(set(g_data.get("numbers") or []))
        entry["groups"].update(set(g_data.get("groups") or []))
        entry["sources"].update(set(g_data.get("sources") or {"Google"}))
        entry["protected"] = bool(g_data.get("protected", False))
        if g_data.get("_raw_row"):
            entry["_raw_row"] = deepcopy(g_data.get("_raw_row"))

    logging.info(f"After initial Google processing, merged has {len(merged)} entries.")

    cmp_to_names = defaultdict(list)
    for nm in list(merged.keys()):
        g_cmp = (google.get(nm, {}).get("_cmp_name") or strip_lab_token(nm)).lower()
        cmp_to_names[g_cmp].append(nm)
    for names in cmp_to_names.values():
        if len(names) > 1:
            canonical = next(
                (n for n in names if not merged[n].get("protected", False)), names[0]
            )
            for n in names:
                if n != canonical and n in merged:
                    _merge_entry_into(merged, n, canonical, phone_to_names)
    logging.info(f"After name-based merging, merged has {len(merged)} entries.")

    for phone, names in list(phone_to_names.items()):
        names_list = [n for n in names if n in merged]
        if len(names_list) > 1:
            canonical = next(
                (n for n in names_list if not merged[n].get("protected", False)),
                names_list[0],
            )
            for n in names_list:
                if n != canonical and n in merged:
                    _merge_entry_into(merged, n, canonical, phone_to_names)
    logging.info(f"After phone-based merging, merged has {len(merged)} entries.")

    for m_name, m_data in mssql.items():
        existing = None
        m_nums = set(expand_normalize_numbers(list(m_data.get("numbers") or [])))
        for num in m_nums:
            if num in phone_to_names:
                existing = list(phone_to_names[num])[0]
                break
        if not existing:
            m_cmp = (m_data.get("_cmp_name") or strip_lab_token(m_name)).lower()
            if m_cmp in google_cmp_index:
                existing = google_cmp_index[m_cmp]

        if existing:
            if not merged[existing]["protected"]:
                original_raw = google.get(existing, {}).get("_raw_row")
                update_data = {
                    "mssql_name": m_name,
                    "mssql_original_name": m_data.get("original_name"),
                    "added_numbers": sorted(
                        m_nums - set(merged[existing].get("numbers") or [])
                    ),
                    "added_first_name": False,
                    "added_last_name": False,
                }
                merged[existing]["numbers"].update(m_nums)
                for n in m_nums:
                    phone_to_names[n].add(existing)
                merged[existing]["sources"].add("MSSQL")
                if not merged[existing].get("first_name") and m_data.get("first_name"):
                    merged[existing]["first_name"] = m_data.get("first_name")
                    update_data["added_first_name"] = True
                if not merged[existing].get("last_name") and m_data.get("last_name"):
                    merged[existing]["last_name"] = m_data.get("last_name")
                    update_data["added_last_name"] = True
                merged[existing]["duplicates"].add(
                    m_data.get("original_name") or m_name
                )
                final_snapshot = {
                    "First Name": existing,
                    "Phones": sorted(list(merged[existing].get("numbers") or [])),
                    "Group Membership": " ::: ".join(
                        sorted(merged[existing].get("groups") or [])
                    ),
                    "Sources": " & ".join(
                        sorted(merged[existing].get("sources") or [])
                    ),
                    "Duplicates": " - ".join(
                        sorted(merged[existing].get("duplicates") or [])
                    ),
                }
                per_row_logs.append(
                    {
                        "google_name": existing,
                        "original_google_row": original_raw,
                        "update_data": update_data,
                        "final_row": final_snapshot,
                    }
                )
            continue

        if not m_nums:
            continue
        entry = merged.setdefault(
            m_name,
            {
                "numbers": set(),
                "groups": set(),
                "sources": set(),
                "duplicates": set(),
                "protected": False,
                "first_name": m_data.get("first_name", ""),
                "last_name": m_data.get("last_name", ""),
            },
        )
        entry["numbers"].update(m_nums)
        for n in m_nums:
            phone_to_names[n].add(m_name)
        entry["groups"].add(DEFAULT_NEW_GROUP)
        entry["sources"].add("MSSQL")
        nm_norm = m_name.strip().lower()
        for existing_name in list(merged.keys()):
            if existing_name.strip().lower() == nm_norm and existing_name != m_name:
                entry["duplicates"].add(existing_name)

    logging.info(f"After processing MSSQL contacts, merged has {len(merged)} entries.")

    for phone, names in list(phone_to_names.items()):
        names_list = [n for n in names if n in merged]
        if len(names_list) > 1:
            canonical = next(
                (n for n in names_list if not merged[n].get("protected", False)),
                names_list[0],
            )
            for n in names_list:
                if n != canonical and n in merged:
                    _merge_entry_into(merged, n, canonical, phone_to_names)

    logging.info(f"Final merged contact count: {len(merged)}")
    return merged, per_row_logs


def write_detailed_log(per_row_logs, summary_data, timestamp):
    """Writes the detailed per-row logs and a summary to a timestamped JSON file."""
    if not per_row_logs:
        logging.info("No detailed merge logs to write.")
        return

    log_path = os.path.join(LOG_DIR, f"merge_log_{timestamp}.json")
    logging.info(f"Writing detailed merge log to '{log_path}'")

    # Prepare the full log data with a summary section
    full_log = {
        "summary": summary_data,
        "details": per_row_logs
    }

    try:
        with open(log_path, "w", encoding="utf-8") as f:
            json.dump(full_log, f, indent=2, ensure_ascii=False)
        logging.info(f"Successfully wrote {len(per_row_logs)} detailed log entries with summary.")
    except Exception as e:
        logging.error(f"Failed to write detailed log: {e}")


def write_detailed_csv_log(per_row_logs, timestamp):
    """Writes the detailed per-row logs to a timestamped CSV file."""
    if not per_row_logs:
        return

    log_path = os.path.join(LOG_DIR, f"merge_log_{timestamp}.csv")
    logging.info(f"Writing detailed CSV log to '{log_path}'")

    rows = []
    for log_entry in per_row_logs:
        google_name = log_entry.get("google_name", "")
        update_data = log_entry.get("update_data", {})
        final_row = log_entry.get("final_row", {})

        row = {
            "Google Contact": google_name,
            "MSSQL Name": update_data.get("mssql_name", ""),
            "MSSQL Original Name": update_data.get("mssql_original_name", ""),
            "Added Numbers": ", ".join(update_data.get("added_numbers", [])),
            "Added First Name": update_data.get("added_first_name", False),
            "Added Last Name": update_data.get("added_last_name", False),
            "Final Phone Numbers": final_row.get("Phones", ""),
            "Final Group Membership": final_row.get("Group Membership", ""),
            "Final Sources": final_row.get("Sources", ""),
            "Final Duplicates": final_row.get("Duplicates", ""),
        }
        rows.append(row)

    if not rows:
        return

    try:
        pd.DataFrame(rows).to_csv(log_path, index=False, encoding="utf-8-sig")
        logging.info(f"Successfully wrote {len(rows)} detailed log entries to CSV.")
    except Exception as e:
        logging.error(f"Failed to write detailed CSV log: {e}")


# ===============================
# 📤 EXPORT
# ===============================


def export_contacts(
    merged, output_file="merged_contacts.csv", template: str | None = None
):
    logging.info(f"Exporting {len(merged)} contacts to '{output_file}'.")
    fieldnames = None
    if template:
        try:
            df = safe_read_csv(template)
            fieldnames = list(df.columns)
            logging.info(f"Using template '{template}' for export fields.")
        except Exception as e:
            logging.warning(f"Could not read template file '{template}': {e}")
    if fieldnames and "Name" in fieldnames:
        try:
            fieldnames = [c for c in fieldnames if c != "Name"]
        except Exception:
            pass

    if not fieldnames:
        fieldnames = [
            "First Name",
            "Middle Name",
            "Last Name",
            "Group Membership",
            "Phone 1 - Type",
            "Phone 1 - Value",
            "Phone 2 - Type",
            "Phone 2 - Value",
            "Phone 3 - Type",
            "Phone 3 - Value",
            "Phone 4 - Type",
            "Phone 4 - Value",
            "Labels",
            "Custom Field 1 - Label",
            "Custom Field 1 - Value",
            "Custom Field 2 - Label",
            "Custom Field 2 - Value",
        ]
        logging.info("No template found. Using default field names for export.")

    required_cols = [
        "First Name",
        "Middle Name",
        "Last Name",
        "Group Membership",
        "Phone 1 - Type",
        "Phone 1 - Value",
        "Phone 2 - Type",
        "Phone 2 - Value",
        "Phone 3 - Type",
        "Phone 3 - Value",
        "Phone 4 - Type",
        "Phone 4 - Value",
        "Labels",
        "Custom Field 1 - Label",
        "Custom Field 1 - Value",
        "Custom Field 2 - Label",
        "Custom Field 2 - Value",
    ]
    for col in required_cols:
        if col not in fieldnames:
            fieldnames.append(col)

    rows = []
    for name, data in merged.items():
        incoming_vals = expand_normalize_numbers(list(data.get("numbers") or []))
        groups = " ::: ".join(sorted(data.get("groups") or [])) or "* myContacts"
        groups = normalize_group_name(groups)
        sources = " & ".join(sorted(set(data.get("sources") or [])))
        duplicates = " - ".join(sorted(data.get("duplicates") or []))

        row = {col: "" for col in fieldnames}
        if original_row := data.get("_raw_row"):
            for col in fieldnames:
                row[col] = original_row.get(col, "")
        else:
            row["First Name"] = name
            if "Labels" in row:
                row["Labels"] = groups or (row.get("Labels") or "Merged Contact")

        existing_cells = [row.get(f"Phone {i} - Value", "") for i in range(1, 5)]
        existing_vals = expand_normalize_numbers(existing_cells)
        final_phones = []
        for v in existing_vals + incoming_vals:
            if v not in final_phones:
                final_phones.append(v)
        final_phones = final_phones[:4]
        cc_only = (
            DEFAULT_COUNTRY_CODE
            if DEFAULT_COUNTRY_CODE.startswith("+")
            else "+" + DEFAULT_COUNTRY_CODE
        )
        final_phones = [p for p in final_phones if p and p != cc_only]

        for i in range(1, 5):
            row[f"Phone {i} - Value"] = ""
        for i, val in enumerate(final_phones, 1):
            row[f"Phone {i} - Value"] = val
            if val and not row.get(f"Phone {i} - Type"):
                row[f"Phone {i} - Type"] = "Mobile"

        if "Labels" in row:
            row["Labels"] = groups
        if "Custom Field 1 - Label" in row and not row.get("Custom Field 1 - Label"):
            row["Custom Field 1 - Label"] = "Duplicate Names"
        if "Custom Field 1 - Value" in row:
            row["Custom Field 1 - Value"] = duplicates
        if "Custom Field 2 - Label" in row and not row.get("Custom Field 2 - Label"):
            row["Custom Field 2 - Label"] = "Sources"
        if "Custom Field 2 - Value" in row:
            row["Custom Field 2 - Value"] = sources
        rows.append(row)

    try:
        pd.DataFrame(rows, columns=fieldnames).to_csv(
            output_file, index=False, encoding="utf-8-sig"
        )
        logging.info(f"Successfully saved {len(rows)} contacts to '{output_file}'")
    except Exception as e:
        logging.error(f"Failed to save contacts to '{output_file}': {e}")



