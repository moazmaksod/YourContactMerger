import pandas as pd
import re
from collections import defaultdict, Counter
from copy import deepcopy

# ===============================
# 🔧 CONFIGURATION SECTION
# ===============================

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

# Protected groups are any that are not Lab; only Lab contacts are modifiable
# Keeping constant for reference but not used for logic anymore
PROTECTED_KEYWORDS = []

# Default group for new MSSQL contacts
DEFAULT_NEW_GROUP = "🧪 Lab ::: * myContacts"

# Default country code (used as fallback for local numbers). Include the leading '+' if you like.
DEFAULT_COUNTRY_CODE = "+20"

# ===============================
# 🧠 UTILITIES
# ===============================


def safe_read_csv(path):
    """Read CSV with automatic encoding detection (for UTF-16 MSSQL exports)."""
    # Try a wider range of encodings and use the python engine which is more tolerant
    encodings_to_try = [
        "utf-8",
        "utf-8-sig",
        "utf-16",
        "utf-16-le",
        "utf-16-be",
        "cp1252",
        "latin1",
    ]
    for enc in encodings_to_try:
        try:
            # use the python engine to be tolerant of irregular separators/quotes
            df = pd.read_csv(
                path, dtype=str, encoding=enc, on_bad_lines="skip", engine="python"
            ).fillna("")
            df.columns = [
                c.strip().replace("\ufeff", "").replace("ÿþ", "") for c in df.columns
            ]
            return df
        except Exception:
            continue
    # Last resort: try opening file as binary and decoding with 'latin1' fallback
    try:
        import io

        with open(path, "rb") as fh:
            data = fh.read()
        text = data.decode("latin1", errors="replace")
        df = pd.read_csv(
            io.StringIO(text), dtype=str, on_bad_lines="skip", engine="python"
        ).fillna("")
        df.columns = [
            c.strip().replace("\ufeff", "").replace("ÿþ", "") for c in df.columns
        ]
        return df
    except Exception:
        raise ValueError(f"Cannot read file {path} with tried encodings.")


def clean_phone(phone: str) -> str:
    """Clean phone number of unwanted characters."""
    return re.sub(r"[^\d+]", "", str(phone)).strip()


def normalize_phone(num: str) -> str | None:
    """Normalize phone number based on pattern and country rules."""
    if num is None:
        return None
    s = str(num).strip()
    if not s or s.lower() == "null":
        return None

    num = clean_phone(s)

    # Already international
    if num.startswith("+"):
        return num

    # Replace 00 with +
    if num.startswith("00"):
        num = "+" + num[2:]
        return num

    # Egypt
    if re.match(r"^01[0-9]{8,}$", num):
        return "+2" + num

    # Saudi Arabia
    if re.match(r"^05[0-9]{8,}$", num):
        return "+966" + num[1:]

    # UAE
    if re.match(r"^971[0-9]{7,}$", num):
        return "+" + num
    if re.match(r"^05[0-9]{7,}$", num):
        return "+971" + num[1:]

    # Turkey
    if re.match(r"^90[0-9]{9}$", num):
        return "+" + num

    # Russia
    if re.match(r"^7[0-9]{9,}$", num):
        return "+" + num

    # UK
    if re.match(r"^44[0-9]{9,}$", num):
        return "+" + num

    # Default: build international with DEFAULT_COUNTRY_CODE only if there are local digits
    if not num.startswith("+"):
        digits_only = re.sub(r"\D", "", num)
        digits_only = digits_only.lstrip("0")
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
    """Convert Arabic/mixed group names to clean English names and remove starred markers."""
    group_value = (group_value or "").strip()
    group_value = group_value.replace("::: * starred", "").strip()

    for old, new in GROUP_MAP.items():
        if old in group_value:
            group_value = group_value.replace(old, new)

    return group_value


def strip_lab_token(name: str) -> str:
    s = re.sub(r"(?i)\blab\b", "", str(name or ""))
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_display_name(raw: str, append_lab: bool = True, preserve_lab: bool = False) -> str:
    """Normalize display name for output.
    - when preserve_lab is True, do not remove existing 'Lab' tokens from the visible name
    - when append_lab is True, ensure a trailing ' Lab' exists
    - always collapse whitespace
    """
    if not raw:
        return " Lab" if append_lab else ""
    s = str(raw)
    if not preserve_lab:
        s = re.sub(r"(?i)\blab\b", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    if append_lab and not s.endswith(" Lab"):
        s = (s + " Lab").strip()
    return s

# ===============================
# 📥 LOAD CONTACTS
# ===============================


def expand_normalize_numbers(seq):
    """Expand ':::' concatenations, strip and normalize with normalize_phone, dedupe preserving order."""
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


def load_google_contacts(path):
    df = safe_read_csv(path)
    contacts = {}
    for _, row in df.iterrows():
        # capture name parts when present; build full name from first/middle/last if available
        first = (row.get("First Name") or "").strip()
        middle = (row.get("Middle Name") or "").strip()
        last = (row.get("Last Name") or "").strip()
        if first or middle or last:
            raw_name = " ".join([p for p in [first, middle, last] if p]).strip()
        else:
            raw_name = (row.get("Name") or "").strip()
        # groups: prefer Labels, fallback to Group Membership
        groups_raw = row.get("Labels") or row.get("Group Membership") or "* myContacts"
        groups = normalize_group_name(groups_raw)
        # robust lab detection over tokens
        def _is_lab_group(g: str) -> bool:
            tokens = [t.strip().lower() for t in (g or "").split(":::")]
            for t in tokens:
                # remove emojis and punctuation for matching
                t_clean = re.sub(r"[^\w\s]", "", t)
                if "lab" in t_clean:
                    return True
            return False
        is_lab = _is_lab_group(groups)
        # Preserve Lab in Google-visible names; append ' Lab' if belongs to Lab group
        name = normalize_display_name(raw_name, append_lab=is_lab, preserve_lab=True)
        # keep cmp name separately for matching only
        cmp_name = strip_lab_token(name).lower()
        # expand phones from google cells (split on :::, normalize, dedupe)
        existing_cells = [row.get(f"Phone {i} - Value", "") for i in range(1, 5)]
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


def load_mssql_contacts(path):
    df = safe_read_csv(path)
    contacts = {}

    for _, row in df.iterrows():
        # assume first column is full name, rest are numbers
        full = str(row.iloc[0]).strip()
        # expand/normalize/dedupe numbers from remaining columns with possible ::: concatenations
        nums_list = expand_normalize_numbers([v for v in list(row.iloc[1:]) if v])
        # skip MSSQL rows with no valid numbers
        if not nums_list:
            continue

        parts = [p for p in full.split() if p]
        first = parts[0] if parts else ""
        middle = " ".join(parts[1:]) if len(parts) > 1 else ""
        # per your rule, set last name to 'Lab' for MSSQL-derived contacts
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

    return contacts


def load_mssql_from_db(
    server: str, database: str, user: str, password: str, query: str | None = None
):
    """
    Load MSSQL contacts directly from a SQL Server database using pyodbc.
    Returns a dict of contacts in the same shape as load_mssql_contacts.
    """
    try:
        import pyodbc
    except Exception as e:
        raise RuntimeError(
            "pyodbc is required to load from MSSQL. Please install it (pip install pyodbc)."
        ) from e

    if not query:
        query = """
select DISTINCT
patientnamear,
patientaddress,
patientphone,
patienttel,
PFax
from patients_1..patientinfo
where
    patienttel != 'NULL'
    or patientaddress != 'NULL'
    or PFax != 'NULL'
    or patientphone != 'NULL'
order by patientnamear asc
"""

    # try common SQL Server ODBC drivers
    drivers = [
        "ODBC Driver 17 for SQL Server",
        "ODBC Driver 13 for SQL Server",
        "SQL Server Native Client 11.0",
        "SQL Server",
    ]

    last_err = None
    conn = None
    for drv in drivers:
        try:
            conn_str = f"DRIVER={{{drv}}};SERVER={server};DATABASE={database};UID={user};PWD={password};TrustServerCertificate=YES"
            conn = pyodbc.connect(conn_str, timeout=10)
            break
        except Exception as e:
            last_err = e
            conn = None

    if conn is None:
        raise RuntimeError(f"Cannot connect to database. Last error: {last_err}")

    contacts = {}
    cur = conn.cursor()
    cur.execute(query)
    cols = [c[0] for c in cur.description]
    for row in cur.fetchall():
        # row is a tuple; first value is name, others may contain phones
        full = str(row[0] or "").strip()
        nums_list = expand_normalize_numbers([str(v) for v in row[1:] if v])
        # skip MSSQL DB rows with no valid numbers
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

    cur.close()
    conn.close()
    return contacts

# ===============================
# 🔄 MERGING LOGIC
# ===============================


def _merge_entry_into(merged, src_name, dst_name, phone_to_names):
    """Merge merged[src_name] into merged[dst_name] and delete src entry. Update phone_to_names."""
    if src_name == dst_name or src_name not in merged or dst_name not in merged:
        return
    src = merged[src_name]
    dst = merged[dst_name]
    # numbers
    for n in list(src.get("numbers") or []):
        dst.setdefault("numbers", set()).add(n)
        phone_to_names[n].add(dst_name)
        if src_name in phone_to_names[n]:
            try:
                phone_to_names[n].remove(src_name)
            except Exception:
                pass
    # groups, sources, duplicates
    dst.setdefault("groups", set()).update(set(src.get("groups") or []))
    dst.setdefault("sources", set()).update(set(src.get("sources") or []))
    dst.setdefault("duplicates", set()).update(set(src.get("duplicates") or []))
    dst["duplicates"].add(src_name)
    # preserve a raw row: prefer dst existing; if none and src has one, take it
    if (not dst.get("_raw_row")) and src.get("_raw_row"):
        dst["_raw_row"] = deepcopy(src.get("_raw_row"))
    # protected flag: if either protected, keep protected
    dst["protected"] = bool(dst.get("protected", False) or src.get("protected", False))
    # remove src
    try:
        del merged[src_name]
    except Exception:
        pass


def merge_contacts(google, mssql):
    phone_to_names = defaultdict(set)
    merged = {}
    per_row_logs = []

    # Build Google comparison index and seed merged entries
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
        # carry over raw row for export preservation
        if g_data.get("_raw_row"):
            entry["_raw_row"] = deepcopy(g_data.get("_raw_row"))

    # Consolidate duplicate Google contacts by comparison name
    cmp_to_names = defaultdict(list)
    for nm in list(merged.keys()):
        # Use google-provided cmp name when available; fallback to stripped visible name
        g_cmp = (google.get(nm, {}).get("_cmp_name") or strip_lab_token(nm)).lower()
        cmp_to_names[g_cmp].append(nm)
    for key, names in cmp_to_names.items():
        if len(names) > 1:
            # choose canonical: prefer unprotected if any, else first
            canonical = None
            for n in names:
                if not merged[n].get("protected", False):
                    canonical = n
                    break
            if not canonical:
                canonical = names[0]
            for n in names:
                if n != canonical and n in merged:
                    _merge_entry_into(merged, n, canonical, phone_to_names)

    # Consolidate by phone overlap among Google contacts
    for phone, names in list(phone_to_names.items()):
        names_list = [n for n in names if n in merged]
        if len(names_list) > 1:
            canonical = None
            for n in names_list:
                if not merged[n].get("protected", False):
                    canonical = n
                    break
            if not canonical:
                canonical = names_list[0]
            for n in names_list:
                if n != canonical and n in merged:
                    _merge_entry_into(merged, n, canonical, phone_to_names)

    # merge MSSQL contacts
    for m_name, m_data in mssql.items():
        existing = None

        # normalize mssql numbers again deterministically
        m_nums = set(expand_normalize_numbers(list(m_data.get("numbers") or [])))

        # Try phone-based match first
        for num in m_nums:
            if num in phone_to_names:
                existing = list(phone_to_names[num])[0]
                break

        # Fallback to name-based comparison
        if not existing:
            m_cmp = (m_data.get("_cmp_name") or strip_lab_token(m_name)).lower()
            if m_cmp in google_cmp_index:
                existing = google_cmp_index[m_cmp]

        if existing:
            # Merge into existing contact if not protected
            if not merged[existing]["protected"]:
                original_raw = None
                try:
                    original_raw = google.get(existing, {}).get("_raw_row")
                except Exception:
                    original_raw = None

                current_numbers = set(merged[existing].get("numbers") or [])
                new_numbers = m_nums
                numbers_to_add = new_numbers - current_numbers

                update_data = {
                    "mssql_name": m_name,
                    "mssql_original_name": m_data.get("original_name"),
                    "added_numbers": sorted(numbers_to_add),
                    "added_first_name": False,
                    "added_last_name": False,
                }

                merged[existing]["numbers"].update(new_numbers)
                for n in new_numbers:
                    phone_to_names[n].add(existing)

                merged[existing]["sources"].add("MSSQL")

                if (
                    not merged[existing].get("first_name")
                    or merged[existing].get("first_name") == ""
                ) and m_data.get("first_name"):
                    merged[existing]["first_name"] = m_data.get("first_name")
                    update_data["added_first_name"] = True
                if (
                    not merged[existing].get("last_name")
                    or merged[existing].get("last_name") == ""
                ) and m_data.get("last_name"):
                    merged[existing]["last_name"] = m_data.get("last_name")
                    update_data["added_last_name"] = True

                merged[existing]["duplicates"].add(m_data.get("original_name") or m_name)

                final_snapshot = {
                    "First Name": existing,
                    "Middle Name": "",
                    "Last Name": "",
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

        # Do not create MSSQL-only entry if no numbers
        if not m_nums:
            continue
        # New contact entirely — create with defaults and mark group
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
        # index new MSSQL-only numbers so global consolidation can detect shared phones
        for n in m_nums:
            phone_to_names[n].add(m_name)
        entry["groups"].add(DEFAULT_NEW_GROUP)
        entry["sources"].add("MSSQL")

        nm_norm = m_name.strip().lower()
        for existing_name in list(merged.keys()):
            if existing_name.strip().lower() == nm_norm and existing_name != m_name:
                entry["duplicates"].add(existing_name)

    # Final global consolidation by shared phone across all sources
    for phone, names in list(phone_to_names.items()):
        names_list = [n for n in names if n in merged]
        if len(names_list) > 1:
            canonical = None
            for n in names_list:
                if not merged[n].get("protected", False):
                    canonical = n
                    break
            if not canonical:
                canonical = names_list[0]
            for n in names_list:
                if n != canonical and n in merged:
                    _merge_entry_into(merged, n, canonical, phone_to_names)

    try:
        global LAST_PER_ROW_LOGS
        LAST_PER_ROW_LOGS = per_row_logs
    except Exception:
        pass
    return merged, per_row_logs

# ===============================
# 📤 EXPORT
# ===============================


def export_contacts(
    merged, output_file="merged_contacts.csv", template: str | None = None
):
    """
    Export merged contacts using the column order from the template CSV (usually Google contacts export).
    If template is None or cannot be read, falls back to a default compact schema.
    """
    # try to load fieldnames from template if provided
    fieldnames = None
    if template:
        try:
            df = safe_read_csv(template)
            fieldnames = list(df.columns)
        except Exception:
            fieldnames = None

    # Do not include a separate 'Name' column in output; user requested full name goes to 'First Name'
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

    # Ensure required Google columns exist (append missing ones while keeping template order)
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
        # Build incoming values from merged data numbers
        incoming_vals = expand_normalize_numbers(list(data.get("numbers") or []))
        groups = " ::: ".join(sorted(data.get("groups") or [])) or "* myContacts"
        # normalize final group membership string using GROUP_MAP
        groups = normalize_group_name(groups)
        sources_set = set(data.get("sources") or [])
        sources = " & ".join(sorted(sources_set))
        duplicates = " - ".join(sorted(data.get("duplicates") or []))

        # Base row: preserve Google row if present; otherwise start empty
        original_row = data.get("_raw_row") if ("Google" in sources_set) else None
        row = {col: "" for col in fieldnames}
        if original_row is not None and ("Google" in sources_set):
            for col in fieldnames:
                row[col] = original_row.get(col, "")
        else:
            # MSSQL only: set minimal fields
            row["First Name"] = name
            row["Middle Name"] = ""
            row["Last Name"] = ""
            # Do not write Group Membership; use Labels only
            if "Labels" in row:
                row["Labels"] = groups or (row.get("Labels") or "Merged Contact")

        # Phones: build existing from row cells, then merge with incoming, normalize/dedupe, overwrite slots
        existing_cells = [row.get(f"Phone {i} - Value", "") for i in range(1, 5)]
        existing_vals = expand_normalize_numbers(existing_cells)
        final_phones = []
        for v in existing_vals + incoming_vals:
            if v not in final_phones:
                final_phones.append(v)
        final_phones = final_phones[:4]
        # filter out bare country code artifact if any
        cc_only = DEFAULT_COUNTRY_CODE if DEFAULT_COUNTRY_CODE.startswith("+") else "+" + DEFAULT_COUNTRY_CODE
        final_phones = [p for p in final_phones if p and p != cc_only]

        # Clear and set phone slots; set default type if empty when value set
        for i in range(1, 5):
            row[f"Phone {i} - Value"] = ""
        for i in range(1, 5):
            val = final_phones[i - 1] if i - 1 < len(final_phones) else ""
            row[f"Phone {i} - Value"] = val
            if val and not row.get(f"Phone {i} - Type", ""):
                row[f"Phone {i} - Type"] = "Mobile"

        # Set Labels to groups for all rows (write groups in Labels only)
        if "Labels" in row:
            row["Labels"] = groups

        # custom fields for duplicates and sources
        if "Custom Field 1 - Label" in row and not row.get("Custom Field 1 - Label"):
            row["Custom Field 1 - Label"] = "Duplicate Names"
        if "Custom Field 1 - Value" in row:
            row["Custom Field 1 - Value"] = duplicates
        if "Custom Field 2 - Label" in row and not row.get("Custom Field 2 - Label"):
            row["Custom Field 2 - Label"] = "Sources"
        if "Custom Field 2 - Value" in row:
            row["Custom Field 2 - Value"] = sources

        rows.append(row)

    # use utf-8-sig for better Excel compatibility with non-ascii characters
    pd.DataFrame(rows, columns=fieldnames).to_csv(
        output_file, index=False, encoding="utf-8-sig"
    )
    print(f"Saved {len(rows)} contacts to '{output_file}'")


# ===============================
# 🚀 MAIN EXECUTION
# ===============================
if __name__ == "__main__":
    print("Loading contacts...")
    google_contacts = load_google_contacts("contacts.csv")
    mssql_contacts = load_mssql_contacts("contact numbers 24-10-2025.csv")

    print(f"Google contacts loaded: {len(google_contacts)}")
    print(f"MSSQL contacts loaded: {len(mssql_contacts)}")

    print("Merging...")
    merged_res = merge_contacts(google_contacts, mssql_contacts)
    # merge_contacts returns (merged, logs)
    if isinstance(merged_res, (tuple, list)):
        merged = merged_res[0]
    else:
        merged = merged_res

    print("Exporting merged contacts...")
    export_contacts(merged)

    print("Merge complete.")
