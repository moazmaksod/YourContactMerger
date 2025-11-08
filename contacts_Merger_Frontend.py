import flet as ft
import threading
import os
import datetime
import Contacts_Merger_Backend as backend
import keyring

LOGO_PATH = "assets/logo.png"

# -------------------------------
# 🎨 Centralized Theme Colors
# -------------------------------
THEME = {
    "primary": ft.Colors.BLUE_600,
    "primary_dark": ft.Colors.BLUE_800,
    "surface": ft.Colors.GREY_50,
    "surface_dark": ft.Colors.BLUE_GREY_900,
    "text_primary": ft.Colors.BLACK,
    "text_primary_dark": ft.Colors.WHITE,
    "text_secondary": ft.Colors.GREY_700,
    "text_secondary_dark": ft.Colors.GREY_400,
    "metric_bg": ft.Colors.GREY_100,
    "metric_bg_dark": ft.Colors.BLUE_GREY_800,
    "success": ft.Colors.GREEN_600,
    "warning": ft.Colors.ORANGE_700,
    "error": ft.Colors.RED_700,
    "accent": ft.Colors.CYAN_700,
}


def theme_color(page, key):
    """Return color based on current theme."""
    dk = f"{key}_dark"
    if page.theme_mode == ft.ThemeMode.DARK and dk in THEME:
        return THEME[dk]
    return THEME.get(key, ft.Colors.GREY_500)


# Globals for theme update
boxes = []
summary_metrics = []


def main(page: ft.Page):
    page.title = "Contacts Merger v3.6.3"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window.width = 980
    page.window.height = 720
    page.window.min_width = 720
    page.window.min_height = 640
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 20
    page.horizontal_alignment = ft.CrossAxisAlignment.STRETCH

    state = {"google": None, "csvs": []}

    # ---------------------------
    # Section helper
    # ---------------------------
    def section(content):
        container = ft.Container(
            content=content,
            bgcolor=theme_color(page, "surface"),
            border_radius=10,
            padding=12,
            shadow=ft.BoxShadow(
                blur_radius=8, color=ft.Colors.BLACK12, offset=ft.Offset(0, 3)
            ),
            expand=True,
        )
        boxes.append(container)
        return container

    # ---------------------------
    # Header
    # ---------------------------
    def toggle_theme(e):
        page.theme_mode = (
            ft.ThemeMode.DARK
            if page.theme_mode == ft.ThemeMode.LIGHT
            else ft.ThemeMode.LIGHT
        )
        refresh_theme()
        page.update()

    logo = (
        ft.Image(src=LOGO_PATH, width=40, height=40)
        if os.path.exists(LOGO_PATH)
        else ft.Icon(ft.Icons.PERSON, size=40, color=ft.Colors.WHITE)
    )
    title = ft.Text(
        "Contacts Merger", size=22, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE
    )
    subtitle = ft.Text(
        "Smart merge tool for Google & MSSQL contacts", size=13, color=ft.Colors.WHITE70
    )
    theme_btn = ft.IconButton(
        icon=ft.Icons.BRIGHTNESS_6,
        icon_color=ft.Colors.WHITE,
        tooltip="Toggle theme",
        on_click=toggle_theme,
    )

    header = ft.Container(
        ft.Row(
            [
                logo,
                ft.Column([title, subtitle], spacing=0),
                ft.Container(expand=True),
                theme_btn,
            ],
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        ),
        bgcolor=theme_color(page, "primary"),
        padding=ft.Padding(16, 12, 16, 12),
        border_radius=10,
    )

    # ------------------------------------
    # Google Contacts Source Selection
    # ------------------------------------
    google_source_method = ft.RadioGroup(
        content=ft.Row(
            [
                ft.Radio(value="csv", label="CSV File"),
                ft.Radio(value="api", label="Google API"),
            ],
            alignment=ft.MainAxisAlignment.START,
        ),
        value="csv",
    )

    # --- CSV Controls ---
    google_field = ft.TextField(
        label="Google Contacts CSV", read_only=True, dense=True, expand=True
    )
    google_picker = ft.FilePicker(on_result=lambda e: pick_google(e))
    google_btn = ft.FilledButton(
        "Browse",
        icon=ft.Icons.UPLOAD_FILE,
        height=40,
        style=ft.ButtonStyle(
            color={ft.ControlState.DEFAULT: ft.Colors.WHITE},
            bgcolor={ft.ControlState.DEFAULT: theme_color(page, "primary")},
            shape=ft.RoundedRectangleBorder(radius=8),
        ),
        on_click=lambda _: google_picker.pick_files(),
    )
    csv_controls = ft.Row([google_field, google_btn], spacing=6, visible=True)

    # --- API Controls ---
    api_status_text = ft.Text("Ready to sync.", size=13, color=theme_color(page, "text_secondary"))
    
    def sync_google_api_click(e):
        # This function will run in a separate thread
        api_status_text.value = "Syncing... Please check your browser to authorize."
        api_status_text.color = theme_color(page, "text_secondary")
        page.update()
        try:
            google_contacts_data = backend.load_google_contacts_from_api()
            if google_contacts_data is not None:
                num_contacts = len(google_contacts_data)
                state["google_contacts"] = google_contacts_data
                state["google_source"] = "api"
                state["google_path"] = None # Clear path if we are using API
                google_field.value = f"API: {num_contacts} contacts synced"
                api_status_text.value = f"✅ Synced {num_contacts} contacts."
                api_status_text.color = theme_color(page, "success")
            else:
                api_status_text.value = "Error: Received no data from API."
                api_status_text.color = theme_color(page, "error")
        except Exception as err:
            api_status_text.value = f"Error: {err}"
            api_status_text.color = theme_color(page, "error")
        page.update()

    def google_api_logout(e):
        try:
            token_path = backend.TOKEN_FILE
            if os.path.exists(token_path):
                os.remove(token_path)
            
            api_status_text.value = "Ready to sync."
            api_status_text.color = theme_color(page, "text_secondary")
            google_field.value = ""
            state["google_contacts"] = None
            
            page.open(ft.SnackBar(ft.Text("Logged out. You can now sync with a different account.")))
        except Exception as err:
            page.open(ft.SnackBar(ft.Text(f"Error: {err}")))
        page.update()

    google_api_btn = ft.FilledButton(
        "Sync from Google",
        icon=ft.Icons.SYNC,
        height=40,
        style=ft.ButtonStyle(
            color={ft.ControlState.DEFAULT: ft.Colors.WHITE},
            bgcolor={ft.ControlState.DEFAULT: theme_color(page, "primary")},
            shape=ft.RoundedRectangleBorder(radius=8),
        ),
        on_click=lambda e: threading.Thread(target=sync_google_api_click, args=(e,), daemon=True).start(),
    )

    logout_btn = ft.TextButton(
        "Switch Account",
        icon=ft.Icons.LOGOUT,
        on_click=google_api_logout,
        tooltip="Log out to connect with a different Google account"
    )

    api_controls = ft.Column(
        [
            ft.Row([google_api_btn, logout_btn], alignment=ft.MainAxisAlignment.START, spacing=10),
            api_status_text
        ], 
        spacing=4, 
        visible=False
    )

    def on_google_source_change(e):
        selected_source = e.control.value
        state["google_source"] = selected_source
        if selected_source == "csv":
            csv_controls.visible = True
            api_controls.visible = False
        else:
            csv_controls.visible = False
            api_controls.visible = True
        page.update()

    google_source_method.on_change = on_google_source_change

    # --- Other File Pickers ---
    csv_list = ft.ListView(height=100, spacing=4)
    csv_picker = ft.FilePicker(on_result=lambda e: pick_csvs(e))
    page.overlay.extend([google_picker, csv_picker])

    def pick_google(e):
        if e.files:
            f = e.files[0]
            google_field.value = f.name
            state["google_path"] = f.path
            state["google_source"] = "csv"
            state["google_contacts"] = None # Clear contacts if path changes
            page.update()

    def pick_csvs(e):
        if e.files:
            for f in e.files:
                if f.path not in [x["path"] for x in state["csvs"]]:
                    state["csvs"].append({"name": f.name, "path": f.path})
            refresh_csvs()

    def refresh_csvs():
        csv_list.controls.clear()
        for i, f in enumerate(state["csvs"]):
            csv_list.controls.append(
                ft.Row(
                    [
                        ft.Icon(
                            ft.Icons.INSERT_DRIVE_FILE,
                            size=16,
                            color=theme_color(page, "primary"),
                        ),
                        ft.Text(
                            f["name"],
                            size=13,
                            expand=True,
                            color=theme_color(page, "text_primary"),
                        ),
                        ft.IconButton(
                            ft.Icons.CLOSE,
                            tooltip="Remove",
                            icon_size=16,
                            on_click=lambda e, idx=i: remove_csv(idx),
                            icon_color=theme_color(page, "text_secondary"),
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                )
            )
        page.update()

    def remove_csv(idx):
        state["csvs"].pop(idx)
        refresh_csvs()

    csv_btn = ft.FilledButton(
        "Add CSVs",
        icon=ft.Icons.LIBRARY_ADD,
        height=40,
        style=ft.ButtonStyle(
            color={ft.ControlState.DEFAULT: ft.Colors.WHITE},
            bgcolor={ft.ControlState.DEFAULT: theme_color(page, "accent")},
            shape=ft.RoundedRectangleBorder(radius=8),
        ),
        on_click=lambda _: csv_picker.pick_files(allow_multiple=True),
    )

    files_box = section(
        ft.Column(
            [
                ft.Text(
                    "🗂 Files",
                    size=15,
                    weight=ft.FontWeight.BOLD,
                    color=theme_color(page, "text_primary"),
                ),
                ft.Text("Google Contacts Source:", weight=ft.FontWeight.BOLD),
                google_source_method,
                csv_controls,
                api_controls,
                ft.Divider(height=10),
                ft.Row(
                    [
                        ft.Text(
                            "🗄️ Additional CSV Files to Merge:",
                            size=15,
                            weight=ft.FontWeight.BOLD,
                            color=theme_color(page, "text_primary"),
                        ),
                        csv_btn,
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                ),
                csv_list,
            ],
            spacing=8,
            expand=True,
        )
    )

    # ---------------------------
    # Options & Actions
    # ---------------------------
    dry = ft.Checkbox(label="Dry-run", value=False)
    log = ft.Checkbox(label="Save log", value=True)
    openf = ft.Checkbox(label="Open folder", value=True)

    # DB connection inputs (optional)
    db_server = ft.TextField(label="DB Server (host[:port])", value="", dense=True)
    db_name = ft.TextField(label="DB Name", value="", dense=True)
    db_user = ft.TextField(label="DB User", value="", dense=True)
    db_pass = ft.TextField(
        label="DB Password",
        value="",
        password=True,
        can_reveal_password=True,
        dense=True,
    )

    # --- Credential Management ---
    KEYRING_SERVICE = "ContactsMerger"

    def save_credentials(e):
        try:
            keyring.set_password(KEYRING_SERVICE, "db_server", db_server.value or "")
            keyring.set_password(KEYRING_SERVICE, "db_name", db_name.value or "")
            keyring.set_password(KEYRING_SERVICE, "db_user", db_user.value or "")
            keyring.set_password(KEYRING_SERVICE, "db_pass", db_pass.value or "")
            page.open(
                ft.SnackBar(
                    ft.Text("✅ Credentials saved securely."),
                    bgcolor=theme_color(page, "success"),
                )
            )
        except Exception as err:
            page.open(
                ft.SnackBar(
                    ft.Text(f"Error saving credentials: {err}"),
                    bgcolor=theme_color(page, "error"),
                )
            )
        page.update()

    def load_credentials():
        try:
            db_server.value = keyring.get_password(KEYRING_SERVICE, "db_server") or ""
            db_name.value = keyring.get_password(KEYRING_SERVICE, "db_name") or ""
            db_user.value = keyring.get_password(KEYRING_SERVICE, "db_user") or ""
            db_pass.value = keyring.get_password(KEYRING_SERVICE, "db_pass") or ""
        except Exception as err:
            print(f"Could not load credentials from store: {err}")
        page.update()

    def delete_credentials(e):
        try:
            keyring.delete_password(KEYRING_SERVICE, "db_server")
            keyring.delete_password(KEYRING_SERVICE, "db_name")
            keyring.delete_password(KEYRING_SERVICE, "db_user")
            keyring.delete_password(KEYRING_SERVICE, "db_pass")
            db_server.value = ""
            db_name.value = ""
            db_user.value = ""
            db_pass.value = ""
            page.open(
                ft.SnackBar(
                    ft.Text("🗑️ Saved credentials cleared."),
                    bgcolor=theme_color(page, "warning"),
                )
            )
        except Exception as err:
            page.open(
                ft.SnackBar(
                    ft.Text(f"Error clearing credentials: {err}"),
                    bgcolor=theme_color(page, "error"),
                )
            )
        page.update()

    # Load credentials on startup
    load_credentials()

    cred_buttons = ft.Row(
        [
            ft.FilledButton(
                "Save Credentials",
                icon=ft.Icons.SAVE,
                on_click=save_credentials,
                style=ft.ButtonStyle(
                    color=ft.Colors.WHITE,
                    bgcolor=theme_color(page, "primary"),
                ),
            ),
            ft.TextButton(
                "Clear Saved",
                icon=ft.Icons.DELETE_FOREVER,
                on_click=delete_credentials,
                style=ft.ButtonStyle(color=theme_color(page, "error")),
            ),
        ],
        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
    )

    options_box = section(
        ft.Column(
            [
                ft.Text(
                    "⚙️ Options",
                    size=15,
                    weight=ft.FontWeight.BOLD,
                    color=theme_color(page, "text_primary"),
                ),
                dry,
                log,
                openf,
                ft.Divider(height=6),
                ft.Text(
                    "MSSQL (optional)",
                    size=12,
                    weight=ft.FontWeight.BOLD,
                    color=theme_color(page, "text_secondary"),
                ),
                db_server,
                db_name,
                db_user,
                db_pass,
                cred_buttons,
            ],
            spacing=4,
            expand=True,
        )
    )

    progress_ring = ft.ProgressRing(width=44, height=44, visible=False)
    progress_text = ft.Text("", size=12, color=theme_color(page, "text_secondary"))

    def run_merge():
        progress_ring.visible = True
        progress_text.value = "Merging..."
        page.update()

        try:
            # --- Load Google Contacts based on selected source ---
            google_contacts = {}
            source_type = state.get("google_source", "csv")
            
            if source_type == "csv":
                google_path = state.get("google_path")
                if not google_path:
                    raise ValueError("Please select a Google Contacts CSV before merging.")
                google_contacts = backend.load_google_contacts(google_path)
            elif source_type == "api":
                google_contacts = state.get("google_contacts")
                if not google_contacts:
                    raise ValueError("Please sync with Google API before merging.")
            else:
                raise ValueError("Invalid Google Contacts source selected.")

            # --- Load MSSQL contacts ---
            mssql_contacts = {}
            for it in state.get("csvs", []):
                path = it.get("path")
                if not path:
                    continue
                try:
                    loaded = backend.load_mssql_contacts(path)
                    for n, d in loaded.items():
                        if n not in mssql_contacts:
                            mssql_contacts[n] = {"numbers": set(), "sources": set()}
                        mssql_contacts[n]["numbers"].update(d.get("numbers", set()))
                        mssql_contacts[n]["sources"].update(d.get("sources", set()))
                        mssql_contacts[n]["first_name"] = d.get("first_name")
                        mssql_contacts[n]["middle_name"] = d.get("middle_name")
                        mssql_contacts[n]["last_name"] = d.get("last_name")
                        mssql_contacts[n]["original_name"] = d.get("original_name")
                except Exception as e:
                    print(f"⚠️ Warning: failed to read '{path}': {e}")

            if (db_server.value or "").strip() and (db_name.value or "").strip():
                try:
                    db_loaded = backend.load_mssql_from_db(
                        (db_server.value or "").strip(),
                        (db_name.value or "").strip(),
                        (db_user.value or "").strip(),
                        (db_pass.value or "") or "",
                    )
                    for n, d in db_loaded.items():
                        if n not in mssql_contacts:
                            mssql_contacts[n] = {"numbers": set(), "sources": set()}
                        mssql_contacts[n]["numbers"].update(d.get("numbers", set()))
                        mssql_contacts[n]["sources"].update(d.get("sources", set()))
                        mssql_contacts[n]["first_name"] = d.get("first_name")
                        mssql_contacts[n]["middle_name"] = d.get("middle_name")
                        mssql_contacts[n]["last_name"] = d.get("last_name")
                        mssql_contacts[n]["original_name"] = d.get("original_name")
                except Exception as e:
                    print(f"⚠️ Warning: failed to load from DB: {e}")

            # --- Merge ---
            merge_result = backend.merge_contacts(google_contacts, mssql_contacts)
            if isinstance(merge_result, tuple) or isinstance(merge_result, list):
                merged = merge_result[0]
                per_row_logs = merge_result[1] if len(merge_result) > 1 else []
            else:
                merged = merge_result
                per_row_logs = getattr(backend, "LAST_PER_ROW_LOGS", [])

            # --- Export ---
            # Determine output folder. Use CWD if path is not available (API case)
            google_path = state.get("google_path")
            base_dir = os.path.dirname(google_path) if google_path else os.getcwd()

            output_dir = os.path.join(base_dir, "output")
            os.makedirs(output_dir, exist_ok=True)
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(output_dir, f"merged_contacts_{ts}.csv")

            if not dry.value:
                # Use template if available, otherwise the backend handles defaults
                template_path = state.get("google_path")
                backend.export_contacts(
                    merged, output_file=output_file, template=template_path
                )

            # --- Logging and Summary ---
            if log.value:
                logp = os.path.join(output_dir, f"merge_log_{ts}.txt")
                with open(logp, "w", encoding="utf-8") as fh:
                    fh.write(f"Merge run: {ts}\n")
                    google_file_info = state.get("google_path") or "Google API Sync"
                    fh.write(f"Google source: {google_file_info}\n")
                    fh.write(
                        f"MSSQL files: {[x['path'] for x in state.get('csvs', [])]}\n"
                    )
                    fh.write(f"Total merged contacts: {len(merged)}\n\n")

                    if per_row_logs:
                        fh.write("DETAILED ROW UPDATES:\n")
                        for rec in per_row_logs:
                            fh.write("---\n")
                            fh.write("1) Original Google row:\n")
                            og = rec.get("original_google_row")
                            fh.write(str(og) + "\n" if og else "<no original row preserved>\n")
                            fh.write("2) Update data applied:\n")
                            fh.write(str(rec.get("update_data", {})) + "\n")
                            fh.write("3) Final row in output:\n")
                            fh.write(str(rec.get("final_row", {})) + "\n")
                        fh.write("---\n")
                    else:
                        fh.write("No detailed per-row logs available.\n")
                print(f"✅ Log saved to {logp}")

            if openf.value and os.path.exists(output_dir):
                try:
                    os.startfile(output_dir)
                except Exception:
                    pass

            data = {}
            data["Google"] = len(google_contacts)
            data["MSSQL"] = len(mssql_contacts)
            data["Total"] = len(merged)
            def _sources_is_only_mssql(v):
                s = set(v.get("sources") or [])
                return s == {"MSSQL"}
            def _sources_have_both(v):
                s = set(v.get("sources") or [])
                return ("Google" in s) and ("MSSQL" in s)
            data["New"] = sum(1 for v in merged.values() if _sources_is_only_mssql(v))
            data["Merged"] = sum(1 for v in merged.values() if _sources_have_both(v))
            data["Protected"] = sum(1 for v in merged.values() if v.get("protected"))

            progress_ring.visible = False
            progress_text.value = ""
            show_summary(data)
            page.open(
                ft.SnackBar(
                    ft.Text("✅ Merge complete!"), bgcolor=theme_color(page, "success")
                )
            )
            page.update()

        except Exception as err:
            progress_ring.visible = False
            progress_text.value = ""
            page.open(
                ft.SnackBar(ft.Text(str(err)), bgcolor=theme_color(page, "error"))
            )
            page.update()

    start_btn = ft.FilledButton(
        "Start Merge",
        icon=ft.Icons.PLAY_ARROW,
        height=40,
        style=ft.ButtonStyle(
            color={ft.ControlState.DEFAULT: ft.Colors.WHITE},
            bgcolor={ft.ControlState.DEFAULT: theme_color(page, "success")},
            shape=ft.RoundedRectangleBorder(radius=8),
        ),
        on_click=lambda e: threading.Thread(target=run_merge, daemon=True).start(),
    )

    def handle_exit(e):
        print("Exit button pressed. Forcing application shutdown via taskkill.")
        os.system(f'taskkill /F /PID {os.getpid()} /T')

    exit_btn = ft.FilledButton(
        "Exit",
        icon=ft.Icons.CLOSE,
        height=40,
        style=ft.ButtonStyle(
            color={ft.ControlState.DEFAULT: ft.Colors.WHITE},
            bgcolor={ft.ControlState.DEFAULT: theme_color(page, "error")},
            shape=ft.RoundedRectangleBorder(radius=8),
        ),
        on_click=handle_exit,
    )

    actions_box = section(
        ft.Column(
            [
                ft.Text(
                    "▶️ Actions",
                    size=15,
                    weight=ft.FontWeight.BOLD,
                    color=theme_color(page, "text_primary"),
                ),
                ft.Row(
                    [start_btn, exit_btn], alignment=ft.MainAxisAlignment.SPACE_EVENLY
                ),
                ft.Row(
                    [progress_ring, progress_text],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
            ],
            spacing=8,
            expand=True,
        )
    )

    right_column = ft.Column(
        [options_box, ft.Container(height=8), actions_box], expand=True
    )
    left_column = ft.Container(files_box, expand=True)
    top_section = ft.Row(
        [
            ft.Container(left_column, expand=True),
            ft.Container(width=12),
            ft.Container(right_column, expand=True),
        ],
        expand=True,
    )

    # ---------------------------
    # Summary (reactive)
    # ---------------------------
    summary = ft.Container(visible=False, expand=True)
    boxes.append(summary)

    def show_summary(data=None):
        if data is None:
            data = {
                "Google": 1700,
                "MSSQL": len(state["csvs"]),
                "Total": 5400,
                "New": 2700,
                "Merged": 600,
                "Protected": 40,
            }

        summary_metrics.clear()

        def metric(icon, label, val, color):
            box = ft.Container(
                ft.Row(
                    [
                        ft.Icon(icon, color=color, size=18),
                        ft.Column(
                            [
                                ft.Text(
                                    label,
                                    size=12,
                                    color=theme_color(page, "text_primary"),
                                ),
                                ft.Text(
                                    str(val),
                                    size=14,
                                    weight=ft.FontWeight.BOLD,
                                    color=color,
                                ),
                            ],
                            spacing=0,
                        ),
                    ],
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=6,
                ),
                bgcolor=theme_color(page, "metric_bg"),
                border_radius=8,
                padding=8,
                expand=True,
            )
            summary_metrics.append(box)
            return box

        metrics_row = ft.Row(
            [
                metric(
                    ft.Icons.PERSON,
                    "Google",
                    data["Google"],
                    theme_color(page, "primary"),
                ),
                metric(
                    ft.Icons.LIST, "MSSQL", data["MSSQL"], theme_color(page, "accent")
                ),
                metric(
                    ft.Icons.INSERT_DRIVE_FILE,
                    "Total",
                    data["Total"],
                    theme_color(page, "primary_dark"),
                ),
                metric(
                    ft.Icons.PERSON_ADD,
                    "New",
                    data["New"],
                    theme_color(page, "success"),
                ),
                metric(
                    ft.Icons.SYNC_ALT,
                    "Merged",
                    data["Merged"],
                    theme_color(page, "warning"),
                ),
                metric(
                    ft.Icons.LOCK,
                    "Protected",
                    data["Protected"],
                    theme_color(page, "error"),
                ),
            ],
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            spacing=10,
        )

        summary.content = ft.Container(
            ft.Column(
                [
                    ft.Text(
                        "🧾 Summary",
                        size=15,
                        weight=ft.FontWeight.BOLD,
                        color=theme_color(page, "text_primary"),
                    ),
                    metrics_row,
                ],
                spacing=10,
            ),
            bgcolor=theme_color(page, "surface"),
            border_radius=10,
            padding=12,
            shadow=ft.BoxShadow(
                blur_radius=8, color=ft.Colors.BLACK12, offset=ft.Offset(0, 3)
            ),
        )
        summary.visible = True
        page.update()

    # ---------------------------
    # Refresh Theme (dynamic recoloring)
    # ---------------------------
    def refresh_theme():
        for box in boxes:
            if isinstance(box, ft.Container):
                box.bgcolor = theme_color(page, "surface")

        if summary.visible and summary.content:
            # Only set bgcolor if the content is actually a Container (narrow the type)
            if isinstance(summary.content, ft.Container):
                summary.content.bgcolor = theme_color(page, "surface")
            for m in summary_metrics:
                m.bgcolor = theme_color(page, "metric_bg")
                for row_item in m.content.controls:
                    if isinstance(row_item, ft.Column):
                        for txt in row_item.controls:
                            if isinstance(txt, ft.Text):
                                txt.color = theme_color(page, "text_primary")

    # ---------------------------
    # Layout
    # ---------------------------
    layout = ft.Column(
        [header, top_section, ft.Container(height=10), summary], spacing=10, expand=True
    )
    page.add(layout)
    page.update()


if __name__ == "__main__":
    ft.app(target=main)
