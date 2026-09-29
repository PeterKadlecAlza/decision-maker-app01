from __future__ import annotations

import json
import hashlib
import hmac
import os
import socket
import re
import unicodedata
from datetime import datetime
from io import BytesIO
from html import unescape
from urllib.parse import quote_plus, urljoin
from urllib.request import Request, urlopen
from typing import Iterable

import pandas as pd
import streamlit as st

from processing import clean_code_value, export_workbook, run_pipeline
from storage import backup_database, initialize_storage, latest_batch_key, load_latest_pipeline_result, load_reviews, save_pipeline_result, save_review, storage_location_label, update_batch_tables, uses_persistent_cloud_storage


st.set_page_config(
    page_title="Packaging Instruction Enrichment & Decision Maker",
    layout="wide",
)


REVIEW_STATUS_OPTIONS = [
    "Neriešené",
    "Rieši sa",
    "Čaká na IT",
    "Schválené",
    "Hotovo",
]

FINAL_ACTION_OPTIONS = [
    "Ponechať balenie",
    "Odstrániť balenie",
    "Upraviť parameter",
    "Odovzdať na IT",
    "Manuálna analýza",
]

DISPLAY_CASE_COLUMNS = [
    "case_id",
    "priority",
    "Kód produktu",
    "Název produktu",
    "instruction_category",
    "report_count",
    "recommended_action",
    "confidence",
    "status",
    "last_updated_by",
]

REVIEW_EXPORT_COLUMNS = [
    "case_id",
    "Kód produktu",
    "Název produktu",
    "instruction_category",
    "report_count",
    "recommended_action",
    "suggested_parameter_to_check",
    "confidence",
    "review_status",
    "final_action",
    "validator_note",
    "decision_reason",
    "validator_checklist",
]


def read_excel_upload(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()
    return pd.read_excel(uploaded_file, dtype=object, engine="openpyxl")


PRODUCT_DESCRIPTION_COLUMNS = [
    "Popis produktu",
    "Produktový popis",
    "Ext. jméno",
    "Externí jméno",
    "External name",
    "Description",
    "Long description",
]


def enrich_current_batch_with_descriptions(result, descriptions_raw: pd.DataFrame) -> dict[str, int]:
    """Merge descriptions into the current batch while preserving cases and reviews."""
    code_column = _resolve_column(descriptions_raw, "Kód produktu", "SKU")
    description_column = _resolve_column(descriptions_raw, *PRODUCT_DESCRIPTION_COLUMNS)
    if code_column is None:
        raise ValueError("V Exceli chýba stĺpec Kód produktu.")
    if description_column is None:
        raise ValueError("V Exceli chýba stĺpec Ext. jméno alebo Popis produktu.")

    source = descriptions_raw.loc[:, [code_column, description_column]].copy()
    source["_code"] = source[code_column].map(clean_code_value)
    source["_description"] = source[description_column].map(
        lambda value: None if pd.isna(value) else str(value).strip()
    )
    source = source.dropna(subset=["_code", "_description"])
    source = source[source["_description"].ne("")]
    source = source.drop_duplicates(subset=["_code"], keep="last")
    description_map = dict(zip(source["_code"].astype(str), source["_description"]))
    if not description_map:
        raise ValueError("Excel neobsahuje žiadne použiteľné kódy a popisy produktov.")

    matched_product_codes: set[str] = set()
    matched_cases = 0
    for table_name in ("products_clean", "cases_enriched", "ai_cases_input"):
        frame = getattr(result, table_name).copy()
        target_code_column = _resolve_column(frame, "Kód produktu", "SKU")
        if target_code_column is None:
            continue
        normalized_codes = frame[target_code_column].map(clean_code_value).astype("string")
        matched_mask = normalized_codes.isin(description_map)
        frame.loc[matched_mask, "Popis produktu"] = normalized_codes[matched_mask].map(description_map)
        setattr(result, table_name, frame)
        if table_name == "products_clean":
            matched_product_codes.update(normalized_codes[matched_mask].dropna().astype(str))
        elif table_name == "cases_enriched":
            matched_cases = int(matched_mask.sum())

    return {
        "input_descriptions": len(description_map),
        "matched_products": len(matched_product_codes),
        "matched_cases": matched_cases,
        "unmatched_products": len(set(description_map) - matched_product_codes),
    }


def _input_signature(reports_file, products_file) -> str:
    digest = hashlib.sha256()
    for uploaded_file in (reports_file, products_file):
        digest.update(uploaded_file.getvalue())
    digest.update(str(st.session_state.get("top_n_input", 100)).encode("utf-8"))
    digest.update(str(bool(st.session_state.get("exclude_stitok_input", True))).encode("utf-8"))
    digest.update(str(st.session_state.get("min_report_count_input", 0)).encode("utf-8"))
    return digest.hexdigest()


@st.cache_data(show_spinner=False)
def load_latest_pipeline_result_cached(batch_key):
    if batch_key is None:
        return None, None
    return load_latest_pipeline_result()


def default_select_all(options: Iterable) -> list:
    return list(options)


def init_session_state() -> None:
    initialize_storage()
    st.session_state.setdefault("review_data", {})
    st.session_state.setdefault("current_batch_id", None)
    st.session_state.setdefault("current_input_signature", None)
    st.session_state.setdefault("current_result", None)
    st.session_state.setdefault("access_role", None)
    st.session_state.setdefault("access_identity", "")
    st.session_state.setdefault("reviewer_link_authorized", False)
    st.session_state.setdefault("processed_export_ready_batch_id", None)
    st.session_state.setdefault("review_list_page_size", 20)
    st.session_state.setdefault("review_visible_case_count", 20)
    st.session_state.setdefault("review_case_filter_signature", None)
    st.session_state.setdefault("selected_case_id", None)
    st.session_state.setdefault("review_widget_case_id", None)
    st.session_state.setdefault("review_status_widget", REVIEW_STATUS_OPTIONS[0])
    st.session_state.setdefault("final_action_widget", FINAL_ACTION_OPTIONS[0])
    st.session_state.setdefault("validator_note_widget", "")
    st.session_state.setdefault("session_saved_case_ids", set())


def render_access_login() -> str | None:
    admin_password = os.environ.get("INSTRUCTION_VALIDATOR_ADMIN_PASSWORD", "")
    reviewer_token = os.environ.get("INSTRUCTION_VALIDATOR_REVIEWER_TOKEN", "")
    st.sidebar.markdown("### Prihlásenie")
    role = st.session_state.get("access_role")
    if role in {"admin", "reviewer"}:
        identity = st.session_state.get("access_identity") or "neznámy používateľ"
        role_label = "Admin" if role == "admin" else "Kvalitár"
        st.sidebar.success(f"{role_label}: {identity}")
        if st.sidebar.button("Odhlásiť"):
            st.session_state.access_role = None
            st.session_state.access_identity = ""
            st.session_state.reviewer_link_authorized = False
            st.rerun()
        return role

    supplied_token = str(st.query_params.get("reviewer_token", "") or "")
    if supplied_token:
        if reviewer_token and hmac.compare_digest(supplied_token, reviewer_token):
            st.session_state.reviewer_link_authorized = True
            st.query_params.clear()
        else:
            st.query_params.clear()
            st.sidebar.error("Prístupový link nie je platný.")

    if st.session_state.get("reviewer_link_authorized"):
        st.sidebar.success("Prístup kvalitára bol overený.")
        identity = st.sidebar.text_input("Meno kvalitára", key="reviewer_identity")
        if st.sidebar.button("Pokračovať ako kvalitár"):
            clean_identity = str(identity or "").strip()
            if not clean_identity:
                st.sidebar.error("Zadajte svoje meno.")
                return None
            st.session_state.access_role = "reviewer"
            st.session_state.access_identity = clean_identity
            st.rerun()
        return None

    if not admin_password:
        st.sidebar.warning("Admin prístup nie je nakonfigurovaný.")
        return None
    st.sidebar.caption("Admin prístup")
    identity = st.sidebar.text_input("Meno admina", key="login_identity")
    password = st.sidebar.text_input("Admin heslo", type="password", key="login_password")
    if st.sidebar.button("Prihlásiť admina"):
        clean_identity = str(identity or "").strip()
        if not clean_identity:
            st.sidebar.error("Zadajte meno admina.")
            return None
        if hmac.compare_digest(password, admin_password):
            st.session_state.access_role = "admin"
            st.session_state.access_identity = clean_identity
            st.rerun()
        else:
            st.sidebar.error("Nesprávne admin heslo.")
    if not reviewer_token:
        st.sidebar.caption("Prístupový link kvalitára zatiaľ nie je nakonfigurovaný.")
    return None


NEW_REVIEW_STATUS_OPTIONS = [
    "Nový",
    "V analýze",
    "Čaká na zmenu v internom systéme",
    "Zmena vykonaná",
    "Zamietnuté",
    "Nie je možné rozhodnúť",
]

LEGACY_REVIEW_STATUS_OPTIONS = [
    "Neriešené",
    "Rieši sa",
    "Čaká na IT",
    "Schválené",
    "Hotovo",
]

REVIEW_STATUS_OPTIONS_V2 = NEW_REVIEW_STATUS_OPTIONS + [
    status for status in LEGACY_REVIEW_STATUS_OPTIONS if status not in NEW_REVIEW_STATUS_OPTIONS
]

FINAL_ACTION_OPTIONS_V2 = [
    "Iba štítok",
    "Prelepiť páskou",
    "Zabaliť do fólie",
    "Použiť bublinkovú fóliu",
    "Použiť obálku",
    "Použiť kartónovú krabičku",
    "Ponechať aktuálnu inštrukciu",
    "Vyžaduje ďalšiu analýzu",
    "Ponechať balenie",
    "Odstrániť balenie",
    "Upraviť parameter",
    "Odovzdať na IT",
    "Manuálna analýza",
]

EXTENDED_REVIEW_EXPORT_COLUMNS = [
    "case_id",
    "SKU",
    "Product ID",
    "Kód produktu",
    "Názov produktu",
    "original_instruction",
    "final_instruction",
    "status",
    "review_status",
    "final_action",
    "review_saved_at",
    "review_saved_by",
    "Validator note",
    "validator_note",
    "change_execution_status",
    "change_executed_at",
    "change_executed_by",
    "instruction_category",
    "report_count",
    "recommended_action",
    "suggested_parameter_to_check",
    "confidence",
    "decision_reason",
    "validator_checklist",
]

INTERNAL_CONSOLE_URL_TEMPLATE = ""
# TODO: Doplniť URL šablónu internej produktovej konzoly


def safe_value(row: pd.Series | None, column_name: str, default="Neuvedené"):
    if row is None or column_name not in row.index:
        return default
    value = row[column_name]
    if pd.isna(value):
        return default
    if isinstance(value, str):
        cleaned = value.strip()
        return cleaned if cleaned else default
    return value


def safe_text_value(row: pd.Series | None, column_name: str, default="Neuvedené") -> str:
    value = safe_value(row, column_name, default=default)
    if value == default:
        return default
    text = str(value).strip()
    return text if text else default


def _row_column_name(row: pd.Series, *candidates: str) -> str | None:
    normalized = {_normalize_for_matching(str(column)): column for column in row.index}
    for candidate in candidates:
        resolved = normalized.get(_normalize_for_matching(candidate))
        if resolved is not None:
            return resolved
    return None


def _row_safe_value(row: pd.Series, *candidates: str, default="Neuvedené"):
    column = _row_column_name(row, *candidates)
    if column is None:
        return default
    return safe_value(row, column, default=default)


def build_console_action_guidance(case_row: pd.Series, review_state: dict | None = None) -> dict:
    """Build a conservative, human-checkable checklist for the internal console."""
    review_state = review_state or {}
    action = str(
        review_state.get("final_action")
        or case_row.get("recommended_action")
        or "Manuálna analýza"
    ).strip()
    product_description = _row_safe_value(
        case_row,
        "Popis produktu",
        "Produktový popis",
        "Ext. jméno",
        "Externí jméno",
        "External name",
        "Description",
        "Long description",
        default="",
    )
    description_available = product_description not in {"", "Neuvedené"}

    if action in {"Ponechať balenie", "Ponechať aktuálnu inštrukciu"}:
        steps = [("Žiadny parameter", "Nič nemeniť", "Odporúčanie ponecháva aktuálne nastavenie.")]
        confidence = "vysoká"
    elif action in {"Použiť bublinkovú fóliu", "Zabaliť do bublinkovej fólie"}:
        steps = [
            (
                "Bublinková fólia (Skladová vlastnost)",
                "Nastaviť na Áno",
                "Hlavný parameter pre navrhované balenie.",
            ),
            (
                "Křehký produkt / Sklo / Porcelán / Lehko poškoditelný",
                "Overiť podľa produktu; nemeníť automaticky",
                "Pomocné rizikové parametre majú zostať predmetom kontroly človeka.",
            ),
        ]
        confidence = "stredná"
    elif action in {"Zabaliť do fólie", "Prelepiť páskou"}:
        steps = [
            (
                "Fólie (Skladová vlastnost)",
                "Overiť alebo nastaviť podľa požadovaného spôsobu balenia",
                "Parameter zodpovedá fóliovému baleniu.",
            ),
            (
                "Dobalovat (Skladová vlastnost)",
                "Overiť; nemeníť automaticky",
                "Treba potvrdiť, či má ísť o dodatočné balenie.",
            ),
        ]
        confidence = "stredná"
    elif action in {"Odstrániť balenie", "Preveriť odstránenie balenia"}:
        steps = [
            (
                "Nebaliť (Skladová vlastnost)",
                "Overiť alebo nastaviť na Áno po manuálnej kontrole",
                "Odstránenie balenia môže byť v konflikte s rizikovými parametrami.",
            ),
            (
                "Aktívny baliaci parameter",
                "Identifikovať a vypnúť iba potvrdený trigger",
                "Aplikácia zatiaľ neurčuje konkrétny trigger bez potvrdenia v konzole.",
            ),
        ]
        confidence = "nízka až stredná"
    else:
        steps = [
            (
                "Konzola produktu",
                "Vykonať manuálnu analýzu",
                "Pre túto akciu zatiaľ nemáme dostatočne bezpečné pravidlo.",
            )
        ]
        confidence = "nízka"

    return {
        "action": action,
        "steps": pd.DataFrame(steps, columns=["Parameter / oblasť", "Odporúčaný krok", "Dôvod"]),
        "confidence": confidence,
        "product_description": product_description if description_available else "",
        "description_available": description_available,
    }


def _current_timestamp() -> str:
    return datetime.now().astimezone().replace(microsecond=0).isoformat()


def _current_user_identifier() -> str:
    for key in ["access_identity", "current_user", "username", "user", "user_email", "email", "name"]:
        value = st.session_state.get(key)
        if value:
            return str(value)
    user_obj = getattr(st, "user", None)
    for attr in ["email", "name", "username", "id"]:
        value = getattr(user_obj, attr, None)
        if value:
            return str(value)
    return socket.gethostname()


def _mark_case_saved_for_session(case_id: str) -> None:
    saved_case_ids = st.session_state.setdefault("session_saved_case_ids", set())
    saved_case_ids.add(str(case_id))


def _review_has_saved_work(review: dict | None) -> bool:
    """Identify cases with an explicit persisted review without hiding untouched cases."""
    if not review:
        return False
    return bool(
        review.get("review_saved_at")
        or review.get("review_saved_by")
        or review.get("change_executed_at")
        or review.get("change_executed_by")
        or review.get("change_execution_status") == "Zmena vykonaná"
        or review.get("review_status") in {"Zmena vykonaná", "Zamietnuté", "Nie je možné rozhodnúť"}
    )


def _normalize_review_entry(entry: dict | None, row: pd.Series | None = None) -> dict:
    recommended_action = safe_text_value(row, "recommended_action", default="") if row is not None else ""
    default_status = "Nový"
    if recommended_action == "Odovzdať na IT":
        default_status = "Čaká na zmenu v internom systéme"
    elif recommended_action in {"Ponechať balenie", "Ponechať aktuálnu inštrukciu"}:
        default_status = "V analýze"
    elif recommended_action in {"Odstrániť balenie", "Prelepiť páskou", "Zabaliť do fólie"}:
        default_status = "V analýze"

    default_final_action = recommended_action if recommended_action in FINAL_ACTION_OPTIONS_V2 else "Vyžaduje ďalšiu analýzu"
    normalized = {
        "review_status": default_status,
        "status": default_status,
        "final_action": default_final_action,
        "validator_note": "",
        "review_saved_at": "",
        "review_saved_by": "",
        "change_execution_status": "",
        "change_executed_at": "",
        "change_executed_by": "",
    }
    if entry:
        normalized.update(entry)
    normalized.setdefault("status", normalized.get("review_status", default_status))
    normalized.setdefault("review_status", normalized.get("status", default_status))
    normalized.setdefault("final_action", default_final_action)
    normalized.setdefault("validator_note", "")
    normalized.setdefault("review_saved_at", "")
    normalized.setdefault("review_saved_by", "")
    normalized.setdefault("change_execution_status", "")
    normalized.setdefault("change_executed_at", "")
    normalized.setdefault("change_executed_by", "")
    return normalized


def ensure_optional_review_columns(df: pd.DataFrame) -> pd.DataFrame:
    optional_columns = {
        "SKU": "",
        "Product ID": "",
        "original_instruction": "",
        "final_instruction": "",
        "status": "",
        "review_status": "",
        "final_action": "",
        "review_saved_at": "",
        "review_saved_by": "",
        "Validator note": "",
        "validator_note": "",
        "change_execution_status": "",
        "change_executed_at": "",
        "change_executed_by": "",
    }
    for column, default_value in optional_columns.items():
        if column not in df.columns:
            df[column] = default_value
    return df


def build_internal_console_url(product_id=None, sku=None):
    template = INTERNAL_CONSOLE_URL_TEMPLATE.strip()
    if not template:
        return None
    if product_id in [None, ""] and sku in [None, ""]:
        return None
    values = {}
    if product_id not in [None, ""]:
        values["product_id"] = quote_plus(str(product_id))
    if sku not in [None, ""]:
        values["sku"] = quote_plus(str(sku))
    try:
        class _BlankMapping(dict):
            def __missing__(self, key):
                return ""

        return template.format_map(_BlankMapping(values))
    except Exception:
        return None


def update_review_execution_fields(review_entry: dict, executed_by: str | None = None, executed_at: str | None = None) -> dict:
    updated = dict(review_entry)
    timestamp = executed_at or _current_timestamp()
    updated["review_status"] = "Zmena vykonaná"
    updated["status"] = "Zmena vykonaná"
    updated["change_execution_status"] = "Zmena vykonaná"
    updated["change_executed_at"] = timestamp
    updated["change_executed_by"] = executed_by or ""
    return updated


def _is_valid_http_url(value: str) -> bool:
    return bool(re.match(r"^https?://", value.strip(), flags=re.I))


def _image_column_name(row: pd.Series) -> str | None:
    preferred_tokens = ["image", "img", "photo", "foto", "fotka", "picture", "obraz"]
    for column in row.index:
        normalized = _normalize_for_matching(str(column))
        if any(token in normalized for token in preferred_tokens):
            return column
    return None


def _display_placeholder_image() -> None:
    st.info("Fotografia produktu nie je k dispozícii.")


def render_product_image(row: pd.Series) -> None:
    image_column = _image_column_name(row)
    if image_column is None:
        # TODO: Napojenie na finálny zdroj produktových fotografií
        _display_placeholder_image()
        return

    image_url = safe_text_value(row, image_column, default="")
    if not image_url or not _is_valid_http_url(image_url):
        _display_placeholder_image()
        return

    try:
        request = Request(
            image_url,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Accept": "image/*,*/*;q=0.8",
            },
        )
        with urlopen(request, timeout=10) as response:
            content = response.read()
        st.image(BytesIO(content), caption="Fotografia produktu", use_container_width=True)
    except Exception:
        _display_placeholder_image()


def render_product_header(case_row: pd.Series) -> None:
    st.markdown("### Produkt")
    left, right = st.columns([1, 2])
    with left:
        render_product_image(case_row)
    with right:
        product_name = _row_safe_value(case_row, "Názov produktu", "Název produktu")
        sku = _row_safe_value(case_row, "Kód produktu")
        product_id = _row_safe_value(case_row, "Product ID", "SEOPrefix_ID", "ID produktu")
        category = _row_safe_value(case_row, "Segment1", "Kategória", "Kategorie")
        product_description = _row_safe_value(
            case_row,
            "Popis produktu",
            "Produktový popis",
            "Ext. jméno",
            "Externí jméno",
            "External name",
            "Description",
            "Long description",
        )
        current_instruction = _row_safe_value(
            case_row,
            "Přeložené instrukce",
            "Peložené instrukce",
            "Přeložená instrukce",
        )
        details = pd.DataFrame(
            [
                {
                    "Názov produktu": product_name,
                    "SKU": sku,
                    "Product ID": product_id,
                    "Kategória produktu": category,
                    "Popis produktu": product_description,
                    "Aktuálna baliaca inštrukcia": current_instruction,
                }
            ]
        )
        st.dataframe(details, use_container_width=True, hide_index=True)

        console_url = build_internal_console_url(
            product_id=None if product_id == "Neuvedené" else product_id,
            sku=None if sku == "Neuvedené" else sku,
        )
        if console_url:
            if hasattr(st, "link_button"):
                st.link_button("Otvoriť produkt v konzole", console_url)
            else:
                st.markdown(f"[Otvoriť produkt v konzole]({console_url})")
        else:
            st.button("Otvoriť produkt v konzole", disabled=True)
            st.caption("Konzola zatiaľ nie je nakonfigurovaná.")


def render_case_reason(case_row: pd.Series) -> None:
    st.markdown("### Dôvod kontroly")
    reason_rows = [
        ("Počet hlásení", _row_safe_value(case_row, "report_count")),
        ("Počet používateľov", _row_safe_value(case_row, "unique_users_count")),
        ("Prvé hlásenie", _row_safe_value(case_row, "first_reported_at")),
        ("Posledné hlásenie", _row_safe_value(case_row, "last_reported_at")),
        ("Hlavný dôvod kontroly", _row_safe_value(case_row, "main_trigger_parameter")),
        ("Identifikované rizikové vlastnosti", _row_safe_value(case_row, "active_risk_parameters")),
        ("Vyhodnotenie rizika", _row_safe_value(case_row, "risk_evaluation")),
        ("AI odporúčanie", _row_safe_value(case_row, "ai_decision_hint")),
        ("Odporúčané opatrenie", _row_safe_value(case_row, "recommended_action")),
        ("Zdôvodnenie rozhodnutia", _row_safe_value(case_row, "decision_reason")),
    ]
    reason_frame = pd.DataFrame(reason_rows, columns=["Pole", "Hodnota"])
    reason_frame["Hodnota"] = reason_frame["Hodnota"].map(str)
    st.dataframe(reason_frame, use_container_width=True, hide_index=True)


def render_instruction_comparison(case_row: pd.Series) -> None:
    st.markdown("### Porovnanie inštrukcií")
    current_instruction = _row_safe_value(case_row, "Přeložené instrukce", "Peložené instrukce", "Přeložená instrukce")
    recommended_instruction = _row_safe_value(case_row, "recommended_action")
    left, right = st.columns(2)
    with left:
        st.markdown("**Aktuálna inštrukcia**")
        st.write(current_instruction)
    with right:
        st.markdown("**Odporúčaná inštrukcia**")
        st.write(recommended_instruction)


def _review_widget_key(case_id: str, field: str) -> str:
    return f"{field}_{case_id}"


def _selectbox_options_with_current(current_value: str, base_options: list[str]) -> list[str]:
    if current_value in base_options:
        return base_options
    return [current_value] + base_options


def render_review_form(case_id: str, case_row: pd.Series) -> None:
    st.markdown("### Review detail")
    review_state = _normalize_review_entry(st.session_state.review_data.get(case_id), case_row)
    st.session_state.review_data[case_id] = review_state

    status_key = _review_widget_key(case_id, "status")
    final_action_key = _review_widget_key(case_id, "final_action")
    validator_note_key = _review_widget_key(case_id, "validator_note")
    status_options = _selectbox_options_with_current(review_state["review_status"], REVIEW_STATUS_OPTIONS_V2)
    final_action_options = _selectbox_options_with_current(review_state["final_action"], FINAL_ACTION_OPTIONS_V2)

    st.session_state.setdefault(status_key, review_state["review_status"])
    st.session_state.setdefault(final_action_key, review_state["final_action"])
    st.session_state.setdefault(validator_note_key, review_state["validator_note"])

    st.selectbox("Status", options=status_options, key=status_key)
    st.selectbox("Final action", options=final_action_options, key=final_action_key)
    st.text_area("Validator note", key=validator_note_key, height=180)

    last_updated_by = review_state.get("change_executed_by") or review_state.get("review_saved_by")
    last_updated_at = review_state.get("change_executed_at") or review_state.get("review_saved_at")
    if last_updated_by:
        timestamp = f" ({last_updated_at})" if last_updated_at else ""
        st.caption(f"Posledný update: {last_updated_by}{timestamp}")

    cols = st.columns([1, 1])
    with cols[0]:
        if st.button("Save review", key=f"save_review_{case_id}"):
            save_current_review(case_id, case_row, explicit_save=True)
            st.success("Rozhodnutie bolo uložené.")
    with cols[1]:
        change_state = st.session_state.review_data.get(case_id, {})
        can_confirm = bool(change_state.get("review_saved_at"))
        already_executed = change_state.get("change_execution_status") == "Zmena vykonaná"
        confirm_key = _review_widget_key(case_id, "confirm_change")
        if already_executed:
            st.info("Zmena už bola označená ako vykonaná.")
        elif can_confirm:
            if st.button("Potvrdiť vykonanie zmeny", key=f"confirm_change_{case_id}"):
                st.session_state[confirm_key] = True
            if st.session_state.get(confirm_key):
                st.warning("Toto označí prípad ako vybavený. Pokračovať?")
                confirm_cols = st.columns(2)
                with confirm_cols[0]:
                    if st.button("Áno, potvrdiť", key=f"confirm_change_yes_{case_id}"):
                        mark_change_as_executed(case_id)
                        st.session_state[confirm_key] = False
                        st.success("Zmena bola označená ako vykonaná.")
                with confirm_cols[1]:
                    if st.button("Zrušiť", key=f"confirm_change_no_{case_id}"):
                        st.session_state[confirm_key] = False
        else:
            st.button("Potvrdiť vykonanie zmeny", disabled=True, key=f"confirm_change_disabled_{case_id}")
            st.caption("Najprv uložte review, až potom je možné potvrdiť vykonanie zmeny.")

    save_current_review(case_id, case_row, explicit_save=False)


def save_current_review(case_id: str, case_row: pd.Series, explicit_save: bool = False) -> None:
    review_entry = _normalize_review_entry(st.session_state.review_data.get(case_id), case_row)
    status_key = _review_widget_key(case_id, "status")
    final_action_key = _review_widget_key(case_id, "final_action")
    validator_note_key = _review_widget_key(case_id, "validator_note")
    review_entry["review_status"] = st.session_state.get(status_key, review_entry["review_status"])
    review_entry["status"] = review_entry["review_status"]
    review_entry["final_action"] = st.session_state.get(final_action_key, review_entry["final_action"])
    review_entry["validator_note"] = str(st.session_state.get(validator_note_key, review_entry["validator_note"]) or "").strip()
    if explicit_save:
        review_entry["review_saved_at"] = _current_timestamp()
        review_entry["review_saved_by"] = _current_user_identifier()
    st.session_state.review_data[case_id] = review_entry
    save_review(st.session_state.get("current_batch_id"), case_id, review_entry)


def mark_change_as_executed(case_id: str) -> None:
    if case_id not in st.session_state.review_data:
        return
    review_entry = _normalize_review_entry(st.session_state.review_data.get(case_id))
    updated = update_review_execution_fields(review_entry, executed_by=_current_user_identifier())
    st.session_state.review_data[case_id] = updated
    st.session_state[_review_widget_key(case_id, "status")] = updated["review_status"]
    save_review(st.session_state.get("current_batch_id"), case_id, updated)


ALZA_BASE_URL = "https://www.alza.sk"
ALZA_SEARCH_URL = f"{ALZA_BASE_URL}/search.htm?exps={{query}}"
ALZA_TIMEOUT_SECONDS = 20

PACKAGING_KEYWORDS = [
    "obal",
    "baleni",
    "balenie",
    "krabic",
    "krabica",
    "box",
    "foli",
    "vypln",
    "poskoz",
    "rozbit",
    "prask",
    "zdeform",
    "chybajuc",
    "chybajuci",
    "prislusenst",
]

QUALITY_KEYWORDS = [
    "nefung",
    "vada",
    "porucha",
    "reklam",
    "zavad",
    "poskoden",
    "pokaz",
    "nevyhov",
]


def _strip_accents(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))


def _normalize_for_matching(value: str | None) -> str:
    if not value:
        return ""
    text = _strip_accents(str(value)).lower()
    text = re.sub(r"[^a-z0-9%]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _resolve_column(df: pd.DataFrame, *candidates: str) -> str | None:
    if df is None or df.empty:
        return None
    normalized = { _normalize_for_matching(column): column for column in df.columns }
    for candidate in candidates:
        resolved = normalized.get(_normalize_for_matching(candidate))
        if resolved is not None:
            return resolved
    return None


def _row_value(row: pd.Series, *candidates: str, default=None):
    for candidate in candidates:
        if candidate in row.index:
            value = row[candidate]
            if not pd.isna(value):
                return value
    return default


def _http_get(url: str) -> str:
    request = Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
            "Accept-Language": "sk-SK,sk;q=0.9,cs;q=0.8,en;q=0.7",
        },
    )
    with urlopen(request, timeout=ALZA_TIMEOUT_SECONDS) as response:
        charset = response.headers.get_content_charset() or "utf-8"
        return response.read().decode(charset, errors="replace")


def _strip_html_tags(html_text: str) -> str:
    text = re.sub(r"(?is)<script.*?>.*?</script>", " ", html_text)
    text = re.sub(r"(?is)<style.*?>.*?</style>", " ", text)
    text = re.sub(r"(?s)<[^>]+>", " ", text)
    text = unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def _extract_json_ld_objects(html_text: str) -> list[dict]:
    objects: list[dict] = []
    for match in re.finditer(r'<script[^>]+type="application/ld\+json"[^>]*>(.*?)</script>', html_text, flags=re.I | re.S):
        raw = match.group(1).strip()
        if not raw:
            continue
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            continue
        if isinstance(parsed, dict):
            objects.append(parsed)
        elif isinstance(parsed, list):
            objects.extend(item for item in parsed if isinstance(item, dict))
    return objects


def _find_product_dict(objects: list[dict]) -> dict | None:
    def _walk(value):
        if isinstance(value, dict):
            if _normalize_for_matching(str(value.get("@type", ""))) == "product":
                return value
            for nested_value in value.values():
                found = _walk(nested_value)
                if found is not None:
                    return found
        elif isinstance(value, list):
            for item in value:
                found = _walk(item)
                if found is not None:
                    return found
        return None

    for obj in objects:
        found = _walk(obj)
        if found is not None:
            return found
    return None


def _best_text_snippets(text: str, keywords: list[str], limit: int = 3) -> list[str]:
    sentences = re.split(r"(?<=[.!?])\s+", text)
    snippets: list[str] = []
    normalized_keywords = [_normalize_for_matching(keyword) for keyword in keywords]
    for sentence in sentences:
        normalized_sentence = _normalize_for_matching(sentence)
        if not normalized_sentence:
            continue
        if any(keyword in normalized_sentence for keyword in normalized_keywords):
            compact = re.sub(r"\s+", " ", sentence).strip()
            if compact and compact not in snippets:
                snippets.append(compact)
        if len(snippets) >= limit:
            break
    return snippets


def _extract_claim_percentages(text: str) -> dict[str, float | None]:
    normalized = _normalize_for_matching(text)
    patterns = {
        "complaints_pct": [
            r"(\d{1,2}(?:[.,]\d+)?)\s*%\s*(?:reklamaci|reklamacie|reklamacii|reklamac)",
            r"reklamaci[^\d]{0,40}(\d{1,2}(?:[.,]\d+)?)\s*%",
        ],
        "withdrawals_pct": [
            r"(\d{1,2}(?:[.,]\d+)?)\s*%\s*(?:odstoupeni|odstupeni|odstupenie|vrateni)",
            r"(?:odstoupeni|odstupeni|odstupenie)[^\d]{0,40}(\d{1,2}(?:[.,]\d+)?)\s*%",
        ],
    }
    result: dict[str, float | None] = {"complaints_pct": None, "withdrawals_pct": None}
    for key, regexes in patterns.items():
        for regex in regexes:
            match = re.search(regex, normalized, flags=re.I)
            if match:
                try:
                    result[key] = float(match.group(1).replace(",", "."))
                except ValueError:
                    result[key] = None
                break
    return result


def _score_packaging_signal(text: str) -> tuple[str, list[str]]:
    normalized = _normalize_for_matching(text)
    packaging_hits = [keyword for keyword in PACKAGING_KEYWORDS if keyword in normalized]
    quality_hits = [keyword for keyword in QUALITY_KEYWORDS if keyword in normalized]
    evidence: list[str] = []
    if packaging_hits:
        evidence.append("Packaging-related keywords: " + ", ".join(sorted(set(packaging_hits))))
    if quality_hits:
        evidence.append("Quality-related keywords: " + ", ".join(sorted(set(quality_hits))))

    if packaging_hits and len(packaging_hits) >= len(quality_hits):
        return "Likely packaging-related issue", evidence
    if quality_hits and len(quality_hits) > len(packaging_hits):
        return "Likely product-quality / other issue", evidence
    if packaging_hits:
        return "Possible packaging issue", evidence
    if quality_hits:
        return "Possible product-quality issue", evidence
    return "Insufficient review evidence", evidence


def _product_report_summary(result) -> pd.DataFrame:
    df = result.cases_enriched.copy()
    code_col = _resolve_column(df, "Kód produktu")
    name_col = _resolve_column(df, "Názov produktu", "Název produktu")
    if df.empty or code_col is None:
        return pd.DataFrame(columns=["Kód produktu", "Názov produktu", "report_count", "case_count"])

    group_cols = [column for column in [code_col, name_col] if column is not None]
    summary = (
        df.groupby(group_cols, dropna=False)
        .agg(
            report_count=("report_count", "sum"),
            case_count=("case_id", "count"),
        )
        .reset_index()
        .sort_values(by=["report_count", "case_count"], ascending=[False, False], kind="mergesort")
    )
    rename_map = {}
    if code_col != "Kód produktu":
        rename_map[code_col] = "Kód produktu"
    if name_col is not None and name_col != "Názov produktu":
        rename_map[name_col] = "Názov produktu"
    if rename_map:
        summary = summary.rename(columns=rename_map)
    if "Názov produktu" not in summary.columns:
        summary["Názov produktu"] = pd.NA
    return summary.reset_index(drop=True)


@st.cache_data(show_spinner=False, ttl=6 * 60 * 60)
def fetch_alza_review_analysis(product_name: str, product_code: str | None = None) -> dict:
    query = product_name.strip() or (product_code or "").strip()
    search_url = ALZA_SEARCH_URL.format(query=quote_plus(query))

    try:
        search_html = _http_get(search_url)
    except Exception as exc:
        return {
            "status": "error",
            "product_query": query,
            "search_url": search_url,
            "error": f"Search request failed: {exc}",
        }

    candidate_urls: list[str] = []
    for href in re.findall(r'href="([^"]+)"', search_html, flags=re.I):
        href = href.strip()
        if not href or href.startswith("#"):
            continue
        if "search.htm" in href.lower():
            continue
        if not re.search(r"/\d+\.htm(?:\?|$)", href):
            continue
        candidate_urls.append(urljoin(ALZA_BASE_URL, href))

    product_url = candidate_urls[0] if candidate_urls else None
    if product_url is None:
        return {
            "status": "not_found",
            "product_query": query,
            "search_url": search_url,
            "error": "No product link found on the Alza search page.",
        }

    try:
        product_html = _http_get(product_url)
    except Exception as exc:
        return {
            "status": "error",
            "product_query": query,
            "search_url": search_url,
            "product_url": product_url,
            "error": f"Product page request failed: {exc}",
        }

    json_ld_objects = _extract_json_ld_objects(product_html)
    product_ld = _find_product_dict(json_ld_objects)

    title_match = re.search(r"<title>(.*?)</title>", product_html, flags=re.I | re.S)
    page_title = unescape(title_match.group(1)).strip() if title_match else None
    if page_title:
        page_title = re.sub(r"\s+\|\s+Alza.*$", "", page_title).strip()

    aggregate = product_ld.get("aggregateRating", {}) if product_ld else {}
    rating_value = aggregate.get("ratingValue")
    review_count = aggregate.get("reviewCount")
    try:
        rating_value = float(str(rating_value).replace(",", ".")) if rating_value is not None else None
    except ValueError:
        rating_value = None
    try:
        review_count = int(float(str(review_count).replace(",", "."))) if review_count is not None else None
    except ValueError:
        review_count = None

    visible_text = _strip_html_tags(product_html)
    snippets = _best_text_snippets(visible_text, PACKAGING_KEYWORDS + QUALITY_KEYWORDS, limit=5)
    percentages = _extract_claim_percentages(visible_text)
    verdict, evidence = _score_packaging_signal(visible_text)

    return {
        "status": "ok",
        "product_query": query,
        "search_url": search_url,
        "product_url": product_url,
        "page_title": page_title,
        "rating_value": rating_value,
        "review_count": review_count,
        "complaints_pct": percentages["complaints_pct"],
        "withdrawals_pct": percentages["withdrawals_pct"],
        "verdict": verdict,
        "evidence": evidence,
        "snippets": snippets,
    }


def get_top_product_for_review(result) -> pd.Series | None:
    summary = _product_report_summary(result)
    if summary.empty:
        return None

    top_row = summary.iloc[0]
    product_code = top_row.get("Kód produktu")
    product_name = top_row.get("Názov produktu")
    cases = result.cases_enriched
    filtered = cases[cases["Kód produktu"].astype(str) == str(product_code)] if "Kód produktu" in cases.columns else cases.head(0)
    selected_case = filtered.sort_values(by=["report_count"], ascending=[False], kind="mergesort").head(1)
    if selected_case.empty:
        return pd.Series(
            {
                "Kód produktu": product_code,
                "Názov produktu": product_name,
                "product_report_count": top_row.get("report_count"),
                "product_case_count": top_row.get("case_count"),
            }
        )

    case_row = selected_case.iloc[0].copy()
    case_row["product_report_count"] = top_row.get("report_count")
    case_row["product_case_count"] = top_row.get("case_count")
    return case_row


def get_uploaded_result() -> tuple[pd.DataFrame, pd.DataFrame, object | None]:
    reports_file = st.session_state.get("reports_file_upload")
    products_file = st.session_state.get("products_file_upload")
    if reports_file is None or products_file is None:
        result, batch_id = load_latest_pipeline_result_cached(latest_batch_key())
        if result is None:
            return pd.DataFrame(), pd.DataFrame(), None
        st.session_state.current_batch_id = batch_id
        st.session_state.current_result = result
        st.session_state.review_data = load_reviews(batch_id)
        return result.reports_clean, result.products_clean, result

    reports_raw = read_excel_upload(reports_file)
    products_raw = read_excel_upload(products_file)
    signature = _input_signature(reports_file, products_file)
    if signature == st.session_state.current_input_signature and st.session_state.current_result is not None:
        result = st.session_state.current_result
        return reports_raw, products_raw, result
    result = run_pipeline(
        reports_raw,
        products_raw,
        top_n=int(st.session_state.get("top_n_input", 100)),
        exclude_stitok=bool(st.session_state.get("exclude_stitok_input", True)),
        instruction_categories=None,
        min_report_count=int(st.session_state.get("min_report_count_input", 0)),
        product_match_status=["Produkt nájdený"],
        has_risk_flag=None,
        trigger_detected=None,
    )
    config = {
        "top_n": int(st.session_state.get("top_n_input", 100)),
        "exclude_stitok": bool(st.session_state.get("exclude_stitok_input", True)),
        "min_report_count": int(st.session_state.get("min_report_count_input", 0)),
    }
    batch_id = save_pipeline_result(result, signature, config)
    if not uses_persistent_cloud_storage():
        backup_database()
    st.session_state.current_batch_id = batch_id
    st.session_state.current_input_signature = signature
    st.session_state.current_result = result
    st.session_state.review_data = load_reviews(batch_id)
    return reports_raw, products_raw, result


def ensure_review_defaults(case_id: str, row: pd.Series) -> dict[str, str]:
    review_data = st.session_state.review_data
    if case_id in review_data:
        return review_data[case_id]

    recommended_action = str(row.get("recommended_action", "") or "")
    default_review_status = "Neriešené"
    if recommended_action == "Odovzdať na IT":
        default_review_status = "Čaká na IT"
    elif recommended_action == "Ponechať balenie":
        default_review_status = "Schválené"
    elif recommended_action in {"Preveriť odstránenie balenia", "Odstrániť balenie"}:
        default_review_status = "Rieši sa"

    default_final_action = recommended_action if recommended_action in FINAL_ACTION_OPTIONS else "Manuálna analýza"
    review_data[case_id] = {
        "review_status": default_review_status,
        "final_action": default_final_action,
        "validator_note": "",
    }
    return review_data[case_id]


def sync_review_widgets(case_id: str, row: pd.Series) -> None:
    if st.session_state.review_widget_case_id == case_id:
        return
    review_values = ensure_review_defaults(case_id, row)
    st.session_state.review_widget_case_id = case_id
    st.session_state.review_status_widget = review_values["review_status"]
    st.session_state.final_action_widget = review_values["final_action"]
    st.session_state.validator_note_widget = review_values["validator_note"]


def save_current_review(case_id: str) -> None:
    st.session_state.review_data[case_id] = {
        "review_status": st.session_state.review_status_widget,
        "final_action": st.session_state.final_action_widget,
        "validator_note": st.session_state.validator_note_widget.strip(),
    }


def build_reviewed_cases(result) -> pd.DataFrame:
    rows = []
    for _, row in result.cases_enriched.iterrows():
        case_id = row.get("case_id")
        review_values = ensure_review_defaults(str(case_id), row)
        rows.append(
            {
                "case_id": case_id,
                "Kód produktu": row.get("Kód produktu"),
                "Název produktu": row.get("Název produktu"),
                "instruction_category": row.get("instruction_category"),
                "report_count": row.get("report_count"),
                "recommended_action": row.get("recommended_action"),
                "suggested_parameter_to_check": row.get("suggested_parameter_to_check"),
                "confidence": row.get("confidence"),
                "review_status": review_values["review_status"],
                "final_action": review_values["final_action"],
                "validator_note": review_values["validator_note"],
                "decision_reason": row.get("decision_reason"),
                "validator_checklist": row.get("validator_checklist"),
            }
        )
    return pd.DataFrame(rows)


def reviewed_cases_bytes(reviewed_cases: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        reviewed_cases.to_excel(writer, sheet_name="reviewed_cases", index=False)
    return buffer.getvalue()


def render_dashboard(result) -> None:
    st.subheader("Key metrics")
    metric_order = [
        ("total_reports", "Total reports"),
        ("total_cases", "Total cases"),
        ("matched_products_count", "Matched products"),
        ("invalid_reports_count", "Invalid reports"),
        ("total_ai_input_rows", "AI input rows"),
    ]
    metric_values = result.summary[result.summary["metric"].isin([item[0] for item in metric_order])].set_index("metric")["value"].to_dict()
    metric_cols = st.columns(len(metric_order))
    for idx, (metric_key, metric_label) in enumerate(metric_order):
        with metric_cols[idx]:
            st.metric(metric_label, int(metric_values.get(metric_key, 0)))

    action_counts = (
        result.cases_enriched.groupby("recommended_action", dropna=False)
        .size()
        .reset_index(name="case_count")
        .sort_values(by=["case_count", "recommended_action"], ascending=[False, True], kind="mergesort")
    )
    st.subheader("Counts by recommended_action")
    if not action_counts.empty:
        action_card_rows = [action_counts.iloc[i:i + 4] for i in range(0, len(action_counts), 4)]
        for chunk in action_card_rows:
            cols = st.columns(len(chunk))
            for col, (_, row) in zip(cols, chunk.iterrows()):
                with col:
                    st.metric(str(row["recommended_action"]), int(row["case_count"]))
    st.dataframe(action_counts, use_container_width=True, hide_index=True)
    if not action_counts.empty:
        chart_data = action_counts.set_index("recommended_action")[["case_count"]]
        st.bar_chart(chart_data)

    matrix = (
        result.cases_enriched.groupby(["instruction_category", "recommended_action"], dropna=False)
        .size()
        .reset_index(name="case_count")
        .sort_values(by=["instruction_category", "case_count", "recommended_action"], ascending=[True, False, True], kind="mergesort")
    )
    st.subheader("Counts by instruction_category and recommended_action")
    st.dataframe(matrix, use_container_width=True, hide_index=True)


def render_customer_online_review(result) -> None:
    st.subheader("Customer online review")
    st.caption("This tab starts with the product that has the most reports and tries to match it with the current Alza product page.")

    summary = _product_report_summary(result)
    if summary.empty:
        st.info("No products available for review yet.")
        return

    code_col = _resolve_column(summary, "Kód produktu")
    name_col = _resolve_column(summary, "Názov produktu", "Název produktu")
    top_options = [
        f"{_row_value(row, name_col, default='(bez názvu)')} [{_row_value(row, code_col, default='')}] - reports: {int(row['report_count'])}, cases: {int(row['case_count'])}"
        for _, row in summary.head(20).iterrows()
    ]
    selected_label = st.selectbox(
        "Product to review",
        options=top_options,
        index=0,
        help="The first item is the product with the highest total report count.",
    )
    selected_index = top_options.index(selected_label)
    selected_summary = summary.iloc[selected_index]

    selected_code = _row_value(selected_summary, code_col)
    selected_name = _row_value(selected_summary, name_col)
    selected_code_text = None if pd.isna(selected_code) else str(selected_code)
    selected_name_text = "" if pd.isna(selected_name) else str(selected_name)
    analysis = fetch_alza_review_analysis(
        selected_name_text,
        selected_code_text,
    )

    info_cols = st.columns(4)
    with info_cols[0]:
        st.metric("Report count", int(selected_summary.get("report_count") or 0))
    with info_cols[1]:
        st.metric("Case count", int(selected_summary.get("case_count") or 0))
    with info_cols[2]:
        st.metric("Alza rating", "-" if analysis.get("rating_value") is None else f"{analysis.get('rating_value'):.1f}/5")
    with info_cols[3]:
        st.metric("Review count", "-" if analysis.get("review_count") is None else int(analysis.get("review_count")))

    st.markdown("**Product**")
    product_details = pd.DataFrame(
        [
            {
                "Kód produktu": selected_code,
                "Názov produktu": selected_name,
                "Search query": analysis.get("product_query"),
                "Search URL": analysis.get("search_url"),
                "Product URL": analysis.get("product_url"),
                "Page title": analysis.get("page_title"),
            }
        ]
    )
    st.dataframe(product_details, use_container_width=True, hide_index=True)

    if analysis.get("status") != "ok":
        st.warning(analysis.get("error", "Online review lookup failed."))
        return

    verdict = analysis.get("verdict", "Insufficient review evidence")
    if verdict == "Likely packaging-related issue":
        st.error(verdict)
    elif verdict == "Possible packaging issue":
        st.warning(verdict)
    elif verdict == "Likely product-quality / other issue":
        st.info(verdict)
    else:
        st.info(verdict)

    review_metrics = st.columns(2)
    with review_metrics[0]:
        complaints_pct = analysis.get("complaints_pct")
        st.metric("% complaints / claims", "-" if complaints_pct is None else f"{complaints_pct:.1f}%")
    with review_metrics[1]:
        withdrawals_pct = analysis.get("withdrawals_pct")
        st.metric("% withdrawals / returns", "-" if withdrawals_pct is None else f"{withdrawals_pct:.1f}%")

    if analysis.get("evidence"):
        st.markdown("**Evidence**")
        for line in analysis["evidence"]:
            st.write(f"- {line}")

    snippets = analysis.get("snippets") or []
    st.markdown("**Review snippets**")
    if snippets:
        for snippet in snippets[:5]:
            st.write(f"- {snippet}")
    else:
        st.write("No review snippets were found in the page text.")

    st.markdown("**Interpretation**")
    if "packaging" in verdict.lower():
        st.success("The current page text suggests packaging could be part of the issue, but this is heuristic and should be confirmed manually.")
    elif "quality" in verdict.lower():
        st.info("The current page text points more toward product quality or a non-packaging issue.")
    else:
        st.info("There is not enough review evidence on the page text to make a reliable call.")


def render_case_detail(case_row: pd.Series) -> None:
    st.subheader("Case detail")

    sections = [
        (
            "Product",
            ["case_id", "Kód produktu", "Název produktu", "Segment1", "Segment2", "Segment3"],
        ),
        (
            "Reported instruction",
            ["Přeložené instrukce", "report_count", "unique_users_count", "first_reported_at", "last_reported_at"],
        ),
        (
            "System reason",
            ["main_trigger_parameter", "trigger_parameter_value", "trigger_detected", "active_risk_parameters", "has_risk_flag"],
        ),
        (
            "Recommended decision",
            ["recommended_action", "suggested_parameter_to_check", "confidence", "decision_reason"],
        ),
    ]

    for title, columns in sections:
        present_columns = [column for column in columns if column in case_row.index]
        st.markdown(f"**{title}**")
        if present_columns:
            data = {column: [case_row.get(column)] for column in present_columns}
            st.dataframe(pd.DataFrame(data), use_container_width=True, hide_index=True)

    st.markdown("**Validator checklist**")
    checklist = str(case_row.get("validator_checklist", "") or "")
    checklist_items = [item.strip() for item in checklist.split(";") if item.strip()]
    if checklist_items:
        for item in checklist_items:
            st.write(f"- {item}")
    else:
        st.write("No checklist available.")


def render_decision_review(result) -> None:
    review_df = result.cases_enriched.copy()

    filter_row_1 = st.columns(4)
    with filter_row_1[0]:
        recommended_action_filter = st.multiselect(
            "Recommended action",
            options=sorted(review_df["recommended_action"].dropna().astype(str).unique().tolist()),
            default=sorted(review_df["recommended_action"].dropna().astype(str).unique().tolist()),
        )
    with filter_row_1[1]:
        instruction_category_filter = st.multiselect(
            "Instruction category",
            options=sorted(review_df["instruction_category"].dropna().astype(str).unique().tolist()),
            default=sorted(review_df["instruction_category"].dropna().astype(str).unique().tolist()),
        )
    with filter_row_1[2]:
        priority_filter = st.multiselect(
            "Priority",
            options=sorted(review_df["priority"].dropna().astype(str).unique().tolist()),
            default=sorted(review_df["priority"].dropna().astype(str).unique().tolist()),
        )
    with filter_row_1[3]:
        confidence_filter = st.multiselect(
            "Confidence",
            options=sorted(review_df["confidence"].dropna().astype(str).unique().tolist()),
            default=sorted(review_df["confidence"].dropna().astype(str).unique().tolist()),
        )

    filter_row_2 = st.columns(4)
    with filter_row_2[0]:
        product_match_status_filter = st.multiselect(
            "Product match status",
            options=sorted(review_df["product_match_status"].dropna().astype(str).unique().tolist()),
            default=sorted(review_df["product_match_status"].dropna().astype(str).unique().tolist()),
        )
    with filter_row_2[1]:
        trigger_detected_filter = st.multiselect(
            "Trigger detected",
            options=sorted(review_df["trigger_detected"].dropna().astype(str).unique().tolist()),
            default=sorted(review_df["trigger_detected"].dropna().astype(str).unique().tolist()),
        )
    with filter_row_2[2]:
        has_risk_flag_filter = st.multiselect(
            "Has risk flag",
            options=sorted(review_df["has_risk_flag"].dropna().astype(str).unique().tolist()),
            default=sorted(review_df["has_risk_flag"].dropna().astype(str).unique().tolist()),
        )
    with filter_row_2[3]:
        minimum_report_count = st.number_input("Minimum report_count", min_value=0, value=0, step=1)

    filtered = review_df.copy()
    for column, values in [
        ("recommended_action", recommended_action_filter),
        ("instruction_category", instruction_category_filter),
        ("priority", priority_filter),
        ("confidence", confidence_filter),
        ("product_match_status", product_match_status_filter),
        ("trigger_detected", trigger_detected_filter),
        ("has_risk_flag", has_risk_flag_filter),
    ]:
        if values:
            filtered = filtered[filtered[column].astype(str).isin(values)]

    if minimum_report_count:
        filtered = filtered[filtered["report_count"] >= minimum_report_count]

    display_columns = [column for column in DISPLAY_CASE_COLUMNS if column in filtered.columns]
    st.subheader("Case list")
    st.dataframe(filtered.loc[:, display_columns], use_container_width=True, hide_index=True)

    case_options = filtered["case_id"].dropna().astype(str).tolist()
    if not case_options:
        st.info("No cases match the current filters.")
        return

    selected_case_id = st.selectbox("Select case_id", options=case_options, index=0)
    selected_row = review_df.loc[review_df["case_id"].astype(str) == selected_case_id].iloc[0]
    sync_review_widgets(selected_case_id, selected_row)

    detail_col_1, detail_col_2 = st.columns([2, 1])
    with detail_col_1:
        render_case_detail(selected_row)
    with detail_col_2:
        st.subheader("Review fields")
        st.selectbox("Review status", options=REVIEW_STATUS_OPTIONS, key="review_status_widget")
        st.selectbox("Final action", options=FINAL_ACTION_OPTIONS, key="final_action_widget")
        st.text_area("Validator note", key="validator_note_widget", height=180)
        if st.button("Save review"):
            save_current_review(selected_case_id)
            st.success("Review saved for this case.")
        else:
            save_current_review(selected_case_id)


def ensure_review_defaults(case_id: str, row: pd.Series) -> dict[str, str]:
    review_data = st.session_state.review_data
    review_data[case_id] = _normalize_review_entry(review_data.get(case_id), row)
    return review_data[case_id]


def save_current_review(case_id: str, case_row: pd.Series, explicit_save: bool = False) -> None:
    review_entry = _normalize_review_entry(st.session_state.review_data.get(case_id), case_row)
    status_key = _review_widget_key(case_id, "status")
    final_action_key = _review_widget_key(case_id, "final_action")
    validator_note_key = _review_widget_key(case_id, "validator_note")

    review_entry["review_status"] = st.session_state.get(status_key, review_entry["review_status"])
    review_entry["status"] = review_entry["review_status"]
    review_entry["final_action"] = st.session_state.get(final_action_key, review_entry["final_action"])
    review_entry["validator_note"] = str(st.session_state.get(validator_note_key, review_entry["validator_note"]) or "").strip()
    if explicit_save:
        review_entry["review_saved_at"] = _current_timestamp()
        review_entry["review_saved_by"] = _current_user_identifier()
    st.session_state.review_data[case_id] = review_entry
    save_review(st.session_state.get("current_batch_id"), case_id, review_entry)
    if explicit_save:
        _mark_case_saved_for_session(case_id)
        if not uses_persistent_cloud_storage():
            backup_database()


def mark_change_as_executed(case_id: str) -> None:
    if case_id not in st.session_state.review_data:
        return
    review_entry = _normalize_review_entry(st.session_state.review_data.get(case_id))
    updated = update_review_execution_fields(review_entry, executed_by=_current_user_identifier())
    st.session_state.review_data[case_id] = updated
    st.session_state[_review_widget_key(case_id, "status")] = updated["review_status"]
    save_review(st.session_state.get("current_batch_id"), case_id, updated)
    _mark_case_saved_for_session(case_id)
    if not uses_persistent_cloud_storage():
        backup_database()


def build_reviewed_cases(result) -> pd.DataFrame:
    rows = []
    for _, row in result.cases_enriched.iterrows():
        case_id = str(row.get("case_id"))
        review_values = ensure_review_defaults(case_id, row)
        original_instruction = _row_safe_value(row, "Přeložené instrukce", "Peložené instrukce", "Přeložená instrukce")
        product_id = _row_safe_value(row, "Product ID", "SEOPrefix_ID", "ID produktu", default="")
        if product_id in {"", "Neuvedené"}:
            product_id = ""
        rows.append(
            {
                "case_id": row.get("case_id"),
                "SKU": _row_safe_value(row, "Kód produktu"),
                "Product ID": product_id,
                "Kód produktu": _row_safe_value(row, "Kód produktu"),
                "Názov produktu": _row_safe_value(row, "Názov produktu", "Název produktu"),
                "original_instruction": original_instruction,
                "final_instruction": review_values["final_action"],
                "status": review_values["review_status"],
                "review_status": review_values["review_status"],
                "final_action": review_values["final_action"],
                "review_saved_at": review_values["review_saved_at"],
                "review_saved_by": review_values["review_saved_by"],
                "Validator note": review_values["validator_note"],
                "validator_note": review_values["validator_note"],
                "change_execution_status": review_values["change_execution_status"],
                "change_executed_at": review_values["change_executed_at"],
                "change_executed_by": review_values["change_executed_by"],
                "instruction_category": row.get("instruction_category"),
                "report_count": row.get("report_count"),
                "recommended_action": row.get("recommended_action"),
                "suggested_parameter_to_check": row.get("suggested_parameter_to_check"),
                "confidence": row.get("confidence"),
                "decision_reason": row.get("decision_reason"),
                "validator_checklist": row.get("validator_checklist"),
            }
        )
    reviewed_cases = pd.DataFrame(rows)
    reviewed_cases = ensure_optional_review_columns(reviewed_cases)
    ordered_columns = [column for column in EXTENDED_REVIEW_EXPORT_COLUMNS if column in reviewed_cases.columns]
    trailing_columns = [column for column in reviewed_cases.columns if column not in ordered_columns]
    return reviewed_cases.loc[:, ordered_columns + trailing_columns]


def reviewed_cases_bytes(reviewed_cases: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        reviewed_cases.to_excel(writer, sheet_name="reviewed_cases", index=False)
    return buffer.getvalue()


def apply_review_statuses(review_df: pd.DataFrame) -> pd.DataFrame:
    """Make the case list reflect persisted review statuses instead of pipeline defaults."""
    updated = review_df.copy()
    if "case_id" not in updated.columns:
        return updated
    status_map = {
        str(case_id): str(review.get("review_status", ""))
        for case_id, review in st.session_state.review_data.items()
        if review.get("review_status")
    }
    persisted_statuses = updated["case_id"].astype(str).map(status_map)
    fallback_statuses = updated["status"] if "status" in updated.columns else pd.Series(
        "Neriešené", index=updated.index
    )
    statuses = persisted_statuses.fillna(fallback_statuses)
    updated["status"] = statuses
    updated["review_status"] = statuses
    updated["last_updated_by"] = updated["case_id"].astype(str).map(
        {
            str(case_id): review.get("change_executed_by") or review.get("review_saved_by", "")
            for case_id, review in st.session_state.review_data.items()
        }
    ).fillna("")
    return updated


def render_case_detail(case_id: str, case_row: pd.Series) -> None:
    st.subheader("Case detail")
    render_product_header(case_row)
    guidance = build_console_action_guidance(
        case_row,
        _normalize_review_entry(st.session_state.review_data.get(case_id), case_row),
    )
    st.divider()
    st.markdown("### Návrh pre internú konzolu")
    st.write(f"**Navrhovaná akcia:** {guidance['action']}")
    st.caption(f"Úroveň istoty návrhu: {guidance['confidence']}")
    if guidance["description_available"]:
        st.markdown("**Popis produktu**")
        st.write(guidance["product_description"])
    else:
        st.info(
            "Popis produktu nie je v aktuálnej produktovej dávke. "
            "Návrh používa názov, kategóriu a dostupné produktové parametre."
        )
    st.dataframe(guidance["steps"], use_container_width=True, hide_index=True)
    st.divider()
    render_case_reason(case_row)
    st.divider()
    render_instruction_comparison(case_row)
    st.divider()
    render_review_form(case_id, case_row)


def render_decision_review(result) -> None:
    review_df = apply_review_statuses(result.cases_enriched)

    def _options_for(column: str) -> list[str]:
        if column not in review_df.columns:
            return []
        return sorted(review_df[column].dropna().astype(str).unique().tolist())

    filter_row_1 = st.columns(4)
    with filter_row_1[0]:
        recommended_action_filter = st.multiselect(
            "Recommended action",
            options=_options_for("recommended_action"),
            default=_options_for("recommended_action"),
        )
    with filter_row_1[1]:
        instruction_category_filter = st.multiselect(
            "Instruction category",
            options=_options_for("instruction_category"),
            default=_options_for("instruction_category"),
        )
    with filter_row_1[2]:
        priority_filter = st.multiselect(
            "Priority",
            options=_options_for("priority"),
            default=_options_for("priority"),
        )
    with filter_row_1[3]:
        confidence_filter = st.multiselect(
            "Confidence",
            options=_options_for("confidence"),
            default=_options_for("confidence"),
        )

    filter_row_2 = st.columns(4)
    with filter_row_2[0]:
        product_match_status_filter = st.multiselect(
            "Product match status",
            options=_options_for("product_match_status"),
            default=_options_for("product_match_status"),
        )
    with filter_row_2[1]:
        trigger_detected_filter = st.multiselect(
            "Trigger detected",
            options=_options_for("trigger_detected"),
            default=_options_for("trigger_detected"),
        )
    with filter_row_2[2]:
        has_risk_flag_filter = st.multiselect(
            "Has risk flag",
            options=_options_for("has_risk_flag"),
            default=_options_for("has_risk_flag"),
        )
    with filter_row_2[3]:
        minimum_report_count = st.number_input("Minimum report_count", min_value=0, value=0, step=1)

    filtered = review_df.copy()
    for column, values in [
        ("recommended_action", recommended_action_filter),
        ("instruction_category", instruction_category_filter),
        ("priority", priority_filter),
        ("confidence", confidence_filter),
        ("product_match_status", product_match_status_filter),
        ("trigger_detected", trigger_detected_filter),
        ("has_risk_flag", has_risk_flag_filter),
    ]:
        if values and column in filtered.columns:
            filtered = filtered[filtered[column].astype(str).isin(values)]

    if minimum_report_count and "report_count" in filtered.columns:
        filtered = filtered[filtered["report_count"] >= minimum_report_count]

    saved_case_ids = {
        str(case_id)
        for case_id, review in st.session_state.review_data.items()
        if _review_has_saved_work(review)
    }
    case_view = st.radio(
        "Case list",
        options=["Cases", "Cases solved"],
        horizontal=True,
        label_visibility="collapsed",
    )
    if case_view == "Cases":
        filtered = filtered[~filtered["case_id"].astype(str).isin(saved_case_ids)]
        list_title = "Cases"
    else:
        filtered = filtered[filtered["case_id"].astype(str).isin(saved_case_ids)]
        list_title = "Cases solved"

    st.subheader(list_title)

    count_columns = st.columns([1, 1, 1, 3])
    with count_columns[0]:
        page_size = st.selectbox(
            "Načítať naraz",
            options=[10, 20, 50, 100],
            key="review_list_page_size",
        )

    filtered_case_ids = (
        filtered["case_id"].dropna().astype(str).tolist() if "case_id" in filtered.columns else []
    )
    filter_signature = hashlib.sha256(
        f"{case_view}|{page_size}|".encode("utf-8")
        + "\0".join(filtered_case_ids).encode("utf-8")
    ).hexdigest()
    if st.session_state.review_case_filter_signature != filter_signature:
        st.session_state.review_case_filter_signature = filter_signature
        st.session_state.review_visible_case_count = page_size

    visible_count = min(int(st.session_state.review_visible_case_count), len(filtered))
    remaining_count = max(len(filtered) - visible_count, 0)
    with count_columns[1]:
        if st.button(
            f"Načítať ďalších {page_size}",
            disabled=remaining_count == 0,
            use_container_width=True,
        ):
            visible_count = min(visible_count + page_size, len(filtered))
            st.session_state.review_visible_case_count = visible_count
            remaining_count = max(len(filtered) - visible_count, 0)
    with count_columns[2]:
        if st.button(
            "Zobraziť všetky",
            disabled=remaining_count == 0,
            use_container_width=True,
        ):
            visible_count = len(filtered)
            st.session_state.review_visible_case_count = visible_count
            remaining_count = 0

    display_columns = [column for column in DISPLAY_CASE_COLUMNS if column in filtered.columns]
    visible_cases = filtered.head(visible_count)
    st.caption(
        f"Zobrazených {len(visible_cases)} z {len(filtered)} filtrovaných cases. "
        f"Zostáva {remaining_count}."
    )
    st.dataframe(visible_cases.loc[:, display_columns], use_container_width=True, hide_index=True)

    case_options = visible_cases["case_id"].dropna().astype(str).tolist() if "case_id" in visible_cases.columns else []
    if not case_options:
        st.info("No cases match the current filters.")
        return

    if st.session_state.get("decision_case_selector") not in case_options:
        st.session_state["decision_case_selector"] = case_options[0]
    selected_case_id = st.selectbox(
        "Select case_id",
        options=case_options,
        key="decision_case_selector",
    )
    selected_row = review_df.loc[review_df["case_id"].astype(str) == selected_case_id].iloc[0]
    ensure_review_defaults(selected_case_id, selected_row)
    render_case_detail(selected_case_id, selected_row)


def render_upload_processing(result) -> None:
    st.subheader("Zdrojové dáta")
    if result is not None:
        st.success("Používa sa posledná uložená dátová dávka. Excel súbory netreba nahrávať znova.")
    st.caption(f"Úložisko: {storage_location_label()}")

    upload_title = "Nahrať novú dátovú dávku" if result is not None else "Nahrať prvú dátovú dávku"
    with st.expander(upload_title, expanded=result is None):
        upload_cols = st.columns(3)
        with upload_cols[0]:
            st.file_uploader("Chybné instrukcie Excel", type=["xlsx"], key="reports_file_upload")
        with upload_cols[1]:
            st.file_uploader("Produkty a vlastnosti Excel", type=["xlsx"], key="products_file_upload")
        with upload_cols[2]:
            st.file_uploader(
                "Baliace pravidlá Excel (optional)",
                type=["xlsx"],
                key="packing_rules_file_upload",
                help="Optional in this MVP. The pipeline does not depend on it.",
            )

        settings_cols = st.columns(3)
        with settings_cols[0]:
            st.number_input("TOP N", min_value=1, value=100, step=1, key="top_n_input")
        with settings_cols[1]:
            st.checkbox("Exclude Štítok", value=True, key="exclude_stitok_input")
        with settings_cols[2]:
            st.number_input("Minimum report_count", min_value=0, value=0, step=1, key="min_report_count_input")

    if result is not None:
        with st.expander("Doplniť iba popisy produktov", expanded=False):
            st.caption(
                "Excel musí obsahovať Kód produktu a Ext. jméno alebo Popis produktu. "
                "Hlásenia, cases a uložené rozhodnutia sa nemenia."
            )
            descriptions_file = st.file_uploader(
                "Excel s popismi produktov",
                type=["xlsx"],
                key="product_descriptions_upload",
            )
            if st.button("Pridať popisy do aktuálnej dávky", type="primary"):
                if descriptions_file is None:
                    st.warning("Najprv vyberte Excel s popismi produktov.")
                else:
                    try:
                        descriptions_raw = read_excel_upload(descriptions_file)
                        import_stats = enrich_current_batch_with_descriptions(result, descriptions_raw)
                        batch_id = st.session_state.get("current_batch_id")
                        if batch_id is None:
                            raise ValueError("Aktuálna dátová dávka nemá batch_id.")
                        update_batch_tables(
                            batch_id,
                            result,
                            ["products_clean", "cases_enriched", "ai_cases_input"],
                        )
                        st.session_state.current_result = result
                        load_latest_pipeline_result_cached.clear()
                        if not uses_persistent_cloud_storage():
                            backup_database()
                        st.success(
                            f"Popisy boli doplnené pre {import_stats['matched_products']} produktov "
                            f"a {import_stats['matched_cases']} cases. "
                            f"Bez zhody: {import_stats['unmatched_products']}."
                        )
                    except Exception as exc:
                        st.error(f"Popisy sa nepodarilo doplniť: {exc}")

    reports_file = st.session_state.get("reports_file_upload")
    products_file = st.session_state.get("products_file_upload")
    if reports_file is None or products_file is None:
        if result is None:
            st.info("Upload Chybné instrukcie and Produkty a vlastnosti to run the pipeline.")
    elif result is not None:
        if result.warnings:
            st.subheader("Warnings")
            for warning in result.warnings:
                st.warning(warning)
        st.success("Processing complete. Open the Dashboard and Decision Review tabs.")


def render_data_exports(result, role: str) -> None:
    st.subheader("Dáta a exporty")
    reviewed_cases = build_reviewed_cases(result)
    current_user = _current_user_identifier()

    if role == "admin":
        saved_mask = (
            reviewed_cases["review_saved_at"].astype(str).ne("")
            | reviewed_cases["change_executed_at"].astype(str).ne("")
            | reviewed_cases["validator_note"].astype(str).ne("")
        )
        executed_mask = reviewed_cases["change_execution_status"].astype(str).eq("Zmena vykonaná")
        metric_cols = st.columns(3)
        metric_cols[0].metric("Všetky prípady", len(reviewed_cases))
        metric_cols[1].metric("Uložené rozhodnutia", int(saved_mask.sum()))
        metric_cols[2].metric("Vykonané zmeny", int(executed_mask.sum()))
        st.download_button(
            "Stiahnuť všetky rozhodnutia (Excel)",
            data=reviewed_cases_bytes(reviewed_cases),
            file_name="all_decisions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        reviewer_token = os.environ.get("INSTRUCTION_VALIDATOR_REVIEWER_TOKEN", "")
        if reviewer_token:
            reviewer_url = f"{st.context.url}?reviewer_token={quote_plus(reviewer_token)}"
            st.markdown("**Prístupový link pre kvalitára**")
            st.code(reviewer_url, language=None, wrap_lines=True)
            st.link_button("Otvoriť testovací link", reviewer_url)
            st.caption("Každý, kto pozná tento link, sa môže prihlásiť ako kvalitár.")
        else:
            st.warning("Prístupový link kvalitára nie je nakonfigurovaný v Streamlit Secrets.")
        if st.button("Pripraviť databázovú zálohu"):
            backup_path = backup_database()
            if backup_path is not None:
                st.session_state["admin_database_backup"] = {
                    "name": backup_path.name,
                    "data": backup_path.read_bytes(),
                }
        prepared_backup = st.session_state.get("admin_database_backup")
        if prepared_backup:
            st.download_button(
                "Stiahnuť databázovú zálohu",
                data=prepared_backup["data"],
                file_name=prepared_backup["name"],
                mime="application/json" if prepared_backup["name"].endswith(".json") else "application/vnd.sqlite3",
            )
        st.caption(f"Aktívne úložisko: {storage_location_label()}")
        st.dataframe(reviewed_cases, use_container_width=True, hide_index=True)
        return

    completed_mask = (
        reviewed_cases["change_execution_status"].astype(str).eq("Zmena vykonaná")
        | reviewed_cases["review_status"].astype(str).eq("Zmena vykonaná")
        | reviewed_cases["status"].astype(str).eq("Zmena vykonaná")
    )
    user_mask = (
        reviewed_cases["review_saved_by"].astype(str).eq(current_user)
        | reviewed_cases["change_executed_by"].astype(str).eq(current_user)
    )
    reviewer_cases = reviewed_cases.loc[completed_mask & user_mask].copy()
    st.caption(f"Rozhodnutia používateľa {current_user}: {len(reviewer_cases)}")
    if reviewer_cases.empty:
        st.info("Zatiaľ nemáte žiadne dokončené rozhodnutia.")
        return
    st.download_button(
        "Stiahnuť moje dokončené rozhodnutia",
        data=reviewed_cases_bytes(reviewer_cases),
        file_name="my_completed_decisions.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.dataframe(reviewer_cases, use_container_width=True, hide_index=True)


def main() -> None:
    init_session_state()
    role = render_access_login()
    st.title("Packaging Instruction Enrichment & Decision Maker")
    st.caption("Turn packaging-instruction reviews into a clearer decision workflow for non-technical users.")

    if role is None:
        st.info("Aplikácia obsahuje firemné dáta. Pre pokračovanie sa prihláste.")
        st.stop()
        return

    reports_raw, products_raw, result = get_uploaded_result()

    section_labels = ["Dashboard", "Decision Review", "Customer online review", "Dáta a exporty"]
    if role == "admin":
        section_labels.insert(0, "Admin upload")
    selected_section = st.radio("Section", section_labels, horizontal=True, label_visibility="collapsed")

    if selected_section == "Admin upload":
        render_upload_processing(result)
    elif selected_section == "Dashboard":
        if result is None:
            st.info("Upload both required files to see the dashboard.")
        else:
            render_dashboard(result)
    elif selected_section == "Decision Review":
        if result is None:
            st.info("Upload both required files to review cases.")
        else:
            render_decision_review(result)
    elif selected_section == "Customer online review":
        if result is None:
            st.info("Upload both required files to look up online reviews.")
        else:
            render_customer_online_review(result)
    elif selected_section == "Dáta a exporty":
        if result is None:
            st.info("Upload both required files to enable downloads.")
        else:
            render_data_exports(result, role)


if __name__ == "__main__":
    main()
