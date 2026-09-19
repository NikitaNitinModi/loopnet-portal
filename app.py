"""
Project Search & Brochure Portal (v2 — real dataset)
------------------------------------------------------
Loads a property list + per-property attachment/brochure links compiled by
the user (from their own LoopNet research, e.g. exported to Excel), and
gives them search-by-area, view-all-attachments, download, and email.

Data files (in data/):
  projects.csv     one row per property
  attachments.csv  one row per (property, attachment) — a property can
                    have zero, one, or many attachments/brochures

Run:
    pip install -r requirements.txt
    streamlit run app.py
"""

import re
import smtplib
import ssl
import zipfile
from email.message import EmailMessage
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

BASE_DIR = Path(__file__).parent
PROJECTS_CSV = BASE_DIR / "data" / "projects.csv"
ATTACHMENTS_CSV = BASE_DIR / "data" / "attachments.csv"

st.set_page_config(page_title="Project Search & Brochure Portal", page_icon="🏢", layout="wide")


# --------------------------------------------------------------------------
# Data loading
# --------------------------------------------------------------------------

@st.cache_data
def load_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    projects = pd.read_csv(PROJECTS_CSV, dtype=str).fillna("")
    if ATTACHMENTS_CSV.exists():
        attachments = pd.read_csv(ATTACHMENTS_CSV, dtype=str).fillna("")
    else:
        attachments = pd.DataFrame(columns=["project_id", "label", "url"])
    return projects, attachments


def filter_projects(df: pd.DataFrame, query: str, area: str, status: str) -> pd.DataFrame:
    out = df
    if area and area != "All areas":
        out = out[out["area"] == area]
    if status and status != "All":
        out = out[out["status"] == status]
    if query.strip():
        q = query.strip().lower()
        mask = (
            out["city"].str.lower().str.contains(q, na=False)
            | out["zipcode"].astype(str).str.contains(q, na=False)
            | out["state"].str.lower().str.contains(q, na=False)
            | out["address"].str.lower().str.contains(q, na=False)
            | out["project_name"].str.lower().str.contains(q, na=False)
            | out["property_type"].str.lower().str.contains(q, na=False)
            | out["area"].str.lower().str.contains(q, na=False)
        )
        out = out[mask]
    return out


@st.cache_data(show_spinner=False)
def fetch_url_bytes(url: str) -> bytes | None:
    """Download a specific, known attachment URL (the direct PDF link
    already compiled for this property) — not a search or crawl."""
    try:
        resp = requests.get(url, timeout=20)
        resp.raise_for_status()
        return resp.content
    except Exception:  # noqa: BLE001
        return None


def is_pdf_url(url: str) -> bool:
    return url.lower().split("?")[0].endswith(".pdf")


def filename_from_url(url: str) -> str:
    name = url.split("/")[-1].split("?")[0]
    name = requests.utils.unquote(name)
    return name or "attachment.pdf"


def sanitize_path_part(name: str) -> str:
    """Make a string safe to use as a folder/file name in a ZIP."""
    name = (name or "untitled").strip()
    name = re.sub(r'[\\/:*?"<>|]', "-", name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    return name[:80] or "untitled"


def build_flyers_zip(projects_df: pd.DataFrame, attachments_df: pd.DataFrame) -> tuple[bytes, int, int]:
    """Build a ZIP of every PDF attachment for the given properties,
    organized as Area/Status/Property Name/Attachment Label.pdf.
    Returns (zip_bytes, property_count_with_attachments, file_count)."""
    buf = BytesIO()
    file_count = 0
    properties_with_files = set()

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for _, row in projects_df.iterrows():
            pid = row["project_id"]
            pdf_rows = attachments_df[
                (attachments_df["project_id"] == pid) & (attachments_df["url"].apply(is_pdf_url))
            ]
            if pdf_rows.empty:
                continue

            folder = "/".join([
                sanitize_path_part(row.get("area", "")),
                sanitize_path_part(row.get("status", "")),
                sanitize_path_part(row.get("project_name", pid)),
            ])

            for _, att in pdf_rows.iterrows():
                data = fetch_url_bytes(att["url"])
                if not data:
                    continue
                label = sanitize_path_part(att["label"])
                ext = Path(filename_from_url(att["url"])).suffix or ".pdf"
                zf.writestr(f"{folder}/{label}{ext}", data)
                file_count += 1
                properties_with_files.add(pid)

        # A short manifest at the root, listing what's inside and where it came from.
        lines = ["Flyers export", "=" * 40, ""]
        for _, row in projects_df.iterrows():
            pid = row["project_id"]
            pdf_rows = attachments_df[
                (attachments_df["project_id"] == pid) & (attachments_df["url"].apply(is_pdf_url))
            ]
            if pdf_rows.empty:
                continue
            lines.append(f"{row['project_name']} ({row['area']}, {row['status']})")
            lines.append(f"  {row['address']}, {row['city']}, {row['state']} {row['zipcode']}")
            if row.get("loopnet_url"):
                lines.append(f"  LoopNet: {row['loopnet_url']}")
            for _, att in pdf_rows.iterrows():
                lines.append(f"  - {att['label']}: {att['url']}")
            lines.append("")
        zf.writestr("README.txt", "\n".join(lines))

    return buf.getvalue(), len(properties_with_files), file_count


# --------------------------------------------------------------------------
# Email
# --------------------------------------------------------------------------

def send_brochures_email(
    smtp_host, smtp_port, smtp_user, smtp_password, use_tls,
    from_addr, to_addr, subject, body, attachments: list[tuple[bytes, str]],
) -> tuple[bool, str]:
    try:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = from_addr
        msg["To"] = to_addr
        msg.set_content(body)
        for data, filename in attachments:
            msg.add_attachment(data, maintype="application", subtype="pdf", filename=filename)

        context = ssl.create_default_context()
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            if use_tls:
                server.starttls(context=context)
            if smtp_user:
                server.login(smtp_user, smtp_password)
            server.send_message(msg)
        return True, f"Email sent to {to_addr} with {len(attachments)} attachment(s)."
    except Exception as exc:  # noqa: BLE001
        return False, f"Failed to send email: {exc}"


def send_zip_email(
    smtp_host, smtp_port, smtp_user, smtp_password, use_tls,
    from_addr, to_addr, subject, body, zip_bytes: bytes, zip_filename: str,
) -> tuple[bool, str]:
    """Send a single ZIP file (the structured flyer folder) as one attachment."""
    try:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = from_addr
        msg["To"] = to_addr
        msg.set_content(body)
        msg.add_attachment(zip_bytes, maintype="application", subtype="zip", filename=zip_filename)

        context = ssl.create_default_context()
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            if use_tls:
                server.starttls(context=context)
            if smtp_user:
                server.login(smtp_user, smtp_password)
            server.send_message(msg)
        return True, f"Email sent to {to_addr} with {zip_filename} attached."
    except Exception as exc:  # noqa: BLE001
        return False, f"Failed to send email: {exc}"


# --------------------------------------------------------------------------
# Popup: email selected properties as one ZIP
# --------------------------------------------------------------------------

@st.dialog("Email selected properties as a ZIP")
def email_selected_dialog(selected_rows: pd.DataFrame, attachments: pd.DataFrame) -> None:
    st.write(
        f"**{len(selected_rows)}** propert{'y' if len(selected_rows)==1 else 'ies'} selected: "
        + ", ".join(selected_rows["project_name"].tolist())
    )

    to_addr = st.text_input("Recipient email", key="dlg_to")
    subject = st.text_input(
        "Subject",
        value=f"Property Brochures — {len(selected_rows)} propert{'y' if len(selected_rows)==1 else 'ies'}",
        key="dlg_subject",
    )
    body = st.text_area(
        "Message",
        value="Hi,\n\nAttached is a ZIP with the brochures for the selected properties, organized by area and property name.\n\nBest regards,",
        height=110,
        key="dlg_body",
    )

    with st.expander("SMTP settings"):
        smtp_host = st.text_input("SMTP host", key="dlg_host")
        smtp_port = st.number_input("SMTP port", value=587, step=1, key="dlg_port")
        smtp_user = st.text_input("SMTP username", key="dlg_user")
        smtp_password = st.text_input("SMTP password", type="password", key="dlg_pass")
        use_tls = st.checkbox("Use TLS", value=True, key="dlg_tls")
        from_addr = st.text_input("From address", key="dlg_from")

    col1, col2 = st.columns(2)
    send_clicked = col1.button("Send ZIP", type="primary", key="dlg_send")
    cancel_clicked = col2.button("Cancel", key="dlg_cancel")

    if cancel_clicked:
        st.rerun()

    if send_clicked:
        if not to_addr:
            st.error("Please enter a recipient email address.")
        elif not smtp_host or not from_addr:
            st.error("Please fill in SMTP host and From address.")
        else:
            with st.spinner("Building ZIP and sending..."):
                zip_bytes, n_props, n_files = build_flyers_zip(selected_rows, attachments)
                if n_files == 0:
                    st.error("None of the selected properties have a downloadable attachment.")
                else:
                    zip_name = f"selected_flyers_{n_props}_properties.zip"
                    ok, message = send_zip_email(
                        smtp_host, int(smtp_port), smtp_user, smtp_password, use_tls,
                        from_addr, to_addr, subject, body, zip_bytes, zip_name,
                    )
                    if ok:
                        st.success(message)
                    else:
                        st.error(message)


# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------

def main() -> None:
    st.title("🏢 Project Search & Brochure Portal")
    st.caption(
        "Search properties by area, city, or zip code. View every attachment and "
        "useful link, download brochures, or email them directly."
    )

    projects, attachments = load_data()

    if projects.empty:
        st.warning("No project data found.")
        return

    areas = ["All areas"] + sorted([a for a in projects["area"].unique() if a])
    statuses = ["All"] + sorted([s for s in projects["status"].unique() if s])

    col_a, col_b, col_c = st.columns([2, 1, 1])
    with col_a:
        query = st.text_input("Search by city, zip code, area, or property type", placeholder="e.g. Tomball, 77484, Retail...")
    with col_b:
        area = st.selectbox("Area", areas)
    with col_c:
        status = st.selectbox("Status", statuses)

    results = filter_projects(projects, query, area, status)
    st.write(f"**{len(results)}** propert{'y' if len(results)==1 else 'ies'} found")

    if results.empty:
        st.info("No matching properties. Try a different search term, area, or status.")
        return

    # --- Bulk flyers for the current filter -----------------------------
    with st.container(border=True):
        filter_label_parts = [p for p in [area if area != "All areas" else "", status if status != "All" else "", query.strip()] if p]
        filter_label = " · ".join(filter_label_parts) if filter_label_parts else "all properties"
        st.markdown(f"**📦 Bulk flyers for this filter** — *{filter_label}* ({len(results)} propert{'y' if len(results)==1 else 'ies'})")

        bulk_cols = st.columns([1, 2])
        with bulk_cols[0]:
            build_clicked = st.button("Build ZIP for this filter", key="build_zip_btn")

        if build_clicked or st.session_state.get("bulk_zip_bytes"):
            if build_clicked:
                with st.spinner("Fetching flyers and building the ZIP..."):
                    zip_bytes, n_props, n_files = build_flyers_zip(results, attachments)
                st.session_state["bulk_zip_bytes"] = zip_bytes
                st.session_state["bulk_zip_stats"] = (n_props, n_files)
                st.session_state["bulk_zip_filter_label"] = filter_label

            zip_bytes = st.session_state["bulk_zip_bytes"]
            n_props, n_files = st.session_state["bulk_zip_stats"]

            if n_files == 0:
                st.warning("None of the properties in this filter have a downloadable flyer/attachment.")
            else:
                st.success(f"ZIP ready: {n_files} file(s) across {n_props} propert{'y' if n_props==1 else 'ies'}, organized as Area/Status/Property Name/.")
                zip_name = f"flyers_{sanitize_path_part(filter_label)}.zip"
                st.download_button(
                    "⬇️ Download ZIP",
                    data=zip_bytes,
                    file_name=zip_name,
                    mime="application/zip",
                    key="download_bulk_zip",
                )

                st.write("**Email this ZIP:**")
                with st.form("bulk_zip_email_form"):
                    zto_addr = st.text_input("Recipient email", key="zto")
                    zsubject = st.text_input("Subject", value=f"Property Flyers — {filter_label}", key="zsubj")
                    zbody = st.text_area(
                        "Message",
                        value=f"Hi,\n\nAttached is a ZIP of flyers for {filter_label} ({n_props} properties, {n_files} files), organized by area and property.\n\nBest regards,",
                        height=110,
                        key="zbody",
                    )
                    with st.expander("SMTP settings"):
                        zsmtp_host = st.text_input("SMTP host", key="zhost")
                        zsmtp_port = st.number_input("SMTP port", value=587, step=1, key="zport")
                        zsmtp_user = st.text_input("SMTP username", key="zuser")
                        zsmtp_password = st.text_input("SMTP password", type="password", key="zpass")
                        zuse_tls = st.checkbox("Use TLS", value=True, key="ztls")
                        zfrom_addr = st.text_input("From address", key="zfrom")

                    zsubmitted = st.form_submit_button("Send ZIP by email")
                    if zsubmitted:
                        if not zto_addr:
                            st.error("Please enter a recipient email address.")
                        elif not zsmtp_host or not zfrom_addr:
                            st.error("Please fill in SMTP host and From address.")
                        else:
                            ok, message = send_zip_email(
                                zsmtp_host, int(zsmtp_port), zsmtp_user, zsmtp_password, zuse_tls,
                                zfrom_addr, zto_addr, zsubject, zbody, zip_bytes, zip_name,
                            )
                            (st.success if ok else st.error)(message)

    st.divider()

    selected_ids = st.session_state.setdefault("selected_ids", set())

    for _, row in results.iterrows():
        pid = row["project_id"]
        row_attachments = attachments[attachments["project_id"] == pid]

        with st.container(border=True):
            col1, col2 = st.columns([4, 1])

            with col1:
                badge = "🟢 For Lease" if row["status"] == "For Lease" else "🔵 For Sale"
                st.subheader(f"{row['project_name']}")
                st.caption(f"{badge} · {row['area']}")
                st.write(f"📍 {row['address']}, {row['city']}, {row['state']} {row['zipcode']}".strip())

                meta_cols = st.columns(4)
                meta_cols[0].metric("Type", row["property_type"] or "—")
                meta_cols[1].metric("Size", row["size"] or "—")
                meta_cols[2].metric("Price/Rent", row["price"] or "—")
                meta_cols[3].metric("Year Built", row["year_built"] or "—")

                if row["description"]:
                    st.caption(row["description"])

                if row["loopnet_url"]:
                    st.write(f"🔗 [View on LoopNet]({row['loopnet_url']})")

                if not row_attachments.empty:
                    st.write("**Attachments:**")
                    for _, att in row_attachments.iterrows():
                        if is_pdf_url(att["url"]):
                            st.write(f"📄 {att['label']} — [{filename_from_url(att['url'])}]({att['url']})")
                        else:
                            st.write(f"🔗 [{att['label']}]({att['url']})")
                else:
                    st.caption("No attachments captured for this property yet.")

            with col2:
                has_pdf = not row_attachments[row_attachments["url"].apply(is_pdf_url)].empty
                checked = st.checkbox(
                    "Select for email",
                    key=f"select_{pid}",
                    value=pid in selected_ids,
                    disabled=not has_pdf,
                )
                if checked:
                    selected_ids.add(pid)
                else:
                    selected_ids.discard(pid)

                pdf_attachments = row_attachments[row_attachments["url"].apply(is_pdf_url)]
                for _, att in pdf_attachments.iterrows():
                    data = fetch_url_bytes(att["url"])
                    if data:
                        st.download_button(
                            f"⬇️ {att['label']}",
                            data=data,
                            file_name=filename_from_url(att["url"]),
                            mime="application/pdf",
                            key=f"dl_{pid}_{att['label']}",
                        )

    st.session_state["selected_ids"] = selected_ids

    # --- Email selected as ZIP (popup) ---------------------------------------
    st.divider()
    st.subheader("📧 Email selected properties")

    selected_rows = results[results["project_id"].isin(selected_ids)]

    if selected_rows.empty:
        st.caption("Select one or more properties above (checkbox) to email their brochures as a single ZIP.")
    else:
        st.write(
            f"**{len(selected_rows)}** propert{'y' if len(selected_rows)==1 else 'ies'} selected: "
            + ", ".join(selected_rows["project_name"].tolist())
        )
        if st.button("📧 Email selected as ZIP", type="primary", key="open_email_dialog"):
            email_selected_dialog(selected_rows, attachments)


if __name__ == "__main__":
    main()
