# -*- coding: utf-8 -*-
import json
import streamlit as st
import pandas as pd
import io
import os
import base64
import requests
from datetime import datetime, timedelta
from typing import Optional, Dict, Any, List, Tuple
from msal import PublicClientApplication, SerializableTokenCache
from dotenv import load_dotenv

st.set_page_config(page_title="Fatura Mailer", page_icon="📧", layout="wide")

GRAPH_AUTHORITY_TEMPLATE = "https://login.microsoftonline.com/{tenant_id}"
GRAPH_SCOPES = ["User.Read", "Mail.Send"]
GRAPH_SENDMAIL_URL = "https://graph.microsoft.com/v1.0/me/sendMail"

# ---------- Utility ----------
def load_env():
    load_dotenv()
    client_id = os.getenv("GRAPH_CLIENT_ID", "").strip()
    tenant_id = os.getenv("GRAPH_TENANT_ID", "").strip()
    if not client_id or not tenant_id:
        st.error("`.env` dosyasında GRAPH_CLIENT_ID ve/veya GRAPH_TENANT_ID yok.")
        st.stop()
    return client_id, tenant_id

def get_token(client_id: str, tenant_id: str, cache_path: str = "token_cache.json") -> Dict[str, Any]:
    cache = SerializableTokenCache()
    if os.path.exists(cache_path):
        cache.deserialize(open(cache_path, "r", encoding="utf-8").read())
    app = PublicClientApplication(
        client_id=client_id,
        authority=GRAPH_AUTHORITY_TEMPLATE.format(tenant_id=tenant_id),
        token_cache=cache
    )
    # Silent attempt
    result = app.acquire_token_silent(GRAPH_SCOPES, account=None)
    if not result:
        flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
        if "user_code" not in flow:
            st.error("Cihaz kodu akışı başlatılamadı. App registration ayarlarını kontrol edin.")
            st.stop()
        # Show the code and verification link to the user
        with st.expander("🔐 Oturum Aç (Microsoft)", expanded=True):
            st.markdown(f"[Doğrulama sayfasını açmak için tıklayın]({flow['verification_uri']})")
            st.code(flow["user_code"], language=None)
            st.info("Yukarıdaki bağlantıyı açın, kodu girip oturum açın ve izinleri onaylayın. Bu paneli kapatmayın; giriş tamamlanınca otomatik devam eder.")
        result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        st.error(f"Token alınamadı: {result.get('error_description', result)}")
        st.stop()
    open(cache_path, "w", encoding="utf-8").write(cache.serialize())
    return result

def parse_recipients(value: Optional[str]) -> List[Dict[str, Dict[str, str]]]:
    if not value:
        return []
    raw = str(value).replace(";", ",")
    addrs = [a.strip() for a in raw.split(",") if a.strip()]
    return [{"emailAddress": {"address": a}} for a in addrs]

def file_to_base64(path: str) -> Tuple[str, str]:
    import mimetypes
    ctype, _ = mimetypes.guess_type(path)
    ctype = ctype or "application/octet-stream"
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("utf-8")
    return b64, ctype

def build_graph_message(to_addr: str, subject: str, body_html: str,
                        attachment_path: Optional[str], cc_list: Optional[List[Dict[str, Dict[str, str]]]] = None) -> Dict[str, Any]:
    message: Dict[str, Any] = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": [{"emailAddress": {"address": to_addr}}],
            "ccRecipients": cc_list or [],
        },
        "saveToSentItems": True
    }
    if attachment_path and os.path.isfile(attachment_path):
        data_b64, ctype = file_to_base64(attachment_path)
        name = os.path.basename(attachment_path)
        attachment = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": name,
            "contentType": ctype,
            "contentBytes": data_b64,
        }
        message["message"]["attachments"] = [attachment]
    else:
        st.warning(f"Ek bulunamadı: {attachment_path}")
    return message

def send_mail_graph(access_token: str, payload: Dict[str, Any]) -> None:
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    r = requests.post(GRAPH_SENDMAIL_URL, headers=headers, json=payload, timeout=45)
    if r.status_code not in (202, 200):
        raise RuntimeError(f"Graph sendMail başarısız: {r.status_code} - {r.text}")

def load_templates(path: str) -> Dict[str, Dict[str, str]]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def choose_template(row: pd.Series, default_template: Optional[str]) -> str:
    tc = str(row.get("template_choice", "") or "").strip().lower()
    if tc in ("first", "second", "final"):
        return tc
    if default_template in ("first", "second", "final"):
        return default_template
    rs = int(row.get("reminders_sent", 0) or 0)
    if rs <= 0:
        return "first"
    elif rs == 1:
        return "second"
    return "final"

def format_template(tpl: Dict[str, str], row: pd.Series) -> Tuple[str, str]:
    subject = tpl["subject"].format(
        name=row.get("name",""), invoice_no=row.get("invoice_no",""), amount=row.get("amount","")
    )
    body_html = tpl["body_html"].format(
        name=row.get("name",""), invoice_no=row.get("invoice_no",""), amount=row.get("amount","")
    )
    return subject, body_html

# ---------- UI ----------
st.title("📧 Otomatik Fatura Mailer (Web)")
st.caption("Excel yükle → şablon seç → Gönder. Microsoft Graph ile gönderim, CC desteği.")

col_left, col_right = st.columns([2,1])

with col_left:
    uploaded = st.file_uploader("Excel dosyasını yükle (xlsx)", type=["xlsx"])
    templates_path = st.text_input("Şablon dosya yolu (email_templates.json)", "email_templates.json")
    default_template = st.radio("Kullanılacak şablon", ["first", "second", "final"], index=0, horizontal=True)
    dry_run = st.toggle("Deneme modu (dry-run) – mail göndermez", value=False)
    send_btn = st.button("📨 SEND", type="primary", use_container_width=True)

with col_right:
    st.markdown("**İpucu:** Excel sütunları: `email, name, invoice_no, amount, invoice_pdf, status, last_sent, reminders_sent, template_choice, cc`")
    st.markdown("- `status=Paid` olanlar atlanır.")
    st.markdown("- `cc` birden çok adresi `;` veya `,` ile ayırın.")
    st.markdown("- Gönderim sonrası rapor kolonu eklenir: `report_note`, `last_template_sent`, `last_sent` ve `reminders_sent` güncellenir.")
    st.divider()
    if os.path.isfile(templates_path):
        tpl_json = load_templates(templates_path)
        st.subheader("Aktif Şablon Önizleme")
        tpl = tpl_json.get(default_template, {})
        st.write("**Subject**:", tpl.get("subject",""))
        st.markdown(tpl.get("body_html",""), unsafe_allow_html=True)

if uploaded:
    try:
        df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Excel okunamadı: {e}")
        st.stop()

    st.subheader("Excel Önizleme")
    st.dataframe(df, use_container_width=True, height=300)

    if send_btn:
        # Load templates
        if not os.path.isfile(templates_path):
            st.error("email_templates.json bulunamadı. Yol doğru mu?")
            st.stop()
        templates = load_templates(templates_path)

        # Token (skip if dry-run? still not needed)
        access_token = None
        if not dry_run:
            client_id, tenant_id = load_env()
            tok = get_token(client_id, tenant_id)
            access_token = tok["access_token"]

        sent_count = 0
        logs = []
        now_str = datetime.now().date().isoformat()

        # Ensure report columns exist
        if "report_note" not in df.columns:
            df["report_note"] = ""
        if "last_template_sent" not in df.columns:
            df["last_template_sent"] = ""

        for idx, row in df.iterrows():
            if str(row.get("status","")).lower() == "paid":
                logs.append(f"SKIP Paid: {row.get('email','')}")
                continue

            tkey = choose_template(row, default_template)
            tpl = templates.get(tkey)
            if not tpl:
                logs.append(f"ERROR template not found: {tkey} -> {row.get('email','')}")
                continue

            subject, body_html = format_template(tpl, row)
            pdf_path = str(row.get("invoice_pdf","")).strip()
            cc_list = parse_recipients(row.get("cc",""))

            if dry_run:
                logs.append(f"[DRY] {row.get('email','')} | {tkey} | {pdf_path} | CC:{[x['emailAddress']['address'] for x in cc_list]}")
            else:
                try:
                    payload = build_graph_message(str(row.get("email","")).strip(), subject, body_html, pdf_path, cc_list)
                    send_mail_graph(access_token, payload)
                    sent_count += 1
                    logs.append(f"OK {row.get('email','')} ({tkey})")
                except Exception as e:
                    logs.append(f"ERROR {row.get('email','')}: {e}")
                    continue

            # Update report columns
            df.at[idx, "last_sent"] = now_str
            rs = int(row.get("reminders_sent", 0) or 0)
            df.at[idx, "reminders_sent"] = min(rs + 1, 3)
            df.at[idx, "last_template_sent"] = tkey
            df.at[idx, "report_note"] = f"{tkey} template has been sent on {now_str}"

            # Optional per-template boolean columns
            colname = f"{tkey}_template_sent"
            if colname not in df.columns:
                df[colname] = False
            df.at[idx, colname] = True

        st.success(f"Gönderim tamamlandı. Adet: {sent_count}")
        st.text_area("Log", "\n".join(logs), height=200)

        # Prepare updated excel for download
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="invoices")
        out.seek(0)
        st.download_button("📥 Güncellenmiş Excel'i indir", out, file_name=f"invoice_mailer_updated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Üstten bir Excel dosyası yükleyin.")
#streamlit run app.py
