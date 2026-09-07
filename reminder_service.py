import os
import re
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Dict, List, Any, Tuple, Optional

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE_PATH = os.path.join(WORKSPACE_DIR, "reminder_config.json")

DEFAULT_CONFIG = {
    "email": {
        "smtp_host": os.environ.get("SMTP_HOST", "smtp.gmail.com"),
        "smtp_port": int(os.environ.get("SMTP_PORT", 587)),
        "use_tls": True,
        "sender_email": os.environ.get("SENDER_EMAIL", ""),
        "sender_password": os.environ.get("SENDER_PASSWORD", ""),
        "sender_name": os.environ.get("SENDER_NAME", "QA FECT - Phòng Đảm Bảo Chất Lượng"),
    },
    "auditors": {
        "ThaoTDP": {
            "name": "Trần Đoàn Phương Thảo",
            "account": "ThaoTDP",
            "email": "thaotdp@fe.edu.vn",
            "enabled": True,
        },
        "DiCQ": {
            "name": "Chung Quốc Di",
            "account": "DiCQ",
            "email": "dicq@fe.edu.vn",
            "enabled": True,
        },
        "ThanhNTD6": {
            "name": "Nguyễn Thị Đan Thanh",
            "account": "ThanhNTD6",
            "email": "thanhntd6@fe.edu.vn",
            "enabled": True,
        },
    },
    "auto_scanner": {
        "enabled": True,
        "scan_time": "08:30",
        "remind_days_in_advance": 1,
        "remind_monday_on_friday": True,
        "last_scan_date": None,
        "last_scan_at": None,
        "last_scan_summary": None,
    }
}


def load_config() -> Dict[str, Any]:
    """Loads configuration from reminder_config.json or initializes with defaults."""
    if not os.path.exists(CONFIG_FILE_PATH):
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()
    try:
        with open(CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
            merged = DEFAULT_CONFIG.copy()
            for k, v in data.items():
                if isinstance(v, dict) and k in merged:
                    merged[k] = {**merged[k], **v}
                else:
                    merged[k] = v
            # Ensure any legacy zalo keys are cleaned out
            merged.pop("zalo_group", None)
            merged.pop("zalo_oa", None)
            return merged
    except Exception:
        return DEFAULT_CONFIG.copy()


def save_config(config: Dict[str, Any]) -> bool:
    """Saves configuration securely to reminder_config.json."""
    try:
        # Strip legacy zalo keys if present
        config_to_save = dict(config)
        config_to_save.pop("zalo_group", None)
        config_to_save.pop("zalo_oa", None)
        with open(CONFIG_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(config_to_save, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


# ==============================================================================
# EMAIL SERVICE (Google / Gmail SMTP & @fe.edu.vn)
# ==============================================================================

def build_email_html(event: Dict[str, Any], auditor_info: Optional[Dict[str, Any]] = None) -> str:
    """
    Builds a professional HTML email template for audit reminder sent to
    both QA Auditor and Unit Representatives (Column C).
    """
    auditor = re.sub(r"ThanhNTD(?!6)", "ThanhNTD6", event.get("auditor", ""))
    auditor_name = auditor_info.get("name", auditor) if auditor_info else auditor
    date_str = event.get("date", "")
    time_str = event.get("time", "")
    session_str = event.get("session", "")
    unit_str = event.get("unit_display", "")
    dept_str = event.get("dept_display", "")
    task_count = event.get("task_count", 0)

    reps_display = event.get("representatives_display", "")
    if not reps_display or reps_display == "Chưa có thông tin":
        reps_display = "Theo phân công của Đơn vị"

    docs = event.get("documents", [])
    contents = event.get("contents", [])
    record_times = event.get("record_times", [])
    record_time_display = event.get("record_time_display", "")

    # Format documents list
    docs_html = ""
    if docs:
        docs_items = "".join([f"<li style='margin-bottom:6px; color:#334155;'><strong>{d}</strong></li>" for d in docs])
        docs_html = f"""
        <div style="margin-top:16px; padding:12px 16px; background-color:#F8FAFC; border-left:4px solid #3B82F6; border-radius:4px;">
            <h4 style="margin:0 0 8px 0; color:#1E3A8A; font-size:14px; text-transform:uppercase;">📂 Danh Mục Tài Liệu Cần Kiểm Tra:</h4>
            <ul style="margin:0; padding-left:20px; font-size:13px; line-height:1.6;">
                {docs_items}
            </ul>
        </div>
        """

    # Format contents list
    contents_html = ""
    if contents:
        contents_items = "".join([f"<li style='margin-bottom:8px; color:#1E293B;'>{c}</li>" for c in contents[:8]])
        contents_html = f"""
        <div style="margin-top:16px; padding:12px 16px; background-color:#FEFCE8; border-left:4px solid #EAB308; border-radius:4px;">
            <h4 style="margin:0 0 8px 0; color:#854D0E; font-size:14px; text-transform:uppercase;">📝 Nội Dung & Checklist Kiểm Soát Trọng Tâm:</h4>
            <ol style="margin:0; padding-left:20px; font-size:13px; line-height:1.6;">
                {contents_items}
            </ol>
        </div>
        """

    # Format record times (Thời gian phát sinh hồ sơ)
    record_times_html = ""
    if record_times:
        rt_items = "".join([f"<li style='margin-bottom:6px; color:#065F46;'><strong>{rt}</strong></li>" for rt in record_times])
        record_times_html = f"""
        <div style="margin-top:16px; padding:12px 16px; background-color:#ECFDF5; border-left:4px solid #10B981; border-radius:4px;">
            <h4 style="margin:0 0 8px 0; color:#065F46; font-size:14px; text-transform:uppercase;">⏳ Thời Gian Phát Sinh Hồ Sơ Cần Kiểm Soát:</h4>
            <ul style="margin:0; padding-left:20px; font-size:13px; line-height:1.6;">
                {rt_items}
            </ul>
        </div>
        """

    record_time_row = ""
    if record_time_display:
        record_time_row = f"""
                    <tr>
                        <td class="label">⏳ Thời gian phát sinh hồ sơ:</td>
                        <td class="value"><span style="background:#ECFDF5; color:#065F46; padding:3px 8px; border-radius:4px; font-weight:600;">{record_time_display}</span></td>
                    </tr>
        """

    # Personal salutation
    if reps_display != "Theo phân công của Đơn vị":
        salutation = f"Kính gửi Cán bộ Đánh giá <strong>{auditor_name}</strong> và Quý Anh/Chị đại diện đơn vị <strong>{reps_display}</strong>,"
    else:
        salutation = f"Kính gửi Cán bộ Đánh giá <strong>{auditor_name}</strong> và Ban Lãnh đạo/Đại diện đơn vị <strong>{unit_str}</strong>,"

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Arial, sans-serif; background-color: #F1F5F9; margin: 0; padding: 20px; }}
            .container {{ max-width: 680px; margin: 0 auto; background: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border: 1px solid #E2E8F0; }}
            .header {{ background: linear-gradient(135deg, #1E40AF 0%, #2563EB 100%); color: #ffffff; padding: 24px; text-align: center; }}
            .content {{ padding: 24px; }}
            .info-table {{ width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 14px; }}
            .info-table td {{ padding: 10px 12px; border-bottom: 1px solid #E2E8F0; }}
            .info-table td.label {{ width: 36%; color: #64748B; font-weight: 600; background-color: #F8FAFC; }}
            .info-table td.value {{ color: #0F172A; font-weight: 700; }}
            .footer {{ background-color: #F8FAFC; padding: 16px; text-align: center; font-size: 12px; color: #64748B; border-top: 1px solid #E2E8F0; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <div style="font-size:12px; letter-spacing:1px; text-transform:uppercase; opacity:0.9; margin-bottom:4px;">Phòng Đảm Bảo Chất Lượng FECT (QA FECT)</div>
                <h2 style="margin:0; font-size:20px; font-weight:700;">THÔNG BÁO LỊCH KIỂM SOÁT NGÀY MAI</h2>
                <div style="font-size:14px; margin-top:6px; opacity:0.95;">Ngày làm việc: <strong>{date_str}</strong></div>
            </div>
            <div class="content">
                <p style="margin-top:0; font-size:15px; color:#334155; line-height:1.5;">
                    {salutation}
                </p>
                <p style="font-size:14px; color:#475569; line-height:1.6;">
                    Phòng Đảm Bảo Chất Lượng FECT xin thông báo lịch kiểm soát, đánh giá nội bộ đã được phê duyệt cho ngày mai như sau:
                </p>

                <table class="info-table">
                    <tr>
                        <td class="label">👤 Cán bộ đánh giá QA:</td>
                        <td class="value"><span style="color:#1E40AF;">{auditor_name} ({auditor})</span></td>
                    </tr>
                    <tr>
                        <td class="label">👥 Đại diện đơn vị:</td>
                        <td class="value"><span style="background:#EEF2FF; color:#4338CA; padding:3px 8px; border-radius:4px;">{reps_display}</span></td>
                    </tr>
                    <tr>
                        <td class="label">📅 Ngày kiểm soát:</td>
                        <td class="value" style="color:#2563EB;">{date_str}</td>
                    </tr>
                    <tr>
                        <td class="label">⏰ Khung giờ:</td>
                        <td class="value"><span style="background:#FEF3C7; color:#92400E; padding:3px 8px; border-radius:4px;">{time_str} ({session_str})</span></td>
                    </tr>
                    <tr>
                        <td class="label">🏢 Đơn vị / Phân hiệu:</td>
                        <td class="value">{unit_str}</td>
                    </tr>
                    <tr>
                        <td class="label">📍 Phòng ban / Bộ phận:</td>
                        <td class="value">{dept_str or "Toàn bộ theo checklist"}</td>
                    </tr>
                    {record_time_row}
                    <tr>
                        <td class="label">📊 Khối lượng:</td>
                        <td class="value">{task_count} đầu việc kiểm tra</td>
                    </tr>
                </table>

                {record_times_html}
                {docs_html}
                {contents_html}

                <div style="margin-top:20px; padding:14px; background-color:#EFF6FF; border-left:4px solid #3B82F6; border-radius:4px; font-size:13px; color:#1E40AF; line-height:1.6;">
                    💡 <strong>Lưu ý phối hợp:</strong><br>
                    • Đề nghị Cán bộ đánh giá và Quý Đơn vị phối hợp chuẩn bị đầy đủ hồ sơ, minh chứng và bố trí nhân sự làm việc đúng khung giờ quy định.<br>
                    • Trường hợp có thay đổi đột xuất, vui lòng phản hồi ngay qua email này để Phòng QA kịp thời hỗ trợ.
                </div>
            </div>
            <div class="footer">
                Email tự động được gửi từ Hệ Thống Lập Kế Hoạch & Giám Sát QA FECT.<br>
                Mọi thắc mắc xin vui lòng liên hệ Phòng Đảm Bảo Chất Lượng FECT.
            </div>
        </div>
    </body>
    </html>
    """
    return html


def send_email_smtp(
    to_email: str,
    subject: str,
    html_content: str,
    config: Optional[Dict[str, Any]] = None
) -> Tuple[bool, str]:
    """
    Sends an email to one or multiple recipients using Gmail / Google Workspace (smtp.gmail.com)
    or custom SMTP. Handles comma-separated email lists.
    """
    if not config:
        config = load_config()

    email_cfg = config.get("email", {})
    smtp_host = email_cfg.get("smtp_host", "smtp.gmail.com").strip()
    smtp_port = int(email_cfg.get("smtp_port", 587))
    sender_email = email_cfg.get("sender_email", "").strip()
    sender_password = email_cfg.get("sender_password", "").replace(" ", "").strip()
    sender_name = email_cfg.get("sender_name", "QA FECT").strip()

    if not sender_email or not sender_password:
        return False, "Chưa cấu hình Địa chỉ Email gửi hoặc Mật khẩu ứng dụng trong phần Cài đặt!"

    if not to_email:
        return False, "Chưa có địa chỉ email người nhận!"

    # Parse and validate recipient list
    raw_list = [e.strip() for e in to_email.replace(";", ",").split(",")]
    recipients = [e for e in raw_list if e and "@" in e]
    if not recipients:
        return False, f"Không tìm thấy địa chỉ email hợp lệ nào trong: {to_email}"

    # Remove duplicates preserving order
    recipients = list(dict.fromkeys(recipients))

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = f"{sender_name} <{sender_email}>"
        msg["To"] = ", ".join(recipients)

        part = MIMEText(html_content, "html", "utf-8")
        msg.attach(part)

        # Connect to SMTP server
        server = smtplib.SMTP(smtp_host, smtp_port, timeout=20)
        server.ehlo()
        if email_cfg.get("use_tls", True):
            server.starttls()
            server.ehlo()

        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipients, msg.as_string())
        server.quit()

        return True, f"Đã gửi email nhắc lịch thành công tới {len(recipients)} người nhận: {', '.join(recipients)}"
    except smtplib.SMTPAuthenticationError:
        return False, (
            "Lỗi xác thực SMTP: Sai email hoặc mật khẩu ứng dụng! "
            "Đối với tài khoản Google / Gmail (hoặc Google Workspace), bạn cần bật 'Xác minh 2 bước' (2-Step Verification) "
            "và tạo 'Mật khẩu ứng dụng (App Password)' 16 ký tự tại: https://myaccount.google.com/apppasswords"
        )
    except Exception as e:
        return False, f"Lỗi gửi email ({type(e).__name__}): {str(e)}"


# ==============================================================================
# SENT REMINDERS LOGGING & IDEMPOTENCY
# ==============================================================================
import datetime

SENT_LOG_PATH = os.path.join(WORKSPACE_DIR, "sent_reminders.json")


def load_sent_log() -> List[Dict[str, Any]]:
    """Loads history of sent reminders from sent_reminders.json."""
    if not os.path.exists(SENT_LOG_PATH):
        return []
    try:
        with open(SENT_LOG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get("history", []) if isinstance(data, dict) else []
    except Exception:
        return []


def log_sent_reminder(entry: Dict[str, Any]) -> bool:
    """Appends a reminder send record to sent_reminders.json."""
    try:
        history = load_sent_log()
        history.append(entry)
        if len(history) > 500:
            history = history[-500:]
        with open(SENT_LOG_PATH, "w", encoding="utf-8") as f:
            json.dump({"history": history}, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def is_already_sent(event_id: str, date_str: str) -> bool:
    """Checks if an event on target date was already successfully sent."""
    history = load_sent_log()
    for item in history:
        if item.get("event_id") == event_id and item.get("status") == "SUCCESS":
            return True
    return False


def clear_sent_log() -> bool:
    """Clears sent reminder history."""
    try:
        with open(SENT_LOG_PATH, "w", encoding="utf-8") as f:
            json.dump({"history": []}, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


# ==============================================================================
# AUTO-SCAN & AUTO-SEND (1 Day in Advance)
# ==============================================================================

def scan_and_send_reminders(
    plan_data: Dict[str, Any],
    config: Optional[Dict[str, Any]] = None,
    target_scan_date: Optional[datetime.date] = None,
    force: bool = False
) -> Dict[str, Any]:
    """
    Scans plan events matching target date (default tomorrow = today + 1 day),
    and automatically sends reminder emails to QA Auditor and Column C Representatives.
    """
    if not config:
        config = load_config()
    if target_scan_date is None:
        target_scan_date = datetime.date.today()

    scanner_cfg = config.get("auto_scanner", {})
    days_in_advance = int(scanner_cfg.get("remind_days_in_advance", 1))

    # Primary reminder date: scan_date + days_in_advance
    primary_target = target_scan_date + datetime.timedelta(days=days_in_advance)
    target_dates = [primary_target]

    # If scan date is Friday (weekday == 4) and remind_monday_on_friday is True:
    # also remind Saturday (+1), Sunday (+2), and Monday (+3)
    if target_scan_date.weekday() == 4 and scanner_cfg.get("remind_monday_on_friday", True):
        for extra_days in [2, 3]:
            extra_d = target_scan_date + datetime.timedelta(days=extra_days)
            if extra_d not in target_dates:
                target_dates.append(extra_d)

    target_date_strs = [f"{d.day:02d}.{d.month:02d}.{d.year}" for d in target_dates]

    events = plan_data.get("events", [])
    matched_events = [e for e in events if e.get("date") in target_date_strs]

    auditors_cfg = config.get("auditors", {})
    now_iso = datetime.datetime.now().isoformat()

    results = []
    sent_count = 0
    skipped_count = 0
    failed_count = 0

    for ev in matched_events:
        ev_id = ev.get("id")
        ev_date = ev.get("date")

        # Check idempotency (already sent)
        if not force and is_already_sent(ev_id, ev_date):
            skipped_count += 1
            results.append({
                "event_id": ev_id,
                "date": ev_date,
                "unit": ev.get("unit_display"),
                "auditor": ev.get("auditor"),
                "status": "SKIPPED",
                "message": "Đã gửi email nhắc việc trước đó (bỏ qua để tránh trùng lặp)",
                "recipients": []
            })
            continue

        # Auditor info lookup
        raw_aud = re.sub(r"ThanhNTD(?!6)", "ThanhNTD6", ev.get("auditor", ""))
        if raw_aud in auditors_cfg:
            auditor_info = auditors_cfg[raw_aud]
        elif "+" in raw_aud:
            parts = [re.sub(r"ThanhNTD(?!6)", "ThanhNTD6", p.strip()) for p in raw_aud.split("+")]
            names = [auditors_cfg.get(p, {}).get("name", p) for p in parts]
            emails = [auditors_cfg.get(p, {}).get("email", "") for p in parts if auditors_cfg.get(p, {}).get("email")]
            auditor_info = {
                "name": " & ".join(names),
                "email": ", ".join(emails),
            }
        else:
            auditor_info = {}

        # Recipients
        qa_emails = [e.strip() for e in auditor_info.get("email", "").split(",") if e.strip() and "@" in e]
        rep_emails = [e.strip() for e in ev.get("representative_emails", []) if e.strip() and "@" in e]
        all_recipients = list(dict.fromkeys(qa_emails + rep_emails))

        if not all_recipients:
            failed_count += 1
            res_entry = {
                "event_id": ev_id,
                "date": ev_date,
                "unit": ev.get("unit_display"),
                "auditor": raw_aud,
                "status": "FAILED",
                "message": f"Không có địa chỉ email người nhận (Cán bộ QA hoặc Đại diện đơn vị)",
                "recipients": []
            }
            results.append(res_entry)
            log_sent_reminder({
                **res_entry,
                "sent_at": now_iso,
                "mode": "AUTO"
            })
            continue

        recipients_str = ", ".join(all_recipients)
        subject = f"[QA FECT] Nhắc lịch kiểm soát ngày mai: {ev.get('date')} - {ev.get('unit_display')} ({raw_aud})"
        html = build_email_html(ev, auditor_info)

        ok, msg = send_email_smtp(recipients_str, subject, html, config)
        status_str = "SUCCESS" if ok else "FAILED"
        if ok:
            sent_count += 1
        else:
            failed_count += 1

        res_entry = {
            "event_id": ev_id,
            "date": ev_date,
            "unit": ev.get("unit_display"),
            "auditor": raw_aud,
            "status": status_str,
            "message": msg,
            "recipients": all_recipients
        }
        results.append(res_entry)
        log_sent_reminder({
            **res_entry,
            "sent_at": now_iso,
            "mode": "AUTO"
        })

    return {
        "success": True,
        "scan_date": f"{target_scan_date.day:02d}.{target_scan_date.month:02d}.{target_scan_date.year}",
        "target_dates": target_date_strs,
        "total_matched": len(matched_events),
        "sent_count": sent_count,
        "skipped_count": skipped_count,
        "failed_count": failed_count,
        "results": results
    }
