import os
import re
import io
from typing import Optional, List
from fastapi import FastAPI, File, UploadFile, Form, Query, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from excel_processor import (
    get_matrix_info,
    extract_tasks_for_month,
    generate_control_plan_workbook,
    DEFAULT_MATRIX_PATH,
    DEFAULT_TEMPLATE_PATH,
)

import sys

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()

app = FastAPI(title="Hệ Thống Lập Kế Hoạch Kiểm Soát FECT")

# In-memory storage for uploaded matrix file (falls back to default if None)
CURRENT_MATRIX_BYTES: Optional[bytes] = None
CURRENT_MATRIX_FILENAME: str = os.path.basename(DEFAULT_MATRIX_PATH)


class ExportRequest(BaseModel):
    month: int = 8
    year: int = 2026
    fsc_mode: str = "fsc_ct"  # "fsc_ct" or "all_fsc"
    include_calendar: bool = True
    include_empty_sheets: bool = True
    selected_units: Optional[List[str]] = None
    auto_schedule: bool = True


@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_path = os.path.join(BASE_DIR, "templates", "index.html")
    if not os.path.exists(index_path):
        # Fallback to local templates
        index_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "index.html")
    if not os.path.exists(index_path):
        raise HTTPException(status_code=404, detail="Template not found")
    with open(index_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/api/matrix-info")
async def api_matrix_info():
    global CURRENT_MATRIX_BYTES, CURRENT_MATRIX_FILENAME
    source = CURRENT_MATRIX_BYTES if CURRENT_MATRIX_BYTES is not None else DEFAULT_MATRIX_PATH
    try:
        info = get_matrix_info(source)
        return {
            "success": True,
            "filename": CURRENT_MATRIX_FILENAME,
            "is_default": CURRENT_MATRIX_BYTES is None,
            "units": info["units"],
            "month_totals": info["month_totals"],
            "all_months": info["all_months"],
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})


@app.post("/api/upload-matrix")
async def api_upload_matrix(file: UploadFile = File(...)):
    global CURRENT_MATRIX_BYTES, CURRENT_MATRIX_FILENAME
    if not file.filename.endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Vui lòng tải lên file Excel (.xlsx)")

    content = await file.read()
    try:
        # Validate that the file has unit sheets
        info = get_matrix_info(content)
        if not info["units"]:
            raise ValueError("Không tìm thấy sheet đơn vị hợp lệ nào trong file (yêu cầu cột tháng T1..T12 ở dòng 2)")

        CURRENT_MATRIX_BYTES = content
        CURRENT_MATRIX_FILENAME = file.filename

        return {
            "success": True,
            "filename": CURRENT_MATRIX_FILENAME,
            "is_default": False,
            "units": info["units"],
            "month_totals": info["month_totals"],
            "all_months": info["all_months"],
            "message": f"Tải lên thành công file '{file.filename}' với {len(info['units'])} đơn vị!",
        }
    except Exception as e:
        return JSONResponse(status_code=400, content={"success": False, "error": f"Lỗi đọc file: {str(e)}"})


@app.post("/api/reset-matrix")
async def api_reset_matrix():
    global CURRENT_MATRIX_BYTES, CURRENT_MATRIX_FILENAME
    CURRENT_MATRIX_BYTES = None
    CURRENT_MATRIX_FILENAME = os.path.basename(DEFAULT_MATRIX_PATH)
    info = get_matrix_info(DEFAULT_MATRIX_PATH)
    return {
        "success": True,
        "filename": CURRENT_MATRIX_FILENAME,
        "is_default": True,
        "units": info["units"],
        "month_totals": info["month_totals"],
        "all_months": info["all_months"],
        "message": "Đã đặt lại về file ma trận mặc định!",
    }


@app.get("/api/preview")
async def api_preview(month: str = Query("T8"), auto_schedule: bool = Query(True)):
    global CURRENT_MATRIX_BYTES
    source = CURRENT_MATRIX_BYTES if CURRENT_MATRIX_BYTES is not None else DEFAULT_MATRIX_PATH
    try:
        from scheduler import get_auditor, schedule_tasks
        tasks = extract_tasks_for_month(source, month)
        m_num = int(month.upper().replace("T", ""))

        # Prepare tasks with target sheet unit names (map FSC -> FSC CT for preview)
        prepared_tasks = {}
        for u, t_list in tasks.items():
            target_u = "FSC CT" if u == "FSC" else u
            u_tasks = []
            for t in t_list:
                t_copy = dict(t)
                t_copy["unit"] = target_u
                t_copy["scheduled_auditor"] = get_auditor(target_u, t.get("department", ""))
                u_tasks.append(t_copy)
            prepared_tasks[target_u] = u_tasks

        if auto_schedule:
            prepared_tasks, _ = schedule_tasks(prepared_tasks, 2026, m_num)

        total_tasks = sum(len(t) for t in prepared_tasks.values())

        # Clean tasks for JSON serialization (remove datetime objects and raw openpyxl values)
        json_units = {}
        for u, t_list in prepared_tasks.items():
            u_clean = []
            for t in t_list:
                u_clean.append({
                    "row_index": t.get("row_index"),
                    "unit": t.get("unit"),
                    "department": t.get("department"),
                    "document": t.get("document"),
                    "content": t.get("content"),
                    "detail": t.get("detail"),
                    "department_html": t.get("department_html"),
                    "document_html": t.get("document_html"),
                    "content_html": t.get("content_html"),
                    "detail_html": t.get("detail_html"),
                    "scheduled_date": t.get("scheduled_date", ""),
                    "scheduled_time": t.get("scheduled_time", ""),
                    "scheduled_auditor": t.get("scheduled_auditor") or get_auditor(u, t.get("department", "")),
                    "scheduled_session": t.get("scheduled_session", ""),
                })
            json_units[u] = u_clean

        return {
            "success": True,
            "month": month.upper(),
            "total_tasks": total_tasks,
            "units": json_units,
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})


@app.post("/api/export")
async def api_export(req: ExportRequest):
    global CURRENT_MATRIX_BYTES
    source = CURRENT_MATRIX_BYTES if CURRENT_MATRIX_BYTES is not None else DEFAULT_MATRIX_PATH
    try:
        output_stream = generate_control_plan_workbook(
            matrix_source=source,
            template_source=DEFAULT_TEMPLATE_PATH,
            month_num=req.month,
            year=req.year,
            fsc_mode=req.fsc_mode,
            include_calendar=req.include_calendar,
            selected_units=req.selected_units,
            include_empty_sheets=req.include_empty_sheets,
            auto_schedule=req.auto_schedule,
        )

        month_str = f"T{req.month:02d}"
        download_filename = f"Ke_hoach_kiem_soat_{month_str}_{req.year}.xlsx"

        headers = {
            "Content-Disposition": f'attachment; filename="{download_filename}"',
            "Access-Control-Expose-Headers": "Content-Disposition",
        }

        return StreamingResponse(
            output_stream,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})


# ==============================================================================
# REMINDER MODULE ENDPOINTS (Email @fe.edu.vn)
# ==============================================================================
import json
import datetime
from plan_parser import parse_plan_excel
import reminder_service

CURRENT_PLAN_DATA: Optional[dict] = None
CURRENT_PLAN_FILENAME: str = "Ke_hoach_kiem_soat_T08_2026.xlsx"


class ReminderSendRequest(BaseModel):
    event_id: str
    recipient_emails: Optional[str] = None


class TestEmailRequest(BaseModel):
    test_email: str
    smtp_host: str = "smtp.gmail.com"
    smtp_port: int = 587
    sender_email: str
    sender_password: str
    sender_name: str = "QA FECT"


def ensure_plan_loaded() -> dict:
    global CURRENT_PLAN_DATA, CURRENT_PLAN_FILENAME
    if CURRENT_PLAN_DATA is None:
        default_plan_path = os.path.join(BASE_DIR, "Ke_hoach_kiem_soat_T08_2026.xlsx")
        if os.path.exists(default_plan_path):
            with open(default_plan_path, "rb") as f:
                CURRENT_PLAN_DATA = parse_plan_excel(f.read())
                CURRENT_PLAN_FILENAME = "Ke_hoach_kiem_soat_T08_2026.xlsx"
        else:
            CURRENT_PLAN_DATA = {
                "success": False,
                "total_tasks": 0,
                "total_events": 0,
                "dates": [],
                "events": [],
                "events_by_date": {}
            }
    return CURRENT_PLAN_DATA


@app.get("/api/reminders/current-plan")
async def api_get_current_plan():
    plan = ensure_plan_loaded()
    config = reminder_service.load_config()
    return {
        "success": True,
        "filename": CURRENT_PLAN_FILENAME,
        "month": plan.get("month", ""),
        "year": plan.get("year", 2026),
        "total_tasks": plan.get("total_tasks", 0),
        "total_events": plan.get("total_events", 0),
        "dates": plan.get("dates", []),
        "events": plan.get("events", []),
        "events_by_date": plan.get("events_by_date", {}),
        "auditors_config": config.get("auditors", {}),
    }


@app.post("/api/reminders/upload-plan")
async def api_upload_plan(file: UploadFile = File(...)):
    global CURRENT_PLAN_DATA, CURRENT_PLAN_FILENAME
    if not file.filename.endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Vui lòng tải lên file Excel kế hoạch (.xlsx)")

    content = await file.read()
    try:
        parsed = parse_plan_excel(content)
        if parsed.get("total_events", 0) == 0:
            raise ValueError("Không tìm thấy dữ liệu ca kiểm soát có Ngày và CB đánh giá trong file kế hoạch!")

        CURRENT_PLAN_DATA = parsed
        CURRENT_PLAN_FILENAME = file.filename

        return {
            "success": True,
            "filename": CURRENT_PLAN_FILENAME,
            "month": parsed.get("month", ""),
            "year": parsed.get("year", 2026),
            "total_tasks": parsed.get("total_tasks", 0),
            "total_events": parsed.get("total_events", 0),
            "dates": parsed.get("dates", []),
            "message": f"Nạp thành công file kế hoạch '{file.filename}' với {parsed.get('total_events')} ca kiểm soát!",
        }
    except Exception as e:
        return JSONResponse(status_code=400, content={"success": False, "error": f"Lỗi đọc file kế hoạch: {str(e)}"})


@app.get("/api/reminders/config")
async def api_get_reminder_config():
    cfg = reminder_service.load_config()
    safe_cfg = json.loads(json.dumps(cfg))
    if safe_cfg.get("email", {}).get("sender_password"):
        safe_cfg["email"]["has_password"] = True
        safe_cfg["email"]["sender_password"] = "******"
    else:
        safe_cfg["email"]["has_password"] = False
    return {"success": True, "config": safe_cfg}


@app.post("/api/reminders/config")
async def api_save_reminder_config(req: dict):
    current = reminder_service.load_config()
    if req.get("email", {}).get("sender_password") == "******":
        req["email"]["sender_password"] = current.get("email", {}).get("sender_password", "")
    ok = reminder_service.save_config(req)
    if ok:
        return {"success": True, "message": "Đã lưu cấu hình nhắc lịch thành công!"}
    return JSONResponse(status_code=500, content={"success": False, "error": "Không thể ghi file cấu hình"})


@app.get("/api/reminders/preview/{event_id}")
async def api_preview_reminder_message(event_id: str):
    plan = ensure_plan_loaded()
    events = plan.get("events", [])
    target_event = next((e for e in events if e.get("id") == event_id), None)
    if not target_event:
        raise HTTPException(status_code=404, detail="Không tìm thấy ca kiểm soát")

    config = reminder_service.load_config()
    auditor = re.sub(r"ThanhNTD(?!6)", "ThanhNTD6", target_event.get("auditor", ""))
    auditors_cfg = config.get("auditors", {})
    if auditor in auditors_cfg:
        auditor_info = auditors_cfg[auditor]
    elif "+" in auditor:
        parts = [re.sub(r"ThanhNTD(?!6)", "ThanhNTD6", p.strip()) for p in auditor.split("+")]
        names = [auditors_cfg.get(p, {}).get("name", p) for p in parts]
        emails = [auditors_cfg.get(p, {}).get("email", "") for p in parts if auditors_cfg.get(p, {}).get("email")]
        auditor_info = {
            "name": " & ".join(names),
            "email": ", ".join(emails),
        }
    else:
        auditor_info = {}

    qa_emails = [e.strip() for e in auditor_info.get("email", "").split(",") if e.strip() and "@" in e]
    rep_emails = [e.strip() for e in target_event.get("representative_emails", []) if e.strip() and "@" in e]
    all_recipients = list(dict.fromkeys(qa_emails + rep_emails))

    email_html = reminder_service.build_email_html(target_event, auditor_info)
    subject = f"[QA FECT] Nhắc lịch kiểm soát ngày mai: {target_event.get('date')} - {target_event.get('unit_display')} ({auditor})"

    return {
        "success": True,
        "event": target_event,
        "auditor_info": auditor_info,
        "qa_emails": qa_emails,
        "rep_emails": rep_emails,
        "recipient_emails": ", ".join(all_recipients),
        "email_subject": subject,
        "email_html": email_html,
        "email_configured": bool(config.get("email", {}).get("sender_email") and config.get("email", {}).get("sender_password")),
    }


@app.post("/api/reminders/send-email")
async def api_send_reminder_email(req: ReminderSendRequest):
    plan = ensure_plan_loaded()
    events = plan.get("events", [])
    target_event = next((e for e in events if e.get("id") == req.event_id), None)
    if not target_event:
        raise HTTPException(status_code=404, detail="Không tìm thấy ca kiểm soát")

    config = reminder_service.load_config()
    auditor = re.sub(r"ThanhNTD(?!6)", "ThanhNTD6", target_event.get("auditor", ""))
    auditors_cfg = config.get("auditors", {})
    if auditor in auditors_cfg:
        auditor_info = auditors_cfg[auditor]
    elif "+" in auditor:
        parts = [re.sub(r"ThanhNTD(?!6)", "ThanhNTD6", p.strip()) for p in auditor.split("+")]
        names = [auditors_cfg.get(p, {}).get("name", p) for p in parts]
        emails = [auditors_cfg.get(p, {}).get("email", "") for p in parts if auditors_cfg.get(p, {}).get("email")]
        auditor_info = {
            "name": " & ".join(names),
            "email": ", ".join(emails),
        }
    else:
        auditor_info = {}

    # Target emails: either manually customized or automatically combined (QA + Column C Representatives)
    if req.recipient_emails and req.recipient_emails.strip():
        target_emails = req.recipient_emails.strip()
    else:
        qa_emails = [e.strip() for e in auditor_info.get("email", "").split(",") if e.strip() and "@" in e]
        rep_emails = [e.strip() for e in target_event.get("representative_emails", []) if e.strip() and "@" in e]
        all_recipients = list(dict.fromkeys(qa_emails + rep_emails))
        target_emails = ", ".join(all_recipients)

    if not target_emails:
        return JSONResponse(
            status_code=400,
            content={"success": False, "error": f"Không tìm thấy địa chỉ email của Cán bộ QA ({auditor}) hoặc Đại diện đơn vị để gửi!"}
        )

    subject = f"[QA FECT] Nhắc lịch kiểm soát ngày mai: {target_event.get('date')} - {target_event.get('unit_display')} ({auditor})"
    email_html = reminder_service.build_email_html(target_event, auditor_info)

    ok, msg = reminder_service.send_email_smtp(target_emails, subject, email_html, config)
    if ok:
        # Also log manual send into sent_reminders.json
        reminder_service.log_sent_reminder({
            "event_id": target_event.get("id"),
            "date": target_event.get("date"),
            "unit": target_event.get("unit_display"),
            "auditor": auditor,
            "status": "SUCCESS",
            "message": msg,
            "recipients": [e.strip() for e in target_emails.split(",") if e.strip()],
            "sent_at": datetime.datetime.now().isoformat(),
            "mode": "MANUAL"
        })
        return {"success": True, "message": msg}
    return JSONResponse(status_code=400, content={"success": False, "error": msg})


@app.post("/api/reminders/test-email")
async def api_test_email(req: TestEmailRequest):
    test_cfg = {
        "email": {
            "smtp_host": req.smtp_host.strip(),
            "smtp_port": req.smtp_port,
            "use_tls": True,
            "sender_email": req.sender_email.strip(),
            "sender_password": req.sender_password.replace(" ", "").strip(),
            "sender_name": req.sender_name.strip(),
        }
    }
    subject = "[QA FECT] Thử nghiệm kết nối máy chủ gửi Email (Google / Gmail SMTP)"
    html = f"""
    <div style="font-family:Arial,sans-serif; padding:20px; border:1px solid #E2E8F0; border-radius:8px;">
        <h3 style="color:#2563EB;">Kết Nối Email Máy Chủ Google / Gmail Thành Công!</h3>
        <p>Hệ thống lập kế hoạch & nhắc lịch kiểm soát FECT đã kết nối thành công với hòm thư <strong>{req.sender_email}</strong> qua máy chủ Google / Gmail SMTP.</p>
        <p style="color:#64748B; font-size:12px;">Thời gian thử nghiệm: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</p>
    </div>
    """
    ok, msg = reminder_service.send_email_smtp(req.test_email, subject, html, test_cfg)
    if ok:
        return {"success": True, "message": f"Kết nối máy chủ thành công! Thư thử nghiệm đã được gửi tới {req.test_email}."}
    return JSONResponse(status_code=400, content={"success": False, "error": msg})


# ==============================================================================
# AUTO-SCANNER CONTROLLER & BACKGROUND DAEMON
# ==============================================================================

class RunAutoScanRequest(BaseModel):
    simulate_date: Optional[str] = None  # e.g. "07.10.2026" or "2026-10-07"
    force: bool = False


class AutoScannerConfigRequest(BaseModel):
    enabled: bool = True
    scan_time: str = "08:30"
    remind_days_in_advance: int = 1
    remind_monday_on_friday: bool = True


@app.get("/api/reminders/auto-scanner-status")
async def api_get_auto_scanner_status():
    cfg = reminder_service.load_config()
    scanner_cfg = cfg.get("auto_scanner", {
        "enabled": True,
        "scan_time": "08:30",
        "remind_days_in_advance": 1,
        "remind_monday_on_friday": True,
    })
    now = datetime.datetime.now()
    history = reminder_service.load_sent_log()
    return {
        "success": True,
        "scanner": scanner_cfg,
        "system_date": now.strftime("%d.%m.%Y"),
        "system_time": now.strftime("%H:%M:%S"),
        "total_sent_all_time": len(history),
        "recent_logs": history[-20:],
    }


@app.post("/api/reminders/auto-scanner-config")
async def api_save_auto_scanner_config(req: AutoScannerConfigRequest):
    cfg = reminder_service.load_config()
    current_scanner = cfg.get("auto_scanner", {})
    current_scanner.update({
        "enabled": req.enabled,
        "scan_time": req.scan_time.strip(),
        "remind_days_in_advance": req.remind_days_in_advance,
        "remind_monday_on_friday": req.remind_monday_on_friday,
    })
    cfg["auto_scanner"] = current_scanner
    ok = reminder_service.save_config(cfg)
    if ok:
        return {"success": True, "message": "Đã cập nhật cấu hình tự động quét thành công!", "scanner": current_scanner}
    return JSONResponse(status_code=500, content={"success": False, "error": "Không thể lưu cấu hình"})


@app.post("/api/reminders/run-auto-scan")
async def api_run_auto_scan(req: RunAutoScanRequest):
    plan = ensure_plan_loaded()
    cfg = reminder_service.load_config()

    target_date = None
    if req.simulate_date and req.simulate_date.strip():
        s = req.simulate_date.strip()
        try:
            if "." in s:
                d, m, y = map(int, s.split("."))
                target_date = datetime.date(y, m, d)
            elif "-" in s:
                y, m, d = map(int, s.split("-"))
                target_date = datetime.date(y, m, d)
            elif "/" in s:
                d, m, y = map(int, s.split("/"))
                target_date = datetime.date(y, m, d)
        except Exception:
            raise HTTPException(
                status_code=400,
                detail=f"Định dạng ngày không hợp lệ: '{s}'. Vui lòng dùng định dạng dd.mm.yyyy (ví dụ 07.10.2026)"
            )

    result = reminder_service.scan_and_send_reminders(
        plan_data=plan,
        config=cfg,
        target_scan_date=target_date,
        force=req.force,
    )
    return result


@app.get("/api/reminders/sent-history")
async def api_get_sent_history():
    return {
        "success": True,
        "history": reminder_service.load_sent_log()
    }


@app.post("/api/reminders/clear-sent-history")
async def api_clear_sent_history():
    ok = reminder_service.clear_sent_log()
    return {"success": ok, "message": "Đã dọn dẹp nhật ký gửi thư!"}


import threading
import time

def start_auto_scanner_daemon():
    def _scanner_worker():
        time.sleep(5)
        while True:
            try:
                cfg = reminder_service.load_config()
                scanner_cfg = cfg.get("auto_scanner", {})
                if scanner_cfg.get("enabled", True):
                    now = datetime.datetime.now()
                    today_str = now.strftime("%Y-%m-%d")
                    last_scan_date = scanner_cfg.get("last_scan_date")

                    scan_time_str = scanner_cfg.get("scan_time", "08:30")
                    try:
                        sh, sm = map(int, scan_time_str.split(":"))
                    except Exception:
                        sh, sm = 8, 30

                    target_time = now.replace(hour=sh, minute=sm, second=0, microsecond=0)

                    # Trigger scan if now is at or past scan_time and not scanned today
                    if now >= target_time and last_scan_date != today_str:
                        plan = ensure_plan_loaded()
                        res = reminder_service.scan_and_send_reminders(
                            plan_data=plan,
                            config=cfg,
                            target_scan_date=now.date(),
                            force=False,
                        )
                        scanner_cfg["last_scan_date"] = today_str
                        scanner_cfg["last_scan_at"] = now.isoformat()
                        scanner_cfg["last_scan_summary"] = res
                        cfg["auto_scanner"] = scanner_cfg
                        reminder_service.save_config(cfg)
            except Exception:
                pass
            time.sleep(30)

    daemon_thread = threading.Thread(target=_scanner_worker, daemon=True)
    daemon_thread.start()

start_auto_scanner_daemon()


if __name__ == "__main__":
    import uvicorn
    import threading
    import webbrowser
    import time

    port = int(os.environ.get("PORT", 8000))
    host = os.environ.get("HOST", "0.0.0.0")

    def open_browser():
        time.sleep(1.5)
        try:
            webbrowser.open(f"http://localhost:{port}")
        except Exception:
            pass

    # Only open browser locally, not in cloud environments (Render, Railway, etc.)
    if not os.environ.get("RENDER") and not os.environ.get("PORT"):
        threading.Thread(target=open_browser, daemon=True).start()

    print("================================================================")
    print("🚀 Đang khởi chạy Hệ Thống Lập Kế Hoạch Kiểm Soát FECT (QA FECT)...")
    print(f"🌐 Server lắng nghe tại: http://{host}:{port}")
    print("💡 Giữ cửa sổ này hoạt động trong suốt quá trình sử dụng.")
    print("================================================================")
    uvicorn.run(app, host=host, port=port, log_level="info")
