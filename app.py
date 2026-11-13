import os
import io
import json
import zipfile
from typing import Optional, List, Dict, Any
import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from pydantic import BaseModel

from agent_engine import (
    SalesProgressAgent,
    load_from_google_sheets_url,
    parse_excel_bytes,
    AUDIT_LOG_FILE,
    DEFAULT_GEMINI_API_KEY
)
from test_runner import load_scenario_dfs, run_all_self_tests
from data_generator import DEFAULT_REPORT_DATE, generate_all_datasets

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
SAMPLE_DIR = os.path.join(BASE_DIR, "sample_data")
DOCS_DIR = os.path.join(BASE_DIR, "docs")

try:
    if not os.path.exists(os.path.join(SAMPLE_DIR, "scenario_1_normal")):
        generate_all_datasets(SAMPLE_DIR)
except Exception:
    pass

app = FastAPI(
    title="Duydudi - AI Agent Phân Tích Tiến Độ Cơ Hội & Dự Kiến Doanh Số (AI301)",
    version="1.0.0"
)

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# Lưu bộ dữ liệu gần nhất được tải lên trong phiên để Bước 4 (Human-in-the-loop) có thể duyệt lại ngay trên dữ liệu vừa upload
LAST_SESSION_DATA: Dict[str, Any] = {
    "source_type": "preset",
    "scenario_id": "scenario_1_normal",
    "opps_df": None,
    "acts_df": None,
    "orders_df": None,
    "kpi_df": None,
    "dataset_label": "Kịch bản 1: Trường hợp thông thường (Normal Case)"
}
LAST_ANALYSIS_RESULT: Optional[Dict[str, Any]] = None


class PresetAnalyzeRequest(BaseModel):
    scenario_id: str = "scenario_1_normal"
    report_date: str = DEFAULT_REPORT_DATE
    inactivity_threshold_days: int = 14
    human_overrides: Dict[str, Dict[str, Any]] = {}
    llm_provider: str = "gemini"
    api_key: str = DEFAULT_GEMINI_API_KEY
    use_last_uploaded: bool = False


class GSheetAnalyzeRequest(BaseModel):
    gsheet_url: str
    report_date: str = DEFAULT_REPORT_DATE
    inactivity_threshold_days: int = 14
    human_overrides: Dict[str, Dict[str, Any]] = {}
    llm_provider: str = "gemini"
    api_key: str = DEFAULT_GEMINI_API_KEY


@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_path = os.path.join(STATIC_DIR, "index.html")
    with open(index_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.post("/api/analyze/preset")
async def analyze_preset(req: PresetAnalyzeRequest):
    global LAST_SESSION_DATA
    try:
        if req.use_last_uploaded and LAST_SESSION_DATA.get("opps_df") is not None:
            opps_df = LAST_SESSION_DATA["opps_df"]
            acts_df = LAST_SESSION_DATA["acts_df"]
            orders_df = LAST_SESSION_DATA["orders_df"]
            kpi_df = LAST_SESSION_DATA["kpi_df"]
            label = LAST_SESSION_DATA["dataset_label"]
        else:
            scenario_map = {
                "scenario_1_normal": "Kịch bản 1: Trường hợp thông thường (45 cơ hội chuẩn)",
                "scenario_0_full300": "Bối cảnh đầy đủ Công ty của Duydudi (10 NVKD - 300 cơ hội)",
                "scenario_2_missing": "Kịch bản 2: Thiếu dữ liệu & Mâu thuẫn phức tạp",
                "scenario_3_boundary": "Kịch bản 3: Tại ranh giới quy tắc (< 5 cơ hội hợp lệ & Cut-off Date)",
            }
            folder = req.scenario_id if req.scenario_id in scenario_map else "scenario_1_normal"
            opps_df, acts_df, orders_df, kpi_df = load_scenario_dfs(folder)
            label = scenario_map.get(folder, folder)
            LAST_SESSION_DATA = {
                "source_type": "preset",
                "scenario_id": folder,
                "opps_df": opps_df,
                "acts_df": acts_df,
                "orders_df": orders_df,
                "kpi_df": kpi_df,
                "dataset_label": label
            }

        agent = SalesProgressAgent(
            report_date=req.report_date,
            inactivity_threshold_days=req.inactivity_threshold_days,
            human_overrides=req.human_overrides,
            llm_provider=req.llm_provider,
            api_key=req.api_key
        )
        result = agent.execute_pipeline(opps_df, acts_df, orders_df, kpi_df, dataset_label=label)
        global LAST_ANALYSIS_RESULT
        LAST_ANALYSIS_RESULT = result
        return JSONResponse(content=result)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


def _read_single_uploaded_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    fname = filename.lower()
    if fname.endswith(".csv"):
        for enc in ("utf-8-sig", "utf-8", "cp1258", "latin1"):
            try:
                return pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
            except Exception:
                continue
        raise ValueError(f"Không thể đọc file CSV: {filename}")
    elif fname.endswith(".xlsx") or fname.endswith(".xls"):
        return pd.read_excel(io.BytesIO(file_bytes))
    else:
        raise ValueError(f"Định dạng file không hỗ trợ: {filename}. Vui lòng tải lên .csv, .xlsx hoặc .zip")


def _process_uploaded_content(file_bytes: bytes, filename: str, tables_dict: Dict[str, pd.DataFrame], logger: Any):
    """
    Tự động giải nén ZIP hoặc đọc XLSX đa sheet hoặc nhận diện CSV theo tên/cột.
    """
    fname = filename.lower()
    
    # 1. Nếu là file nén ZIP chứa nhiều CSV/XLSX
    if fname.endswith(".zip"):
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            for item_name in zf.namelist():
                if not item_name.startswith("__MACOSX") and item_name.lower().endswith((".csv", ".xlsx", ".xls")):
                    item_bytes = zf.read(item_name)
                    _process_uploaded_content(item_bytes, item_name, tables_dict, logger)
        return

    # 2. Nếu là file Excel (.xlsx, .xls) đa sheet
    if fname.endswith((".xlsx", ".xls")):
        wb_tables = parse_excel_bytes(file_bytes, filename, logger)
        for k, v in wb_tables.items():
            if k not in tables_dict or tables_dict[k] is None:
                tables_dict[k] = v
        return

    # 3. Nếu là file CSV
    if fname.endswith(".csv"):
        df = _read_single_uploaded_file(file_bytes, filename)
        # Nhận diện theo tên file
        if any(k in fname for k in ["opp", "co_hoi", "cơ hội", "pipeline", "deals"]):
            tables_dict["opportunities"] = df
        elif any(k in fname for k in ["act", "hoat_dong", "hoạt động", "liên hệ", "contact"]):
            tables_dict["activities"] = df
        elif any(k in fname for k in ["ord", "don_hang", "đơn hàng", "orders"]):
            tables_dict["orders"] = df
        elif any(k in fname for k in ["kpi", "target", "muc_tieu", "mục tiêu", "nhân viên", "sales"]):
            tables_dict["kpi_targets"] = df
        else:
            # Nhận diện theo cột
            cols_lower = [str(c).lower().strip() for c in df.columns]
            if any(k in cols_lower for k in ["deal_value", "expected_close_date", "opportunity_id", "opp_id"]):
                tables_dict["opportunities"] = df
            elif any(k in cols_lower for k in ["activity_date", "outcome_signal", "activity_type"]):
                tables_dict["activities"] = df
            elif any(k in cols_lower for k in ["order_amount", "order_date", "order_id"]):
                tables_dict["orders"] = df
            elif any(k in cols_lower for k in ["kpi_target", "rep_target"]):
                tables_dict["kpi_targets"] = df


@app.post("/api/analyze/upload")
async def analyze_upload(
    report_date: str = Form(DEFAULT_REPORT_DATE),
    inactivity_threshold_days: int = Form(14),
    llm_provider: str = Form("gemini"),
    api_key: str = Form(DEFAULT_GEMINI_API_KEY),
    human_overrides: str = Form("{}"),
    human_overrides_json: str = Form("{}"),
    files: Optional[List[UploadFile]] = File(None),
    workbook_file: Optional[UploadFile] = File(None),
    opportunities_file: Optional[UploadFile] = File(None),
    activities_file: Optional[UploadFile] = File(None),
    orders_file: Optional[UploadFile] = File(None),
    kpi_file: Optional[UploadFile] = File(None),
):
    global LAST_SESSION_DATA
    raw_override_str = human_overrides if human_overrides != "{}" else human_overrides_json
    try:
        overrides = json.loads(raw_override_str or "{}")
    except Exception:
        overrides = {}

    agent = SalesProgressAgent(
        report_date=report_date,
        inactivity_threshold_days=inactivity_threshold_days,
        human_overrides=overrides,
        llm_provider=llm_provider,
        api_key=api_key
    )

    # Thu thập tất cả các file được gửi lên
    all_uploads: List[UploadFile] = []
    if files:
        all_uploads.extend(files)
    for single_f in [workbook_file, opportunities_file, activities_file, orders_file, kpi_file]:
        if single_f and single_f.filename:
            all_uploads.append(single_f)

    if not all_uploads:
        raise HTTPException(status_code=400, detail="Vui lòng chọn ít nhất một file .xlsx, .csv hoặc .zip để tải lên.")

    tables_dict: Dict[str, pd.DataFrame] = {}
    uploaded_names: List[str] = []

    try:
        for up_file in all_uploads:
            if not up_file.filename:
                continue
            uploaded_names.append(up_file.filename)
            content = await up_file.read()
            _process_uploaded_content(content, up_file.filename, tables_dict, agent.logger)

        # Mặc định lấy từ bộ chuẩn nếu người dùng chưa tải đủ 4 bảng
        default_opps, default_acts, default_orders, default_kpi = load_scenario_dfs("scenario_1_normal")
        opps_df = tables_dict.get("opportunities", default_opps)
        acts_df = tables_dict.get("activities", default_acts)
        orders_df = tables_dict.get("orders", default_orders)
        kpi_df = tables_dict.get("kpi_targets", default_kpi)

        label = f"Dữ liệu tải lên ({', '.join(uploaded_names)})"
        LAST_SESSION_DATA = {
            "source_type": "upload",
            "scenario_id": "custom_upload",
            "opps_df": opps_df,
            "acts_df": acts_df,
            "orders_df": orders_df,
            "kpi_df": kpi_df,
            "dataset_label": label
        }

        result = agent.execute_pipeline(opps_df, acts_df, orders_df, kpi_df, dataset_label=label)
        result["uploaded_files"] = uploaded_names
        result["detected_tables"] = list(tables_dict.keys())
        global LAST_ANALYSIS_RESULT
        LAST_ANALYSIS_RESULT = result
        return JSONResponse(content=result)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/analyze/gsheet")
async def analyze_gsheet(req: GSheetAnalyzeRequest):
    global LAST_SESSION_DATA
    try:
        agent = SalesProgressAgent(
            report_date=req.report_date,
            inactivity_threshold_days=req.inactivity_threshold_days,
            human_overrides=req.human_overrides,
            llm_provider=req.llm_provider,
            api_key=req.api_key
        )
        tables = load_from_google_sheets_url(req.gsheet_url, agent.logger)
        default_opps, default_acts, default_orders, default_kpi = load_scenario_dfs("scenario_1_normal")
        opps_df = tables.get("opportunities", default_opps)
        acts_df = tables.get("activities", default_acts)
        orders_df = tables.get("orders", default_orders)
        kpi_df = tables.get("kpi_targets", default_kpi)

        label = f"Google Sheets ({req.gsheet_url[:45]}...)"
        LAST_SESSION_DATA = {
            "source_type": "gsheet",
            "scenario_id": "custom_gsheet",
            "opps_df": opps_df,
            "acts_df": acts_df,
            "orders_df": orders_df,
            "kpi_df": kpi_df,
            "dataset_label": label
        }
        result = agent.execute_pipeline(opps_df, acts_df, orders_df, kpi_df, dataset_label=label)
        result["gsheet_url"] = req.gsheet_url
        result["detected_tables"] = list(tables.keys())
        global LAST_ANALYSIS_RESULT
        LAST_ANALYSIS_RESULT = result
        return JSONResponse(content=result)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/api/mock-gsheet")
@app.get("/api/mock-gsheet/export")
async def mock_gsheet_export(scenario_id: str = "scenario_1_normal"):
    """
    Cung cấp link Google Sheets mẫu hoạt động offline hoặc demo trực tiếp.
    """
    file_map = {
        "scenario_1_normal": ("scenario_1_normal", "BinhMinh_Scenario1_Normal.xlsx"),
        "scenario_0_full300": ("scenario_0_full300", "BinhMinh_Full300_Context.xlsx"),
        "scenario_2_missing": ("scenario_2_missing", "BinhMinh_Scenario2_MissingConflict.xlsx"),
        "scenario_3_boundary": ("scenario_3_boundary", "BinhMinh_Scenario3_Boundary.xlsx"),
    }
    folder, fname = file_map.get(scenario_id, file_map["scenario_1_normal"])
    fpath = os.path.join(SAMPLE_DIR, folder, fname)
    if os.path.exists(fpath):
        return FileResponse(
            fpath,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=fname
        )
    raise HTTPException(status_code=404, detail="File không tồn tại")


@app.get("/api/self-test")
async def api_self_test(report_date: str = DEFAULT_REPORT_DATE):
    suites = run_all_self_tests(report_date=report_date)
    return JSONResponse(content={"report_date": report_date, "suites": suites})


@app.get("/api/task1-doc")
async def api_task1_doc():
    doc_path = os.path.join(DOCS_DIR, "TASK1_BAO_CAO_THIET_KE.md")
    if os.path.exists(doc_path):
        with open(doc_path, "r", encoding="utf-8") as f:
            return JSONResponse(content={"markdown": f.read()})
    return JSONResponse(content={"markdown": "# Chưa tìm thấy tài liệu Task 1"})


@app.get("/api/download-sample/{scenario_id}")
async def download_sample(scenario_id: str, format: str = "xlsx"):
    file_map = {
        "scenario_1_normal": ("scenario_1_normal", "BinhMinh_Scenario1_Normal.xlsx", "BinhMinh_Scenario1_CSVs.zip"),
        "scenario_0_full300": ("scenario_0_full300", "BinhMinh_Full300_Context.xlsx", "BinhMinh_Full300_CSVs.zip"),
        "scenario_2_missing": ("scenario_2_missing", "BinhMinh_Scenario2_MissingConflict.xlsx", "BinhMinh_Scenario2_CSVs.zip"),
        "scenario_3_boundary": ("scenario_3_boundary", "BinhMinh_Scenario3_Boundary.xlsx", "BinhMinh_Scenario3_CSVs.zip"),
    }
    if scenario_id not in file_map:
        scenario_id = "scenario_1_normal"
    folder, xlsx_name, zip_name = file_map[scenario_id]
    
    if format.lower() in ["zip", "csv"]:
        zip_path = os.path.join(SAMPLE_DIR, folder, zip_name)
        if os.path.exists(zip_path):
            return FileResponse(
                zip_path,
                media_type="application/zip",
                filename=zip_name
            )
            
    fpath = os.path.join(SAMPLE_DIR, folder, xlsx_name)
    if not os.path.exists(fpath):
        raise HTTPException(status_code=404, detail="File mẫu không tồn tại")
    return FileResponse(
        fpath,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=xlsx_name
    )


# =============================================================================
# CÁC ENDPOINT CHO KIẾN TRÚC END-TO-END DATA ANALYTICS AI AGENT (4 GIAI ĐOẠN)
# =============================================================================

@app.post("/api/webhook/ingest")
async def webhook_ingest(
    report_date: str = Form(DEFAULT_REPORT_DATE),
    callback_url: Optional[str] = Form(None),
    partner_id: str = Form("EXTERNAL_PARTNER"),
    files: Optional[List[UploadFile]] = File(None)
):
    """
    Tầng 1: DATA INGESTION GATEWAY cho đối tác bên ngoài đẩy dữ liệu tự động.
    Gán session_id và file_id duy nhất, thẩm định cách ly lỗi (Tầng 2) và chạy Agent (Tầng 3).
    """
    import uuid
    session_id = f"SESS-{uuid.uuid4().hex[:8].upper()}"
    file_id = f"FILE-{uuid.uuid4().hex[:8].upper()}"

    agent = SalesProgressAgent(report_date=report_date)
    tables_dict: Dict[str, pd.DataFrame] = {}
    uploaded_names: List[str] = []

    if files:
        for f in files:
            if f.filename:
                uploaded_names.append(f.filename)
                content = await f.read()
                _process_uploaded_content(content, f.filename, tables_dict, agent.logger)

    default_opps, default_acts, default_orders, default_kpi = load_scenario_dfs("scenario_1_normal")
    opps_df = tables_dict.get("opportunities", default_opps)
    acts_df = tables_dict.get("activities", default_acts)
    orders_df = tables_dict.get("orders", default_orders)
    kpi_df = tables_dict.get("kpi_targets", default_kpi)

    label = f"Webhook Ingest từ {partner_id} ({session_id})"
    result = agent.execute_pipeline(opps_df, acts_df, orders_df, kpi_df, dataset_label=label)
    result["session_id"] = session_id
    result["file_id"] = file_id
    result["partner_id"] = partner_id
    result["uploaded_files"] = uploaded_names
    result["quarantine_count"] = len(result.get("anomalies", []))
    result["clean_count"] = len(result.get("opportunities_clean", []))

    global LAST_ANALYSIS_RESULT
    LAST_ANALYSIS_RESULT = result

    # Nếu có callback_url, gọi webhook trả kết quả ngược lại cho đối tác
    if callback_url:
        import requests
        try:
            requests.post(callback_url, json={"status": "SUCCESS", "session_id": session_id, "result": result}, timeout=10)
        except Exception:
            pass

    return JSONResponse(content={
        "status": "SUCCESS",
        "message": "Đã tiếp nhận, kiểm chứng dữ liệu và thực thi phân tích AI Agent thành công.",
        "session_id": session_id,
        "file_id": file_id,
        "quarantine_count": result["quarantine_count"],
        "clean_count": result["clean_count"],
        "summary": result["company_kpis"],
        "top_5_priorities": result["priority_recommendations"]["top_opportunities"],
        "executive_brief": result.get("ai_executive_brief", {}).get("content", "")[:300] + "..."
    })


@app.get("/api/export/quarantine.csv")
async def export_quarantine_csv():
    """
    Tầng 2: Xuất danh sách bản ghi lỗi / nghi vấn bị cách ly (Quarantine Zone).
    """
    global LAST_ANALYSIS_RESULT
    anomalies = []
    if LAST_ANALYSIS_RESULT and "anomalies" in LAST_ANALYSIS_RESULT:
        anomalies = LAST_ANALYSIS_RESULT["anomalies"]
    else:
        # Nếu chưa chạy, lấy từ kịch bản chuẩn
        agent = SalesProgressAgent()
        o, a, ords, k = load_scenario_dfs("scenario_1_normal")
        res = agent.execute_pipeline(o, a, ords, k)
        anomalies = res.get("anomalies", [])

    rows = []
    for a in anomalies:
        rows.append({
            "opp_id": a.get("opp_id", ""),
            "customer_name": a.get("customer_name", ""),
            "deal_value": a.get("deal_value", 0),
            "anomaly_type": a.get("anomaly_type", ""),
            "title": a.get("title", ""),
            "description": a.get("description", ""),
            "suggested_action": a.get("suggested_action", "Cần xác minh"),
            "status": "QUARANTINED"
        })

    df = pd.DataFrame(rows)
    stream = io.StringIO()
    df.to_csv(stream, index=False, encoding="utf-8-sig")
    
    return HTMLResponse(
        content=stream.getvalue(),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=data_quarantine.csv"}
    )


@app.get("/api/export/cleaned.csv")
async def export_cleaned_csv():
    """
    Tầng 2: Xuất tập dữ liệu sạch sau chuẩn hóa (Cleaned Dataset).
    """
    global LAST_ANALYSIS_RESULT
    clean_opps = []
    if LAST_ANALYSIS_RESULT and "opportunities_clean" in LAST_ANALYSIS_RESULT:
        clean_opps = LAST_ANALYSIS_RESULT["opportunities_clean"]
    else:
        agent = SalesProgressAgent()
        o, a, ords, k = load_scenario_dfs("scenario_1_normal")
        res = agent.execute_pipeline(o, a, ords, k)
        clean_opps = res.get("opportunities_clean", [])

    df = pd.DataFrame(clean_opps)
    stream = io.StringIO()
    df.to_csv(stream, index=False, encoding="utf-8-sig")

    return HTMLResponse(
        content=stream.getvalue(),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=cleaned_data.csv"}
    )


@app.get("/api/architecture")
async def get_architecture_spec():
    """
    Trả về sơ đồ Mermaid và tài liệu kiến trúc 4 tầng chuẩn hóa.
    """
    return JSONResponse(content={
        "title": "Kiến Trúc End-to-End Data Analytics AI Agent (4 Tầng)",
        "tiers": [
            {
                "tier": "1. Data Ingestion",
                "components": ["Unified File Gateway", "Ingestion Router (API / Webhook / Upload / GSheet)"],
                "role": "Tiếp nhận đa nguồn, gán session_id và file_id duy nhất"
            },
            {
                "tier": "2. Validation & Cleaning",
                "components": ["Deterministic Python Gatekeeper", "Data Profiling", "Quarantine Zone"],
                "role": "Kiểm tra schema, lọc trùng, cách ly bản ghi lỗi vào data_quarantine.csv"
            },
            {
                "tier": "3. Agent Reasoning Loop",
                "components": ["Context Builder", "LLM Planner (Google Gemini Live)", "Sandbox Interpreter"],
                "role": "Sinh mã phân tích, tự sửa lỗi runtime, tổng hợp insight & biểu đồ"
            },
            {
                "tier": "4. Output & Delivery",
                "components": ["AY Studio Dashboard", "Report & Artifact Generator", "Webhook Callback"],
                "role": "Trình bày trực tiếp trên UI, xuất file Excel/CSV, trả kết quả qua Webhook"
            }
        ]
    })

