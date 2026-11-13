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

if not os.path.exists(os.path.join(SAMPLE_DIR, "scenario_1_normal")):
    generate_all_datasets(SAMPLE_DIR)

app = FastAPI(
    title="AI301 - Đề Dự Án 09: AI Agent Phân Tích Tiến Độ Cơ Hội & Dự Kiến Doanh Số (Công ty Bình Minh)",
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
                "scenario_0_full300": "Bối cảnh đầy đủ Công ty Bình Minh (10 NVKD - 300 cơ hội)",
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
        raise ValueError(f"Định dạng file không hỗ trợ: {filename}. Vui lòng tải lên .csv hoặc .xlsx")


@app.post("/api/analyze/upload")
async def analyze_upload(
    report_date: str = Form(DEFAULT_REPORT_DATE),
    inactivity_threshold_days: int = Form(14),
    llm_provider: str = Form("gemini"),
    api_key: str = Form(DEFAULT_GEMINI_API_KEY),
    human_overrides_json: str = Form("{}"),
    workbook_file: Optional[UploadFile] = File(None),
    opportunities_file: Optional[UploadFile] = File(None),
    activities_file: Optional[UploadFile] = File(None),
    orders_file: Optional[UploadFile] = File(None),
    kpi_file: Optional[UploadFile] = File(None),
):
    global LAST_SESSION_DATA
    try:
        overrides = json.loads(human_overrides_json or "{}")
    except Exception:
        overrides = {}

    agent = SalesProgressAgent(
        report_date=report_date,
        inactivity_threshold_days=inactivity_threshold_days,
        human_overrides=overrides,
        llm_provider=llm_provider,
        api_key=api_key
    )

    # Mặc định lấy từ bộ chuẩn nếu người dùng chỉ upload 1 vài bảng riêng lẻ
    default_opps, default_acts, default_orders, default_kpi = load_scenario_dfs("scenario_1_normal")
    opps_df, acts_df, orders_df, kpi_df = default_opps, default_acts, default_orders, default_kpi
    uploaded_names = []

    try:
        if workbook_file and workbook_file.filename:
            wb_bytes = await workbook_file.read()
            uploaded_names.append(workbook_file.filename)
            if workbook_file.filename.lower().endswith((".xlsx", ".xls")):
                tables = parse_excel_bytes(wb_bytes, workbook_file.filename, agent.logger)
                if "opportunities" in tables:
                    opps_df = tables["opportunities"]
                if "activities" in tables:
                    acts_df = tables["activities"]
                if "orders" in tables:
                    orders_df = tables["orders"]
                if "kpi_targets" in tables:
                    kpi_df = tables["kpi_targets"]
            else:
                opps_df = _read_single_uploaded_file(wb_bytes, workbook_file.filename)

        if opportunities_file and opportunities_file.filename:
            b = await opportunities_file.read()
            opps_df = _read_single_uploaded_file(b, opportunities_file.filename)
            uploaded_names.append(opportunities_file.filename)

        if activities_file and activities_file.filename:
            b = await activities_file.read()
            acts_df = _read_single_uploaded_file(b, activities_file.filename)
            uploaded_names.append(activities_file.filename)

        if orders_file and orders_file.filename:
            b = await orders_file.read()
            orders_df = _read_single_uploaded_file(b, orders_file.filename)
            uploaded_names.append(orders_file.filename)

        if kpi_file and kpi_file.filename:
            b = await kpi_file.read()
            kpi_df = _read_single_uploaded_file(b, kpi_file.filename)
            uploaded_names.append(kpi_file.filename)

        label = f"Dữ liệu tải lên ({', '.join(uploaded_names)})" if uploaded_names else "Dữ liệu tải lên"
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
        return JSONResponse(content=result)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


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
async def download_sample_excel(scenario_id: str):
    file_map = {
        "scenario_1_normal": ("scenario_1_normal", "BinhMinh_Scenario1_Normal.xlsx"),
        "scenario_0_full300": ("scenario_0_full300", "BinhMinh_Full300_Context.xlsx"),
        "scenario_2_missing": ("scenario_2_missing", "BinhMinh_Scenario2_MissingConflict.xlsx"),
        "scenario_3_boundary": ("scenario_3_boundary", "BinhMinh_Scenario3_Boundary.xlsx"),
    }
    folder, fname = file_map.get(scenario_id, file_map["scenario_1_normal"])
    fpath = os.path.join(SAMPLE_DIR, folder, fname)
    if not os.path.exists(fpath):
        raise HTTPException(status_code=404, detail="File mẫu không tồn tại")
    return FileResponse(
        fpath,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=fname
    )

