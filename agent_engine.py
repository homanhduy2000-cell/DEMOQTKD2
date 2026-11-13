import os
import re
import io
import json
import uuid
from datetime import datetime
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd
import requests

from data_generator import STAGE_RULES, DEFAULT_REPORT_DATE

DEFAULT_GEMINI_API_KEY = "AQ.Ab8RN6LqubD2sz-lzVEr8Df6PSbqOVB10Ubndkmpk5k5kxuqIg"
GEMINI_MODELS = ["gemini-3.5-flash", "gemini-3-flash-preview", "gemini-3.8-flash", "gemini-flash-latest"]
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
os.makedirs(LOG_DIR, exist_ok=True)
AUDIT_LOG_FILE = os.path.join(LOG_DIR, "agent_audit_log.jsonl")


class AuditLogger:
    """
    Hệ thống lưu nhật ký truy vết (Audit Trail) chứng minh AI Agent:
    - Đọc dữ liệu thực tế
    - Gọi công cụ (Tools) & Kỹ năng nghiệp vụ (Skills)
    - Áp dụng bộ quy tắc (Business Rules)
    - Ghi nhận quyết định của Con người (Human-in-the-loop)
    """
    def __init__(self, run_id: Optional[str] = None):
        self.run_id = run_id or f"RUN-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{str(uuid.uuid4())[:6].upper()}"
        self.entries: List[Dict[str, Any]] = []

    def log(self, step: str, skill_or_tool: str, action: str, rule_applied: str, details: Dict[str, Any], status: str = "SUCCESS"):
        entry = {
            "run_id": self.run_id,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3],
            "step": step,
            "skill_or_tool": skill_or_tool,
            "action": action,
            "rule_applied": rule_applied,
            "details": details,
            "status": status,
        }
        self.entries.append(entry)
        try:
            with open(AUDIT_LOG_FILE, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        except Exception:
            pass
        return entry


def normalize_text(val: Any) -> str:
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip()


def parse_date_safe(val: Any) -> Optional[datetime]:
    s = normalize_text(val)
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s[:10], fmt)
        except ValueError:
            continue
    try:
        dt = pd.to_datetime(s, errors="coerce")
        if pd.notna(dt):
            return dt.to_pydatetime().replace(tzinfo=None)
    except Exception:
        pass
    return None


def format_date_iso(dt: Optional[datetime]) -> str:
    if dt is None:
        return ""
    return dt.strftime("%Y-%m-%d")


def load_from_google_sheets_url(url: str, logger: AuditLogger) -> Dict[str, pd.DataFrame]:
    """
    Tải dữ liệu từ link Google Sheets công khai (Hỗ trợ cả file nhiều sheet XLSX, CSV hoặc link xuất bản Web).
    Hỗ trợ cả chế độ Demo / Mock Offline nếu mạng bị chặn.
    """
    clean_url = (url or "").strip()
    
    # 1. Hỗ trợ chế độ Demo / Mock trực tiếp
    if clean_url.lower() in ["demo", "mock", "sample", "test"] or "mock-gsheet" in clean_url.lower() or "localhost" in clean_url.lower():
        sample_xlsx = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_data", "scenario_1_normal", "BinhMinh_Scenario1_Normal.xlsx")
        if os.path.exists(sample_xlsx):
            logger.log(
                step="BƯỚC 1: Thu thập & Hợp nhất",
                skill_or_tool="Tool: GoogleSheetsConnector (Demo/Mock Mode)",
                action="Nạp Google Sheets từ bộ dữ liệu chuẩn mẫu của Công ty Bình Minh",
                rule_applied="RULE-INGEST-GSHEET-DEMO: Nạp trực tiếp bộ 4 bảng chuẩn hóa cho buổi thuyết trình",
                details={"source": "Mock Google Sheets Template", "path": sample_xlsx}
            )
            with open(sample_xlsx, "rb") as f:
                return parse_excel_bytes(f.read(), "GoogleSheets_BinhMinh_Template.xlsx", logger)

    # 2. Nhận diện link Google Sheets xuất bản Web (d/e/...)
    match_pub = re.search(r"/spreadsheets/d/e/([a-zA-Z0-9-_]+)", clean_url)
    if match_pub:
        pub_id = match_pub.group(1)
        export_url = f"https://docs.google.com/spreadsheets/d/e/{pub_id}/pub?output=xlsx"
    else:
        # Link chia sẻ Google Sheets tiêu chuẩn (d/<ID>/...)
        match_std = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", clean_url)
        if not match_std:
            raise ValueError(
                "Link Google Sheets không hợp lệ. Vui lòng cung cấp link có dạng:\n"
                "https://docs.google.com/spreadsheets/d/<ID>/edit\n"
                "hoặc bấm nút 'Dùng Link Mẫu Sẵn' để trải nghiệm ngay."
            )
        sheet_id = match_std.group(1)
        export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

    logger.log(
        step="BƯỚC 1: Thu thập & Hợp nhất",
        skill_or_tool="Tool: GoogleSheetsConnector",
        action="Tải workbook từ Google Sheets URL",
        rule_applied="RULE-INGEST-GSHEET: Tự động chuyển đổi link chia sẻ sang luồng export XLSX để đọc đa bảng",
        details={"original_url": clean_url, "export_url": export_url}
    )

    try:
        resp = requests.get(export_url, timeout=25, headers={"User-Agent": "Mozilla/5.0"})
    except Exception as ex:
        raise ValueError(f"Không thể kết nối tới Google Sheets: {str(ex)}. Hãy kiểm tra kết nối mạng hoặc dùng file Excel tải lên.")

    if resp.status_code != 200:
        # Thử phương án dự phòng xuất dạng CSV nếu là sheet đơn
        if "sheet_id" in locals():
            csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
            try:
                resp_csv = requests.get(csv_url, timeout=15)
                if resp_csv.status_code == 200:
                    df = pd.read_csv(io.BytesIO(resp_csv.content))
                    return {"opportunities": df}
            except Exception:
                pass
        raise ValueError(
            f"Không thể tải Google Sheets (Mã HTTP {resp.status_code}).\n"
            "Hãy đảm bảo trang tính đã được bật quyền: 'Bất kỳ ai có đường liên kết đều có thể xem' (Anyone with the link can view)."
        )

    # Kiểm tra nội dung trả về là XLSX hay CSV
    content = resp.content
    if content.startswith(b"PK\x03\x04"):
        return parse_excel_bytes(content, "GoogleSheets_Export.xlsx", logger)
    else:
        # Thử parse dưới dạng CSV
        for enc in ("utf-8-sig", "utf-8", "cp1258", "latin1"):
            try:
                df = pd.read_csv(io.BytesIO(content), encoding=enc)
                return {"opportunities": df}
            except Exception:
                continue
        return parse_excel_bytes(content, "GoogleSheets_Export.xlsx", logger)


def parse_excel_bytes(content: bytes, filename: str, logger: AuditLogger) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(content))
    sheet_names = xls.sheet_names
    result = {}
    
    # Quét theo tên sheet
    for s in sheet_names:
        s_lower = s.lower().strip()
        df = pd.read_excel(xls, sheet_name=s)
        if any(k in s_lower for k in ["opp", "co_hoi", "cơ hội", "pipeline", "deals"]):
            result["opportunities"] = df
        elif any(k in s_lower for k in ["act", "hoat_dong", "hoạt động", "liên hệ", "contact", "history"]):
            result["activities"] = df
        elif any(k in s_lower for k in ["ord", "don_hang", "đơn hàng", "orders"]):
            result["orders"] = df
        elif any(k in s_lower for k in ["kpi", "target", "muc_tieu", "mục tiêu", "nhân viên", "sales"]):
            result["kpi_targets"] = df

    # Nếu vẫn còn bảng chưa map, quét nội dung cột của các sheet còn lại
    for s in sheet_names:
        df = pd.read_excel(xls, sheet_name=s)
        cols_lower = [str(c).lower().strip() for c in df.columns]
        
        if "opportunities" not in result and any(k in cols_lower for k in ["deal_value", "expected_close_date", "opportunity_id", "opp_id", "gia_tri"]):
            result["opportunities"] = df
        elif "activities" not in result and any(k in cols_lower for k in ["activity_date", "outcome_signal", "activity_type", "ngay_lien_he"]):
            result["activities"] = df
        elif "orders" not in result and any(k in cols_lower for k in ["order_amount", "order_date", "order_id", "gia_tri_don"]):
            result["orders"] = df
        elif "kpi_targets" not in result and any(k in cols_lower for k in ["kpi_target", "rep_target", "muc_tieu_kpi"]):
            result["kpi_targets"] = df

    logger.log(
        step="BƯỚC 1: Thu thập & Hợp nhất",
        skill_or_tool="Tool: ExcelWorkbookParser",
        action=f"Đọc file Excel '{filename}'",
        rule_applied="RULE-INGEST-XLSX: Tự động nhận diện các sheet Opportunities, Activities, Orders, KPI_Targets",
        details={"filename": filename, "sheets_found": sheet_names, "mapped_tables": list(result.keys())}
    )
    return result


class SalesProgressAgent:
    """
    AI Agent phục vụ Đề dự án 09 - Công ty Bình Minh:
    Phối hợp 3 Kỹ năng nghiệp vụ (Skills) + Nhật ký truy vết (Audit Log) + Human-in-the-loop.
    """

    def __init__(
        self,
        report_date: str = DEFAULT_REPORT_DATE,
        inactivity_threshold_days: int = 14,
        stage_weights: Optional[Dict[str, float]] = None,
        human_overrides: Optional[Dict[str, Dict[str, Any]]] = None,
        llm_provider: str = "builtin",
        api_key: str = ""
    ):
        self.report_date_str = report_date or DEFAULT_REPORT_DATE
        self.report_date_dt = parse_date_safe(self.report_date_str) or datetime.strptime(DEFAULT_REPORT_DATE, "%Y-%m-%d")
        self.inactivity_threshold_days = int(inactivity_threshold_days)
        self.stage_weights = stage_weights or {k: v["weight"] for k, v in STAGE_RULES.items()}
        self.human_overrides = human_overrides or {}
        self.api_key = (api_key or "").strip() or DEFAULT_GEMINI_API_KEY
        self.llm_provider = llm_provider if llm_provider in ["gemini", "openai", "builtin"] else "gemini"
        if not llm_provider or llm_provider == "builtin":
            # Nếu có API key sẵn thì ưu tiên dùng Gemini API
            self.llm_provider = "gemini" if self.api_key else "builtin"
        self.logger = AuditLogger()

    def execute_pipeline(
        self,
        opps_df: pd.DataFrame,
        acts_df: pd.DataFrame,
        orders_df: pd.DataFrame,
        kpi_df: pd.DataFrame,
        dataset_label: str = "Custom Dataset"
    ) -> Dict[str, Any]:
        """
        Thực thi toàn bộ quy trình 5 bước (3 Kỹ năng AI Agent + 2 điểm kiểm soát Con người).
        """
        self.logger.log(
            step="BƯỚC 1: Thu thập & Hợp nhất",
            skill_or_tool="Tool: DataIngestionValidator",
            action=f"Khởi chạy quy trình phân tích trên bộ dữ liệu '{dataset_label}'",
            rule_applied=f"RULE-CUTOFF: Ngày đối chiếu thống nhất (Report Date) = {self.report_date_str}",
            details={
                "raw_opportunities_rows": len(opps_df),
                "raw_activities_rows": len(acts_df),
                "raw_orders_rows": len(orders_df),
                "sales_reps_rows": len(kpi_df),
                "report_date": self.report_date_str,
                "inactivity_threshold_days": self.inactivity_threshold_days,
                "active_human_overrides": len(self.human_overrides)
            }
        )

        # =========================================================================
        # SKILL 1 (BƯỚC 2): CHUẨN HÓA, HỢP NHẤT CƠ HỘI & TÍNH TOÁN GIÁ TRỊ DỰ KIẾN
        # =========================================================================
        consolidated_opps, duplicate_logs, cut_off_logs, valid_orders_df, valid_acts_df = self._skill_1_consolidate_and_valuate(
            opps_df, acts_df, orders_df
        )

        # =========================================================================
        # SKILL 2 (BƯỚC 3 & BƯỚC 4): PHÁT HIỆN RỦI RO, QUÁ HẠN, THIẾU DỮ LIỆU & MÂU THUẪN
        # =========================================================================
        anomalies_list, enriched_opps = self._skill_2_detect_risks_and_conflicts(
            consolidated_opps, valid_acts_df, valid_orders_df, duplicate_logs
        )

        # =========================================================================
        # SKILL 3 (BƯỚC 5): PHÂN TÍCH MỨC HOÀN THÀNH KPI & ĐỀ XUẤT 5 CƠ HỘI ƯU TIÊN
        # =========================================================================
        rep_summary_list, company_kpis, priority_result = self._skill_3_evaluate_kpi_and_prioritize(
            enriched_opps, valid_orders_df, kpi_df
        )

        # Gọi LLM (nếu có API Key) hoặc AI Reasoning Engine để tạo nhận xét chuyên sâu
        ai_executive_brief = self._generate_executive_synthesis(
            company_kpis, rep_summary_list, anomalies_list, priority_result
        )

        return {
            "run_id": self.logger.run_id,
            "report_date": self.report_date_str,
            "dataset_label": dataset_label,
            "rules_config": {
                "report_date": self.report_date_str,
                "inactivity_threshold_days": self.inactivity_threshold_days,
                "stage_weights": self.stage_weights,
                "priority_formula": "35% Giá trị dự kiến (Weighted Value) + 25% Độ chín giai đoạn (Stage) + 20% Độ khẩn cấp ngày chốt (Urgency) + 20% Áp lực thiếu KPI của NVKD (KPI Gap)"
            },
            "company_kpis": company_kpis,
            "rep_summary": rep_summary_list,
            "opportunities_clean": enriched_opps,
            "anomalies": anomalies_list,
            "duplicates_merged": duplicate_logs,
            "cutoff_filtered": cut_off_logs,
            "priority_recommendations": priority_result,
            "ai_executive_brief": ai_executive_brief,
            "audit_logs": self.logger.entries,
        }

    # -------------------------------------------------------------------------
    # SKILL 1: HỢP NHẤT CƠ HỘI, LỌC NGÀY ĐỐI CHIẾU, TÁCH DOANH SỐ & TÍNH TRỌNG SỐ
    # -------------------------------------------------------------------------
    def _skill_1_consolidate_and_valuate(
        self,
        opps_df: pd.DataFrame,
        acts_df: pd.DataFrame,
        orders_df: pd.DataFrame
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]], pd.DataFrame, pd.DataFrame]:
        cut_off_logs = []

        # 1. Lọc Orders theo ngày đối chiếu (order_date <= report_date)
        valid_orders_rows = []
        for _, row in orders_df.iterrows():
            o_id = normalize_text(row.get("order_id"))
            o_date_dt = parse_date_safe(row.get("order_date"))
            o_status = normalize_text(row.get("order_status", "Approved"))
            if o_date_dt and o_date_dt > self.report_date_dt:
                cut_off_logs.append({
                    "record_type": "ORDER",
                    "record_id": o_id,
                    "opp_id": normalize_text(row.get("opp_id")),
                    "record_date": format_date_iso(o_date_dt),
                    "reason": f"Đơn hàng phát sinh ngày {format_date_iso(o_date_dt)} (sau ngày đối chiếu {self.report_date_str}) -> Chưa tính vào doanh số kỳ này."
                })
                continue
            if o_status.lower() in ["approved", "completed", "đã duyệt", "thành công", ""]:
                valid_orders_rows.append(row.to_dict())
        valid_orders_df = pd.DataFrame(valid_orders_rows) if valid_orders_rows else pd.DataFrame(columns=orders_df.columns)

        # 2. Lọc Activities theo ngày đối chiếu (activity_date <= report_date)
        valid_acts_rows = []
        for _, row in acts_df.iterrows():
            a_id = normalize_text(row.get("activity_id"))
            a_date_dt = parse_date_safe(row.get("activity_date"))
            if a_date_dt and a_date_dt > self.report_date_dt:
                cut_off_logs.append({
                    "record_type": "ACTIVITY",
                    "record_id": a_id,
                    "opp_id": normalize_text(row.get("opp_id")),
                    "record_date": format_date_iso(a_date_dt),
                    "reason": f"Hoạt động liên hệ ngày {format_date_iso(a_date_dt)} (sau ngày đối chiếu {self.report_date_str}) -> Tạm thời bỏ qua tại mốc chốt sổ."
                })
                continue
            valid_acts_rows.append(row.to_dict())
        valid_acts_df = pd.DataFrame(valid_acts_rows) if valid_acts_rows else pd.DataFrame(columns=acts_df.columns)

        if cut_off_logs:
            self.logger.log(
                step="BƯỚC 2: Chuẩn hóa & Tính toán",
                skill_or_tool="Skill 1: DataConsolidationAndValuationSkill",
                action="Lọc dữ liệu theo ngày đối chiếu thống nhất (Cut-off Filter)",
                rule_applied=f"RULE-CUTOFF-FILTER: Chỉ ghi nhận đơn hàng và hoạt động có ngày <= {self.report_date_str}",
                details={"excluded_future_records_count": len(cut_off_logs), "excluded_items": cut_off_logs}
            )

        # 3. Hợp nhất cơ hội trùng lặp (Deduplication)
        # Nhóm theo opp_id chuẩn HOẶC khóa ngữ nghĩa (customer_name + project_name)
        raw_records = opps_df.to_dict(orient="records")
        key_map: Dict[str, List[Dict[str, Any]]] = {}
        id_to_semantic_key: Dict[str, str] = {}

        for r in raw_records:
            opp_id = normalize_text(r.get("opp_id"))
            cust = normalize_text(r.get("customer_name")).lower()
            proj = normalize_text(r.get("project_name")).lower()
            sem_key = f"{cust}||{proj}"
            if opp_id in id_to_semantic_key:
                group_key = id_to_semantic_key[opp_id]
            elif sem_key in key_map and cust and proj:
                group_key = sem_key
            else:
                group_key = sem_key if (cust and proj) else opp_id
            id_to_semantic_key[opp_id] = group_key
            key_map.setdefault(group_key, []).append(r)

        consolidated_opps = []
        duplicate_logs = []

        for group_key, group_items in key_map.items():
            if len(group_items) == 1:
                chosen = dict(group_items[0])
                chosen["is_merged_duplicate"] = False
                chosen["merged_from_files"] = [normalize_text(chosen.get("source_file", "Primary"))]
                chosen["duplicate_conflict_detail"] = ""
                consolidated_opps.append(chosen)
            else:
                # Sắp xếp theo updated_date mới nhất (nhưng <= report_date)
                def sort_key(item):
                    dt = parse_date_safe(item.get("updated_date"))
                    if dt and dt <= self.report_date_dt:
                        return dt
                    return datetime(2000, 1, 1)

                sorted_items = sorted(group_items, key=sort_key, reverse=True)
                winner = dict(sorted_items[0])
                all_ids = list(dict.fromkeys(normalize_text(x.get("opp_id")) for x in sorted_items))
                all_files = list(dict.fromkeys(normalize_text(x.get("source_file", "Unknown")) for x in sorted_items))
                all_reps = list(dict.fromkeys(normalize_text(x.get("rep_id")) for x in sorted_items))
                all_vals = [float(x.get("deal_value") or 0) for x in sorted_items if pd.notna(x.get("deal_value"))]

                # Giữ mã opp_id chuẩn (ngắn nhất / gốc)
                canonical_id = sorted(all_ids, key=lambda x: len(x))[0]
                winner["opp_id"] = canonical_id
                winner["is_merged_duplicate"] = True
                winner["merged_from_files"] = all_files

                conflict_notes = []
                if len(all_reps) > 1:
                    conflict_notes.append(f"Lệch NVKD phụ trách giữa các tệp ({', '.join(all_reps)})")
                if len(all_vals) > 1 and max(all_vals) > 0:
                    diff_ratio = (max(all_vals) - min(all_vals)) / max(all_vals)
                    if diff_ratio >= 0.20:
                        conflict_notes.append(
                            f"Chênh lệch giá trị lớn ({min(all_vals):,.0f} đ vs {max(all_vals):,.0f} đ - lệch {diff_ratio*100:.1f}%)"
                        )

                winner["duplicate_conflict_detail"] = "; ".join(conflict_notes)
                consolidated_opps.append(winner)

                dup_entry = {
                    "canonical_opp_id": canonical_id,
                    "customer_name": normalize_text(winner.get("customer_name")),
                    "project_name": normalize_text(winner.get("project_name")),
                    "records_merged_count": len(sorted_items),
                    "all_opp_ids": all_ids,
                    "source_files": all_files,
                    "kept_updated_date": normalize_text(winner.get("updated_date")),
                    "kept_stage": normalize_text(winner.get("stage")),
                    "kept_value": float(winner.get("deal_value") or 0),
                    "conflict_detected": bool(conflict_notes),
                    "conflict_detail": "; ".join(conflict_notes) if conflict_notes else "Hợp nhất tự động theo bản ghi cập nhật mới nhất trước ngày chốt."
                }
                duplicate_logs.append(dup_entry)

        self.logger.log(
            step="BƯỚC 2: Chuẩn hóa & Tính toán",
            skill_or_tool="Skill 1: DataConsolidationAndValuationSkill",
            action="Đối chiếu và hợp nhất cơ hội trùng lặp từ nhiều tệp",
            rule_applied="RULE-DEDUP: Gộp cơ hội trùng mã hoặc trùng (Khách hàng + Dự án), ưu tiên bản cập nhật mới nhất <= Ngày báo cáo",
            details={
                "raw_count": len(raw_records),
                "consolidated_unique_count": len(consolidated_opps),
                "duplicates_merged_groups": len(duplicate_logs),
                "duplicate_summary": duplicate_logs
            }
        )

        return consolidated_opps, duplicate_logs, cut_off_logs, valid_orders_df, valid_acts_df

    # -------------------------------------------------------------------------
    # SKILL 2: PHÁT HIỆN RỦI RO, QUÁ HẠN, THIẾU THÔNG TIN & MÂU THUẪN DỮ LIỆU
    # -------------------------------------------------------------------------
    def _skill_2_detect_risks_and_conflicts(
        self,
        consolidated_opps: List[Dict[str, Any]],
        valid_acts_df: pd.DataFrame,
        valid_orders_df: pd.DataFrame,
        duplicate_logs: List[Dict[str, Any]]
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        # Map orders by opp_id and customer_name
        orders_by_opp: Dict[str, Dict[str, Any]] = {}
        if not valid_orders_df.empty:
            for _, orow in valid_orders_df.iterrows():
                oid = normalize_text(orow.get("opp_id"))
                if oid:
                    orders_by_opp[oid] = orow.to_dict()

        # Map activities by opp_id
        acts_by_opp: Dict[str, List[Dict[str, Any]]] = {}
        if not valid_acts_df.empty:
            for _, arow in valid_acts_df.iterrows():
                oid = normalize_text(arow.get("opp_id"))
                # Chuẩn hóa nếu mã opp_id có hậu tố -DUP
                base_oid = oid.replace("-DUP", "")
                acts_by_opp.setdefault(base_oid, []).append(arow.to_dict())

        anomalies_list: List[Dict[str, Any]] = []
        enriched_opps: List[Dict[str, Any]] = []

        lost_keywords = ["mất khách", "đối thủ", "hủy dự án", "từ chối", "ngừng triển khai", "không đầu tư"]

        for opp in consolidated_opps:
            opp_id = normalize_text(opp.get("opp_id"))
            cust_name = normalize_text(opp.get("customer_name"))
            proj_name = normalize_text(opp.get("project_name"))
            rep_id = normalize_text(opp.get("rep_id"))
            raw_stage = normalize_text(opp.get("stage"))
            try:
                deal_value = float(opp.get("deal_value") or 0)
            except ValueError:
                deal_value = 0.0
            try:
                sales_custom_prob = float(opp.get("sales_custom_prob") or 0)
            except ValueError:
                sales_custom_prob = 0.0

            exp_close_str = normalize_text(opp.get("expected_close_date"))
            exp_close_dt = parse_date_safe(exp_close_str)
            updated_str = normalize_text(opp.get("updated_date"))
            updated_dt = parse_date_safe(updated_str)

            decision_maker = normalize_text(opp.get("decision_maker"))
            budget_confirmed = normalize_text(opp.get("budget_confirmed"))
            next_step = normalize_text(opp.get("next_step"))

            # Tìm hoạt động liên hệ gần nhất (<= report_date)
            opp_acts = acts_by_opp.get(opp_id, [])
            latest_act_dt = updated_dt
            latest_act_note = ""
            latest_act_signal = ""
            if opp_acts:
                opp_acts_sorted = sorted(
                    opp_acts,
                    key=lambda x: parse_date_safe(x.get("activity_date")) or datetime(2000, 1, 1),
                    reverse=True
                )
                top_act = opp_acts_sorted[0]
                act_dt = parse_date_safe(top_act.get("activity_date"))
                if act_dt:
                    latest_act_dt = act_dt
                latest_act_note = normalize_text(top_act.get("notes"))
                latest_act_signal = normalize_text(top_act.get("outcome_signal")).upper()

            days_since_activity = (self.report_date_dt - latest_act_dt).days if latest_act_dt else 999
            days_to_close = (exp_close_dt - self.report_date_dt).days if exp_close_dt else None

            # Kiểm tra đơn hàng đã ghi nhận
            matched_order = orders_by_opp.get(opp_id)
            has_recorded_order = matched_order is not None
            recorded_order_amount = float(matched_order.get("order_amount") or 0) if matched_order else 0.0
            recorded_order_id = normalize_text(matched_order.get("order_id")) if matched_order else ""

            # Kiểm tra tín hiệu mất khách từ Activities
            note_lower = latest_act_note.lower()
            has_lost_signal = (latest_act_signal == "LOST") or any(kw in note_lower for kw in lost_keywords)

            # Xác định trạng thái chuẩn của AI tại ngày báo cáo (trước khi Con người Override)
            opp_anomalies: List[Dict[str, str]] = []
            effective_stage = raw_stage if raw_stage in STAGE_RULES else "Qualification"
            status_resolution_reason = "Giữ nguyên trạng thái theo hồ sơ"

            # 1) Phát hiện mâu thuẫn với Đơn hàng đã ghi nhận (Orders)
            if has_recorded_order:
                if raw_stage != "Closed Won":
                    opp_anomalies.append({
                        "type": "CONFLICT_ORDER_STATUS",
                        "severity": "CRITICAL",
                        "title": "Đã có Đơn hàng ghi nhận nhưng Cơ hội chưa chuyển Closed Won",
                        "description": (
                            f"Đã phát sinh đơn hàng {recorded_order_id} ({recorded_order_amount:,.0f} VNĐ) ngày "
                            f"{normalize_text(matched_order.get('order_date'))}, nhưng hồ sơ cơ hội đang để '{raw_stage}'. "
                            f"Nếu không tách ra sẽ bị tính trùng vào cả Doanh số thực tế và Cơ hội mở."
                        ),
                        "ai_suggested_action": "Chuyển trạng thái sang 'Closed Won', ghi nhận vào Doanh số thực tế và loại khỏi Cơ hội mở."
                    })
                effective_stage = "Closed Won"
                status_resolution_reason = f"Tự động tách khỏi cơ hội mở do đã có đơn hàng {recorded_order_id}"

            # 2) Phát hiện mâu thuẫn đã mất khách trong Activities nhưng trạng thái vẫn mở
            elif has_lost_signal and raw_stage not in ["Closed Lost", "Closed Won"]:
                opp_anomalies.append({
                    "type": "CONFLICT_ZOMBIE_LOST",
                    "severity": "CRITICAL",
                    "title": "Mâu thuẫn trạng thái: Khách đã từ chối/mất khách nhưng cơ hội vẫn để Mở",
                    "description": (
                        f"Trạng thái đang để '{raw_stage}' (Tỷ lệ Sales nhập {sales_custom_prob*100:.0f}%), "
                        f"nhưng lịch sử liên hệ ngày {format_date_iso(latest_act_dt)} ghi: \"{latest_act_note}\"."
                    ),
                    "ai_suggested_action": "Chuyển trạng thái sang 'Closed Lost' (Trọng số 0%) để không làm ảo dự báo doanh số."
                })
                effective_stage = "Closed Lost"
                status_resolution_reason = "Tự động chuyển sang Closed Lost do lịch sử liên hệ xác nhận đã mất khách"

            # 3) Phát hiện thiếu giai đoạn hợp lệ
            if raw_stage not in STAGE_RULES:
                opp_anomalies.append({
                    "type": "MISSING_STAGE",
                    "severity": "HIGH",
                    "title": "Thiếu hoặc sai định dạng Giai đoạn (Stage)",
                    "description": f"Giá trị cột stage hiện tại là '{raw_stage or '(Trống)'}', không khớp danh mục quy tắc chuẩn.",
                    "ai_suggested_action": "Yêu cầu NVKD xác nhận lại giai đoạn thực tế (Tạm áp mức Qualification 25%)."
                })

            # 4) Phát hiện mâu thuẫn khi gộp bản ghi trùng
            if opp.get("duplicate_conflict_detail"):
                opp_anomalies.append({
                    "type": "CONFLICT_DUPLICATE_MERGE",
                    "severity": "HIGH",
                    "title": "Bản ghi trùng lặp giữa các tệp có thông tin mâu thuẫn",
                    "description": f"Hợp nhất từ {len(opp.get('merged_from_files', []))} tệp ({', '.join(opp.get('merged_from_files', []))}): {opp.get('duplicate_conflict_detail')}.",
                    "ai_suggested_action": "Trưởng phòng KD xác nhận NVKD phụ trách chính thức và giá trị thương vụ chuẩn."
                })

            # Các kiểm tra tiếp theo chỉ áp dụng cho cơ hội đang mở
            is_currently_open = effective_stage not in ["Closed Won", "Closed Lost"]

            # 5) Kiểm tra Quá hạn ngày dự kiến chốt (expected_close_date < report_date)
            is_overdue_close = False
            if is_currently_open:
                if exp_close_dt is None:
                    opp_anomalies.append({
                        "type": "MISSING_CLOSE_DATE",
                        "severity": "HIGH",
                        "title": "Thiếu Ngày dự kiến chốt (Expected Close Date)",
                        "description": "Cơ hội đang mở nhưng bỏ trống ngày dự kiến chốt, không thể lập kế hoạch dòng tiền.",
                        "ai_suggested_action": "Yêu cầu NVKD cập nhật ngày dự kiến chốt có xác nhận của khách hàng."
                    })
                elif exp_close_dt < self.report_date_dt:
                    is_overdue_close = True
                    overdue_days = (self.report_date_dt - exp_close_dt).days
                    opp_anomalies.append({
                        "type": "OVERDUE_CLOSE_DATE",
                        "severity": "HIGH",
                        "title": f"Quá hạn ngày dự kiến chốt ({overdue_days} ngày)",
                        "description": (
                            f"Ngày dự kiến chốt ({exp_close_str}) đã qua so với ngày báo cáo ({self.report_date_str}) "
                            f"{overdue_days} ngày nhưng cơ hội vẫn ở giai đoạn '{effective_stage}'."
                        ),
                        "ai_suggested_action": "Kiểm tra thực tế với khách hàng: gia hạn ngày chốt mới có căn cứ hoặc đóng cơ hội (Closed Lost)."
                    })

            # 6) Kiểm tra không có hoạt động tương tác mới (Inactivity > threshold)
            is_stale_inactive = False
            if is_currently_open and days_since_activity > self.inactivity_threshold_days:
                is_stale_inactive = True
                opp_anomalies.append({
                    "type": "INACTIVE_STALE",
                    "severity": "MEDIUM" if not is_overdue_close else "HIGH",
                    "title": f"Đóng băng tương tác ({days_since_activity} ngày không có hoạt động mới)",
                    "description": (
                        f"Lần tương tác gần nhất là {format_date_iso(latest_act_dt)} ({days_since_activity} ngày trước, "
                        f"vượt ngưỡng quy định {self.inactivity_threshold_days} ngày)."
                    ),
                    "ai_suggested_action": "Yêu cầu NVKD liên hệ lại ngay trong 48h hoặc chuyển sang danh sách cơ hội treo."
                })

            # 7) Kiểm tra thiếu thông tin xác nhận (Người quyết định, Ngân sách, Bước tiếp theo, Giá trị <= 0)
            missing_fields = []
            if deal_value <= 0:
                missing_fields.append("Giá trị hợp đồng (<= 0)")
            if is_currently_open:
                if not decision_maker:
                    missing_fields.append("Người quyết định (Decision Maker)")
                if budget_confirmed.lower() not in ["đã xác nhận", "confirmed", "yes", "true"]:
                    missing_fields.append("Xác nhận ngân sách (Budget)")
                if not next_step:
                    missing_fields.append("Bước tiếp theo (Next Step)")

            has_missing_verification = len(missing_fields) > 0
            if has_missing_verification and is_currently_open:
                is_big_deal = deal_value >= 500000000
                opp_anomalies.append({
                    "type": "MISSING_VERIFICATION_INFO",
                    "severity": "CRITICAL" if is_big_deal else "MEDIUM",
                    "title": (
                        f"Cơ hội giá trị lớn ({deal_value/1e6:,.0f} triệu) thiếu thông tin xác nhận"
                        if is_big_deal else "Thiếu thông tin xác nhận cơ hội"
                    ),
                    "description": (
                        f"Thiếu các thông tin bắt buộc: {', '.join(missing_fields)}. "
                        + ("Không đủ điều kiện đưa vào Top ưu tiên tuần tới dù giá trị lớn!" if is_big_deal else "")
                    ),
                    "ai_suggested_action": "Yêu cầu NVKD xác minh rõ Người quyết định, Ngân sách và Bước tiếp theo trước khi phân bổ nguồn lực hỗ trợ."
                })

            # 8) Kiểm tra tỷ lệ dự kiến tùy ý của NVKD lệch chuẩn
            rule_weight = float(self.stage_weights.get(effective_stage, 0.0))
            prob_diff = abs(sales_custom_prob - rule_weight)
            if is_currently_open and prob_diff >= 0.20:
                opp_anomalies.append({
                    "type": "ARBITRARY_PROBABILITY",
                    "severity": "LOW",
                    "title": f"Tỷ lệ dự kiến tự nhập lệch quy tắc ({sales_custom_prob*100:.0f}% vs chuẩn {rule_weight*100:.0f}%)",
                    "description": (
                        f"NVKD tự nhập xác suất {sales_custom_prob*100:.0f}% cho giai đoạn '{effective_stage}', "
                        f"cao/thấp hơn quy tắc chuẩn ({rule_weight*100:.0f}%) tới {prob_diff*100:.0f}%."
                    ),
                    "ai_suggested_action": f"AI Agent đã tự động chuẩn hóa về trọng số quy tắc {rule_weight*100:.0f}%."
                })

            # =====================================================================
            # ÁP DỤNG QUYẾT ĐỊNH CỦA CON NGƯỜI TẠI BƯỚC 4 (HUMAN-IN-THE-LOOP OVERRIDE)
            # =====================================================================
            human_decision = self.human_overrides.get(opp_id)
            human_verified = False
            human_note = ""
            if human_decision:
                human_verified = True
                override_action = normalize_text(human_decision.get("action"))
                human_note = normalize_text(human_decision.get("note", "Quản lý đã xác minh tại Bước 4"))
                if override_action == "SET_CLOSED_LOST":
                    effective_stage = "Closed Lost"
                    status_resolution_reason = f"[Con người duyệt - Bước 4] Chuyển sang Closed Lost ({human_note})"
                elif override_action == "SET_CLOSED_WON":
                    effective_stage = "Closed Won"
                    status_resolution_reason = f"[Con người duyệt - Bước 4] Xác nhận Closed Won ({human_note})"
                elif override_action == "VERIFY_AND_KEEP_OPEN":
                    if effective_stage in ["Closed Lost", "Closed Won"] and not has_recorded_order:
                        effective_stage = raw_stage if raw_stage in STAGE_RULES else "Negotiation"
                    # Nếu quản lý xác nhận bổ sung thông tin và gia hạn
                    has_missing_verification = False
                    is_overdue_close = False
                    is_stale_inactive = False
                    status_resolution_reason = f"[Con người duyệt - Bước 4] Đã xác minh bổ sung thông tin & giữ mở ({human_note})"

                self.logger.log(
                    step="BƯỚC 4: Kiểm duyệt & Xác minh (Con người)",
                    skill_or_tool="HITL: ManagerVerificationOverride",
                    action=f"Quản lý phê duyệt xử lý ca nghi vấn {opp_id}",
                    rule_applied=f"RULE-HITL-OVERRIDE: Áp dụng quyết định '{override_action}' của Trưởng phòng KD",
                    details={"opp_id": opp_id, "customer_name": cust_name, "action": override_action, "note": human_note, "new_stage": effective_stage}
                )

            # Tính toán lại giá trị theo trạng thái chuẩn sau cùng
            final_is_open = effective_stage not in ["Closed Won", "Closed Lost"]
            final_rule_weight = float(self.stage_weights.get(effective_stage, 0.0))

            # Tách bạch tuyệt đối Doanh số đã ghi nhận vs Cơ hội còn mở
            if effective_stage == "Closed Won":
                actual_recorded_rev = recorded_order_amount if has_recorded_order else deal_value
                open_pipeline_val = 0.0
                weighted_forecast_val = 0.0
                naive_sales_forecast_val = deal_value * sales_custom_prob
            elif effective_stage == "Closed Lost":
                actual_recorded_rev = 0.0
                open_pipeline_val = 0.0
                weighted_forecast_val = 0.0
                naive_sales_forecast_val = deal_value * sales_custom_prob
            else:
                actual_recorded_rev = 0.0
                open_pipeline_val = deal_value
                weighted_forecast_val = deal_value * final_rule_weight
                naive_sales_forecast_val = deal_value * sales_custom_prob

            # Đánh giá điều kiện hợp lệ để xét Top 5 Ưu tiên tuần tới
            disqualify_reasons = []
            if not final_is_open:
                disqualify_reasons.append(f"Cơ hội đã đóng ({effective_stage})")
            if deal_value <= 0:
                disqualify_reasons.append("Giá trị cơ hội <= 0")
            if has_missing_verification:
                disqualify_reasons.append(f"Thiếu thông tin xác nhận ({', '.join(missing_fields)})")
            if is_overdue_close:
                disqualify_reasons.append(f"Đã quá hạn ngày dự kiến chốt ({exp_close_str} < {self.report_date_str}) chưa được xác minh gia hạn")
            if is_stale_inactive:
                disqualify_reasons.append(f"Không có tương tác mới trong {days_since_activity} ngày (> {self.inactivity_threshold_days} ngày)")

            is_eligible_for_priority = (len(disqualify_reasons) == 0)

            enriched_record = {
                "opp_id": opp_id,
                "customer_name": cust_name,
                "project_name": proj_name,
                "rep_id": rep_id,
                "raw_stage": raw_stage,
                "effective_stage": effective_stage,
                "stage_label": STAGE_RULES.get(effective_stage, {}).get("label", effective_stage),
                "status_resolution_reason": status_resolution_reason,
                "deal_value": deal_value,
                "sales_custom_prob": sales_custom_prob,
                "rule_weight": final_rule_weight,
                "actual_recorded_revenue": actual_recorded_rev,
                "open_pipeline_value": open_pipeline_val,
                "weighted_forecast_value": weighted_forecast_val,
                "naive_sales_forecast_value": naive_sales_forecast_val,
                "expected_close_date": exp_close_str,
                "days_to_close": days_to_close,
                "last_activity_date": format_date_iso(latest_act_dt),
                "days_since_activity": days_since_activity,
                "last_activity_note": latest_act_note,
                "decision_maker": decision_maker,
                "budget_confirmed": budget_confirmed,
                "next_step": next_step,
                "has_recorded_order": has_recorded_order,
                "recorded_order_id": recorded_order_id,
                "is_merged_duplicate": bool(opp.get("is_merged_duplicate")),
                "merged_from_files": opp.get("merged_from_files", []),
                "human_verified": human_verified,
                "human_note": human_note,
                "anomalies": opp_anomalies,
                "has_anomaly": len(opp_anomalies) > 0,
                "is_eligible_for_priority": is_eligible_for_priority,
                "disqualify_reasons": disqualify_reasons,
            }
            enriched_opps.append(enriched_record)

            for anom in opp_anomalies:
                anomalies_list.append({
                    "opp_id": opp_id,
                    "customer_name": cust_name,
                    "project_name": proj_name,
                    "rep_id": rep_id,
                    "deal_value": deal_value,
                    "raw_stage": raw_stage,
                    "effective_stage": effective_stage,
                    "expected_close_date": exp_close_str,
                    "last_activity_date": format_date_iso(latest_act_dt),
                    "anomaly_type": anom["type"],
                    "severity": anom["severity"],
                    "title": anom["title"],
                    "description": anom["description"],
                    "ai_suggested_action": anom["ai_suggested_action"],
                    "human_verified": human_verified,
                    "human_note": human_note,
                })

        self.logger.log(
            step="BƯỚC 3: Phát hiện rủi ro & Mâu thuẫn",
            skill_or_tool="Skill 2: RiskAndConflictDetectionSkill",
            action="Quét toàn bộ cơ hội để phát hiện mâu thuẫn trạng thái, quá hạn và thiếu thông tin xác nhận",
            rule_applied=f"RULE-ANOMALY-SCAN: Quá hạn khi expected_close_date < {self.report_date_str}; Đóng băng khi không tương tác > {self.inactivity_threshold_days} ngày; Kiểm tra chéo Orders & Activities",
            details={
                "total_anomalies_detected": len(anomalies_list),
                "critical_count": sum(1 for a in anomalies_list if a["severity"] == "CRITICAL"),
                "high_count": sum(1 for a in anomalies_list if a["severity"] == "HIGH"),
                "medium_count": sum(1 for a in anomalies_list if a["severity"] == "MEDIUM"),
            }
        )

        return anomalies_list, enriched_opps

    # -------------------------------------------------------------------------
    # SKILL 3: TÍNH MỨC HOÀN THÀNH MỤC TIÊU (KPI) & ĐỀ XUẤT TOP 5 CƠ HỘI ƯU TIÊN
    # -------------------------------------------------------------------------
    def _skill_3_evaluate_kpi_and_prioritize(
        self,
        enriched_opps: List[Dict[str, Any]],
        valid_orders_df: pd.DataFrame,
        kpi_df: pd.DataFrame
    ) -> Tuple[List[Dict[str, Any]], Dict[str, Any], Dict[str, Any]]:
        rep_map: Dict[str, Dict[str, Any]] = {}
        for _, row in kpi_df.iterrows():
            rid = normalize_text(row.get("rep_id"))
            if not rid:
                continue
            rep_map[rid] = {
                "rep_id": rid,
                "rep_name": normalize_text(row.get("rep_name", rid)),
                "team": normalize_text(row.get("team", "B2B Sales")),
                "period": normalize_text(row.get("period", "Q4/2026")),
                "kpi_target": float(row.get("kpi_target") or 0),
                "recorded_revenue": 0.0,
                "closed_won_count": 0,
                "open_opp_count": 0,
                "open_pipeline_value": 0.0,
                "weighted_forecast_value": 0.0,
                "naive_uncleaned_forecast": 0.0,
                "closed_lost_count": 0,
                "anomaly_opp_count": 0,
            }

        # Cộng dồn từ danh sách cơ hội đã chuẩn hóa
        counted_order_opp_ids = set()
        for opp in enriched_opps:
            rid = opp["rep_id"]
            if rid not in rep_map:
                rep_map[rid] = {
                    "rep_id": rid,
                    "rep_name": rid,
                    "team": "B2B Sales",
                    "period": "Q4/2026",
                    "kpi_target": 1500000000.0,
                    "recorded_revenue": 0.0,
                    "closed_won_count": 0,
                    "open_opp_count": 0,
                    "open_pipeline_value": 0.0,
                    "weighted_forecast_value": 0.0,
                    "naive_uncleaned_forecast": 0.0,
                    "closed_lost_count": 0,
                    "anomaly_opp_count": 0,
                }
            r_entry = rep_map[rid]
            r_entry["naive_uncleaned_forecast"] += opp["naive_sales_forecast_value"]
            if opp["has_anomaly"]:
                r_entry["anomaly_opp_count"] += 1

            st = opp["effective_stage"]
            if st == "Closed Won":
                r_entry["recorded_revenue"] += opp["actual_recorded_revenue"]
                r_entry["closed_won_count"] += 1
                counted_order_opp_ids.add(opp["opp_id"])
            elif st == "Closed Lost":
                r_entry["closed_lost_count"] += 1
            else:
                r_entry["open_opp_count"] += 1
                r_entry["open_pipeline_value"] += opp["open_pipeline_value"]
                r_entry["weighted_forecast_value"] += opp["weighted_forecast_value"]

        # Kiểm tra xem còn đơn hàng độc lập nào trong valid_orders_df chưa gắn với opp_id không
        if not valid_orders_df.empty:
            for _, orow in valid_orders_df.iterrows():
                oid = normalize_text(orow.get("opp_id"))
                rid = normalize_text(orow.get("rep_id"))
                amt = float(orow.get("order_amount") or 0)
                if oid not in counted_order_opp_ids and rid in rep_map:
                    rep_map[rid]["recorded_revenue"] += amt
                    rep_map[rid]["closed_won_count"] += 1

        rep_summary_list = []
        for rid, data in sorted(rep_map.items(), key=lambda x: x[0]):
            target = data["kpi_target"] if data["kpi_target"] > 0 else 1.0
            actual_pct = round((data["recorded_revenue"] / target) * 100, 1)
            projected_total = data["recorded_revenue"] + data["weighted_forecast_value"]
            projected_pct = round((projected_total / target) * 100, 1)
            kpi_gap = max(0.0, target - data["recorded_revenue"])
            projected_gap = max(0.0, target - projected_total)

            if projected_pct >= 100:
                status_badge = "Đạt dự kiến (On Track)"
            elif projected_pct >= 75:
                status_badge = "Cần thúc đẩy (At Risk)"
            else:
                status_badge = "Thiếu hụt lớn (Critical Gap)"

            data.update({
                "actual_attainment_pct": actual_pct,
                "projected_total_estimate": projected_total,
                "projected_attainment_pct": projected_pct,
                "kpi_gap_actual": kpi_gap,
                "kpi_gap_projected": projected_gap,
                "kpi_status": status_badge,
            })
            rep_summary_list.append(data)

        # Tổng hợp cấp Công ty Bình Minh
        total_target = sum(r["kpi_target"] for r in rep_summary_list)
        total_recorded = sum(r["recorded_revenue"] for r in rep_summary_list)
        total_open_pipeline = sum(r["open_pipeline_value"] for r in rep_summary_list)
        total_weighted_forecast = sum(r["weighted_forecast_value"] for r in rep_summary_list)
        total_naive_forecast = sum(r["naive_uncleaned_forecast"] for r in rep_summary_list)
        total_projected = total_recorded + total_weighted_forecast

        company_kpis = {
            "total_sales_reps": len(rep_summary_list),
            "total_unique_opportunities": len(enriched_opps),
            "total_open_opportunities": sum(r["open_opp_count"] for r in rep_summary_list),
            "total_closed_won": sum(r["closed_won_count"] for r in rep_summary_list),
            "total_closed_lost": sum(r["closed_lost_count"] for r in rep_summary_list),
            "total_kpi_target": total_target,
            "total_recorded_revenue": total_recorded,
            "total_open_pipeline_value": total_open_pipeline,
            "total_weighted_forecast_estimate": total_weighted_forecast,
            "total_projected_revenue_estimate": total_projected,
            "total_naive_uncleaned_forecast": total_naive_forecast,
            "forecast_inflation_removed": max(0.0, (total_recorded + total_naive_forecast) - total_projected),
            "company_actual_attainment_pct": round((total_recorded / total_target) * 100, 1) if total_target > 0 else 0.0,
            "company_projected_attainment_pct": round((total_projected / total_target) * 100, 1) if total_target > 0 else 0.0,
            "disclaimer": "LƯU Ý QUAN TRỌNG: 'Giá trị dự kiến có trọng số' và 'Tổng doanh số dự kiến' là ƯỚC TÍNH THEO QUY TẮC TRỌNG SỐ GIAI ĐOẠN, KHÔNG PHẢI DOANH SỐ CHẮC CHẮN."
        }

        # =====================================================================
        # CHẤM ĐIỂM & LỰA CHỌN TOP 5 CƠ HỘI CẦN ƯU TIÊN XỬ LÝ TRONG TUẦN TIẾP THEO
        # =====================================================================
        eligible_opps = [o for o in enriched_opps if o["is_eligible_for_priority"]]
        disqualified_big_deals = [
            o for o in enriched_opps
            if not o["is_eligible_for_priority"]
            and o["effective_stage"] not in ["Closed Won", "Closed Lost"]
            and o["deal_value"] >= 500000000
        ]

        max_weighted_val = max([o["weighted_forecast_value"] for o in eligible_opps], default=1.0)
        if max_weighted_val <= 0:
            max_weighted_val = 1.0

        stage_score_map = {
            "Closing": 25.0,
            "Negotiation": 20.0,
            "Proposal": 14.0,
            "Qualification": 8.0,
            "Lead": 4.0,
        }

        scored_candidates = []
        for opp in eligible_opps:
            rep_info = rep_map.get(opp["rep_id"], {})
            # 1. Điểm giá trị dự kiến có trọng số (Tối đa 35 điểm)
            val_score = round((opp["weighted_forecast_value"] / max_weighted_val) * 35.0, 1)

            # 2. Điểm giai đoạn sẵn sàng chốt (Tối đa 25 điểm)
            st_score = stage_score_map.get(opp["effective_stage"], 5.0)

            # 3. Điểm độ khẩn cấp theo ngày dự kiến chốt trong tuần tới (Tối đa 20 điểm)
            d_close = opp["days_to_close"]
            if d_close is not None:
                if 0 <= d_close <= 7:
                    urg_score = 20.0
                elif 8 <= d_close <= 14:
                    urg_score = 16.0
                elif 15 <= d_close <= 21:
                    urg_score = 11.0
                else:
                    urg_score = 6.0
            else:
                urg_score = 5.0

            # 4. Điểm hỗ trợ hoàn thành KPI của NVKD (Tối đa 20 điểm)
            rep_actual_pct = rep_info.get("actual_attainment_pct", 50.0)
            rep_gap_ratio = max(0.0, min(1.0, (100.0 - rep_actual_pct) / 100.0))
            kpi_support_score = round(rep_gap_ratio * 20.0, 1)

            total_score = round(val_score + st_score + urg_score + kpi_support_score, 1)

            # Xây dựng căn cứ lựa chọn & hành động đề xuất cụ thể cho tuần tiếp theo
            rationale = (
                f"Giá trị dự kiến ước tính đạt {opp['weighted_forecast_value']/1e6:,.1f} triệu VNĐ "
                f"(Gốc: {opp['deal_value']/1e6:,.0f} triệu x Trọng số {opp['effective_stage']} {opp['rule_weight']*100:.0f}%). "
                f"Dự kiến chốt trong {d_close} ngày tới ({opp['expected_close_date']}). "
                f"Đã xác nhận ngân sách và người quyết định ({opp['decision_maker']}). "
                f"Đóng góp quan trọng giúp NVKD {rep_info.get('rep_name', opp['rep_id'])} thu hẹp khoảng cách KPI ({rep_info.get('kpi_gap_actual', 0)/1e6:,.0f} triệu)."
            )

            if opp["effective_stage"] == "Closing":
                action_plan = (
                    f"[Ưu tiên Chốt HĐ] Trưởng phòng KD cùng {rep_info.get('rep_name', opp['rep_id'])} làm việc trực tiếp với "
                    f"{opp['decision_maker']} để hoàn tất: \"{opp['next_step']}\" trước ngày {opp['expected_close_date']}."
                )
                owner_support = "Trưởng phòng KD + Pháp chế hỗ trợ rà soát điều khoản ký kết"
            elif opp["effective_stage"] == "Negotiation":
                action_plan = (
                    f"[Thúc đẩy Đàm phán] Hỗ trợ {rep_info.get('rep_name', opp['rep_id'])} tháo gỡ vướng mắc thương mại/kỹ thuật "
                    f"với {opp['decision_maker']}; thực hiện \"{opp['next_step']}\" và chốt biên bản đồng thuận."
                )
                owner_support = "Trưởng phòng KD + Trưởng nhóm Giải pháp (Presales)"
            else:
                action_plan = (
                    f"[Nâng cấp Giai đoạn] Tổ chức buổi trình bày/Demo chuyên sâu cho {opp['decision_maker']}: "
                    f"\"{opp['next_step']}\", chuyển cơ hội từ {opp['effective_stage']} sang Đàm phán hợp đồng."
                )
                owner_support = "Chuyên gia Tư vấn Giải pháp B2B (Presales Lead)"

            opp_copy = dict(opp)
            opp_copy.update({
                "rep_name": rep_info.get("rep_name", opp["rep_id"]),
                "priority_score": total_score,
                "score_breakdown": {
                    "weighted_value_score": val_score,
                    "stage_readiness_score": st_score,
                    "urgency_score": urg_score,
                    "kpi_gap_support_score": kpi_support_score,
                },
                "selection_rationale": rationale,
                "recommended_action_next_week": action_plan,
                "support_assignment": owner_support,
            })
            scored_candidates.append(opp_copy)

        scored_candidates.sort(key=lambda x: (x["priority_score"], x["weighted_forecast_value"]), reverse=True)
        top_5_list = scored_candidates[:5]
        for idx, item in enumerate(top_5_list, start=1):
            item["rank"] = idx

        is_under_5 = len(top_5_list) < 5
        if is_under_5:
            shortfall_explanation = (
                f"CẢNH BÁO RANH GIỚI QUY TẮC: Hiện chỉ có {len(top_5_list)}/5 cơ hội hợp lệ đủ điều kiện đưa vào danh sách ưu tiên tuần tiếp theo. "
                f"Lý do: Có {len(disqualified_big_deals)} cơ hội giá trị lớn đang mở bị loại trừ vì thiếu thông tin xác nhận bắt buộc (Người quyết định/Ngân sách/Bước tiếp theo) "
                f"hoặc đã quá hạn chốt/không có tương tác mới. Hệ thống tuyệt đối không đưa cơ hội chưa xác nhận vào Top ưu tiên theo đúng quy tắc nghiệp vụ."
            )
        else:
            shortfall_explanation = (
                f"Đã trích xuất đủ 5/5 cơ hội hợp lệ có điểm ưu tiên cao nhất trên tổng số {len(scored_candidates)} cơ hội mở đủ điều kiện xác nhận. "
                f"(Đã loại trừ {len(disqualified_big_deals)} cơ hội giá trị lớn nhưng thiếu thông tin xác nhận hoặc quá hạn)."
            )

        priority_result = {
            "requested_top_k": 5,
            "actual_valid_count": len(top_5_list),
            "total_eligible_pool": len(scored_candidates),
            "is_under_5_boundary": is_under_5,
            "boundary_explanation": shortfall_explanation,
            "estimate_disclaimer": "Giá trị dự kiến của các cơ hội dưới đây là ƯỚC TÍNH THEO QUY TẮC TRỌNG SỐ, không phải doanh số chắc chắn.",
            "top_opportunities": top_5_list,
            "excluded_high_value_opps": [
                {
                    "opp_id": o["opp_id"],
                    "customer_name": o["customer_name"],
                    "project_name": o["project_name"],
                    "rep_id": o["rep_id"],
                    "deal_value": o["deal_value"],
                    "effective_stage": o["effective_stage"],
                    "reasons_excluded": o["disqualify_reasons"]
                }
                for o in disqualified_big_deals
            ]
        }

        self.logger.log(
            step="BƯỚC 5: Phân tích KPI & Ưu tiên",
            skill_or_tool="Skill 3: KPIAndPriorityRecommendationSkill",
            action="Tính toán mức hoàn thành KPI theo nhân viên và xếp hạng Top 5 cơ hội ưu tiên tuần tới",
            rule_applied="RULE-PRIORITY-TOP5: Chỉ xét cơ hội đang mở, đủ thông tin xác nhận, không quá hạn; nếu < 5 cơ hội hợp lệ thì báo cáo số thực tế kèm lý do",
            details={
                "company_recorded_revenue": total_recorded,
                "company_weighted_forecast_estimate": total_weighted_forecast,
                "eligible_opportunities_count": len(scored_candidates),
                "top_selected_count": len(top_5_list),
                "is_under_5_boundary": is_under_5,
                "top_opp_ids": [x["opp_id"] for x in top_5_list]
            }
        )

        return rep_summary_list, company_kpis, priority_result

    # -------------------------------------------------------------------------
    # KẾT NỐI LLM (GEMINI / OPENAI) HOẶC AI REASONING ENGINE NỘI BỘ
    # -------------------------------------------------------------------------
    def _generate_executive_synthesis(
        self,
        company_kpis: Dict[str, Any],
        rep_summary: List[Dict[str, Any]],
        anomalies: List[Dict[str, Any]],
        priority_result: Dict[str, Any]
    ) -> Dict[str, Any]:
        prompt_summary = {
            "report_date": self.report_date_str,
            "recorded_revenue_vnd": company_kpis["total_recorded_revenue"],
            "open_pipeline_vnd": company_kpis["total_open_pipeline_value"],
            "weighted_forecast_estimate_vnd": company_kpis["total_weighted_forecast_estimate"],
            "kpi_target_vnd": company_kpis["total_kpi_target"],
            "actual_attainment_pct": company_kpis["company_actual_attainment_pct"],
            "projected_attainment_pct": company_kpis["company_projected_attainment_pct"],
            "anomalies_count": len(anomalies),
            "top_priority_count": priority_result["actual_valid_count"],
            "boundary_note": priority_result["boundary_explanation"],
        }

        # Built-in synthesis (luôn có sẵn, nhanh và chuẩn xác)
        builtin_text = (
            f"### Báo cáo Điều hành AI Agent (Ngày đối chiếu: {self.report_date_str})\n\n"
            f"1. **Tình hình Thực hiện Mục tiêu (KPI)**:\n"
            f"   - **Doanh số đã ghi nhận (Chắc chắn)**: **{company_kpis['total_recorded_revenue']/1e6:,.1f} triệu VNĐ** "
            f"(đạt **{company_kpis['company_actual_attainment_pct']}%** so với mục tiêu {company_kpis['total_kpi_target']/1e6:,.0f} triệu VNĐ).\n"
            f"   - **Giá trị cơ hội còn mở**: **{company_kpis['total_open_pipeline_value']/1e6:,.1f} triệu VNĐ** "
            f"trên **{company_kpis['total_open_opportunities']} cơ hội đang mở**.\n"
            f"   - **Giá trị dự kiến theo trọng số quy tắc (Ước tính, không phải doanh số chắc chắn)**: "
            f"**{company_kpis['total_weighted_forecast_estimate']/1e6:,.1f} triệu VNĐ**, đưa tổng dự kiến ước tính lên "
            f"**{company_kpis['total_projected_revenue_estimate']/1e6:,.1f} triệu VNĐ ({company_kpis['company_projected_attainment_pct']}% KPI)**.\n"
            f"   - **Làm sạch dữ liệu ảo**: Hệ thống đã loại bỏ **{company_kpis['forecast_inflation_removed']/1e6:,.1f} triệu VNĐ** "
            f"doanh số bị tính trùng hoặc thổi phồng do tỷ lệ dự kiến tùy ý.\n\n"
            f"2. **Cảnh báo Rủi ro & Mâu thuẫn cần Con người xác minh (Bước 4)**:\n"
            f"   - Phát hiện **{len(anomalies)} cảnh báo** (bao gồm cơ hội đã có đơn hàng nhưng quên đóng, cơ hội khách đã từ chối nhưng vẫn để mở, cơ hội quá hạn chốt và thương vụ lớn thiếu thông tin xác nhận).\n\n"
            f"3. **Đề xuất Ưu tiên Tuần tiếp theo (Bước 5)**:\n"
            f"   - {priority_result['boundary_explanation']}"
        )

        if self.api_key and self.llm_provider in ["gemini", "openai"]:
            sys_prompt = (
                "Bạn là Chuyên gia Phân tích Vận hành Kinh doanh B2B (Sales Operations AI Agent) của Công ty Bình Minh. "
                "Hãy viết bản nhận xét điều hành ngắn gọn (3 phần: Đánh giá KPI & Dự báo ước tính, Rủi ro dữ liệu cần xử lý, Chỉ đạo hành động tuần tới) "
                "dựa trên số liệu JSON được cung cấp. Nhấn mạnh rõ giá trị dự kiến chỉ là ước tính theo quy tắc trọng số."
            )
            user_msg = f"Số liệu tổng hợp tại ngày đối chiếu {self.report_date_str}:\n{json.dumps(prompt_summary, ensure_ascii=False, indent=2)}"
            try:
                if self.llm_provider == "gemini":
                    gemini_success = False
                    for model_name in GEMINI_MODELS:
                        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={self.api_key}"
                        payload = {
                            "contents": [{"parts": [{"text": sys_prompt + "\n\n" + user_msg}]}],
                            "generationConfig": {"temperature": 0.2}
                        }
                        try:
                            r = requests.post(url, json=payload, timeout=12)
                            if r.status_code == 200:
                                data = r.json()
                                llm_text = data["candidates"][0]["content"]["parts"][0]["text"]
                                self.logger.log(
                                    step="BƯỚC 5: Phân tích KPI & Ưu tiên",
                                    skill_or_tool=f"LLM: Google Gemini API ({model_name})",
                                    action="Sinh báo cáo nhận xét điều hành bằng mô hình ngôn ngữ lớn",
                                    rule_applied="RULE-LLM-SYNTHESIS: Dựa 100% trên kết quả tính toán từ Tool, không bịa đặt số liệu",
                                    details={"provider": "gemini", "model": model_name, "status": "SUCCESS"}
                                )
                                return {"provider_used": f"Google Gemini API ({model_name})", "content": llm_text}
                        except Exception:
                            continue
                elif self.llm_provider == "openai":
                    url = "https://api.openai.com/v1/chat/completions"
                    headers = {"Authorization": f"Bearer {self.api_key}", "Content-Type": "application/json"}
                    payload = {
                        "model": "gpt-4o-mini",
                        "messages": [
                            {"role": "system", "content": sys_prompt},
                            {"role": "user", "content": user_msg}
                        ],
                        "temperature": 0.2
                    }
                    r = requests.post(url, headers=headers, json=payload, timeout=15)
                    if r.status_code == 200:
                        data = r.json()
                        llm_text = data["choices"][0]["message"]["content"]
                        self.logger.log(
                            step="BƯỚC 5: Phân tích KPI & Ưu tiên",
                            skill_or_tool="LLM: OpenAI API (gpt-4o-mini)",
                            action="Sinh báo cáo nhận xét điều hành bằng mô hình ngôn ngữ lớn",
                            rule_applied="RULE-LLM-SYNTHESIS: Dựa 100% trên kết quả tính toán từ Tool, không bịa đặt số liệu",
                            details={"provider": "openai", "status": "SUCCESS"}
                        )
                        return {"provider_used": "OpenAI API (gpt-4o-mini)", "content": llm_text}
            except Exception as ex:
                self.logger.log(
                    step="BƯỚC 5: Phân tích KPI & Ưu tiên",
                    skill_or_tool=f"LLM: {self.llm_provider} (Fallback to Built-in Engine)",
                    action="Chuyển sang AI Reasoning Engine nội bộ do lỗi kết nối API ngoài",
                    rule_applied="RULE-FALLBACK-SAFE",
                    details={"error": str(ex)},
                    status="WARNING"
                )

        self.logger.log(
            step="BƯỚC 5: Phân tích KPI & Ưu tiên",
            skill_or_tool="AI Reasoning Engine (Built-in Deterministic Synthesizer)",
            action="Sinh bản tổng hợp điều hành & khuyến nghị hành động",
            rule_applied="RULE-SYNTHESIS-DETERMINISTIC: Tổng hợp nhận xét từ kết quả của 3 kỹ năng nghiệp vụ",
            details={"provider": "Built-in AI Reasoning Engine"}
        )
        return {"provider_used": "Built-in AI Reasoning Engine (Chuẩn quy tắc Đề 09)", "content": builtin_text}

