import pandas as pd
import numpy as np
import json
import requests
from pathlib import Path

CONFIG_PATH = Path(__file__).parent / "config.json"

def load_config():
    with open(CONFIG_PATH, "r") as f:
        return json.load(f)

def save_config(config):
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

def load_product_list(config):
    prod = pd.read_excel(config["product_list_path"], header=0)
    prod.columns = ["상품", "구분", "이용권명", "TRNT대구분", "TRNT중구분"]
    prod["이용권명"] = prod["이용권명"].str.strip()
    prod_map = dict(zip(prod["이용권명"], prod[["TRNT대구분", "TRNT중구분"]].values.tolist()))
    manual = {
        "[개인] 개인 레슨 10회": ["개인", "개인레슨"],
        "[개인] 패키지) 개인 레슨 5회": ["개인", "개인레슨"],
        "[OT] 개인레슨 무료쿠폰 1 회": ["개인", "개인레슨"],
        "[개인] A_교육_IT 본교육 메이크업 보강 _2회권": ["아카데미", "딥코칭"],
        "[쿠폰] 자유그룹쿠폰 20회": ["그룹", "그룹레슨"],
        "[개인] IT 개인레슨 1회": ["개인", "개인레슨"],
        "[그룹] A_교육_그룹 취업컨설팅": ["아카데미", "그룹"],
    }
    prod_map.update(manual)
    # config에 저장된 수동 매핑도 적용
    config_manual = config.get("manual_product_map", {})
    for k, v in config_manual.items():
        prod_map[k] = v
    return prod_map

def classify_lesson(name, prod_map):
    name = name.strip()
    if name in prod_map:
        return prod_map[name]
    clean = " ".join(name.split())
    if clean in prod_map:
        return prod_map[clean]
    for k, v in prod_map.items():
        if k.strip() == name:
            return v
    return ["기타", "기타"]

def load_excel_data(config):
    prod_map = load_product_list(config)
    dfs = []
    for month, files in config["data_files"].items():
        for f in files:
            if not Path(f).exists():
                continue
            df = pd.read_excel(f, header=0)
            df["월"] = month
            dfs.append(df)
    if not dfs:
        return pd.DataFrame()
    all_data = pd.concat(dfs, ignore_index=True)
    all_data["수업일자"] = pd.to_datetime(all_data["수업일자"])
    all_data["이용권명"] = all_data["이용권명"].str.strip()
    classifications = all_data["이용권명"].apply(lambda x: classify_lesson(x, prod_map))
    all_data["TRNT대구분"] = [c[0] for c in classifications]
    all_data["TRNT중구분"] = [c[1] for c in classifications]
    all_data = all_data[all_data["강사명"] != "강사 공용"].copy()
    all_data["강사"] = all_data["강사명"].str.replace(" 선생님", "", regex=False)
    # 정산승인 + 정산취소(결석) 모두 수업 완료로 카운트
    all_data["수업완료"] = all_data["정산현황"].isin(["정산승인", "정산취소"])
    return all_data

def fetch_instructor_info(config):
    """노션 인적사항 DB에서 강사 이름/이메일 가져오기"""
    token = config.get("notion_api_token", "")
    db_id = config.get("notion_staff_db_id", "")
    if not token or not db_id:
        return {}
    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json",
    }
    result = {}
    has_more = True
    start_cursor = None
    while has_more:
        body = {"page_size": 100}
        if start_cursor:
            body["start_cursor"] = start_cursor
        try:
            resp = requests.post(
                f"https://api.notion.com/v1/databases/{db_id}/query",
                headers=headers, json=body, timeout=15
            )
            data = resp.json()
        except Exception:
            break
        for r in data.get("results", []):
            props = r["properties"]
            # 이름 필드 찾기 (title 타입)
            name = ""
            email = ""
            for key, val in props.items():
                if val.get("type") == "title":
                    title_parts = val.get("title", [])
                    name = "".join([t["plain_text"] for t in title_parts]).strip()
                if "이메일" in key or key.lower() in ("email",):
                    e = ""
                    if val.get("type") == "email":
                        e = val.get("email") or ""
                    elif val.get("type") == "rich_text":
                        rt = val.get("rich_text", [])
                        e = "".join([t["plain_text"] for t in rt]).strip()
                    if e and not email:
                        email = e
            if name and email:
                result[name] = email
        has_more = data.get("has_more", False)
        start_cursor = data.get("next_cursor")
    return result


def fetch_notion_data(config):
    token = config["notion_api_token"]
    db_id = config["notion_db_id"]
    headers = {
        "Authorization": f"Bearer {token}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json",
    }
    results = []
    has_more = True
    start_cursor = None
    while has_more:
        body = {"page_size": 100}
        if start_cursor:
            body["start_cursor"] = start_cursor
        try:
            resp = requests.post(
                f"https://api.notion.com/v1/databases/{db_id}/query",
                headers=headers, json=body, timeout=15
            )
            data = resp.json()
        except Exception:
            break
        for r in data.get("results", []):
            props = r["properties"]
            title_parts = props.get("레슨마감 \u2018월\u2019+강사명", {}).get("title", [])
            title = "".join([t["plain_text"] for t in title_parts])
            date_obj = props.get("레슨 마감 일자", {}).get("date")
            results.append({
                "title": title,
                "전체고객수": props.get("전체 고객 수", {}).get("number"),
                "홀딩고객수": props.get("홀딩 고객 수", {}).get("number"),
                "재등예정수": props.get("재등 예정 수", {}).get("number"),
                "재등완료수": props.get("재등 완료 수", {}).get("number"),
                "체험수업수": props.get("체험 수업 수", {}).get("number"),
                "체험등록수": props.get("체험 등록 수", {}).get("number"),
                "마감일": date_obj["start"] if date_obj else None,
            })
        has_more = data.get("has_more", False)
        start_cursor = data.get("next_cursor")
    notion_by_month = {}
    for r in results:
        title = r["title"]
        parts = title.split(" ", 1)
        if len(parts) != 2:
            continue
        instructor = parts[1]
        entry = {k: r[k] for k in ["전체고객수", "홀딩고객수", "재등예정수", "재등완료수", "체험수업수", "체험등록수"]}
        entry["마감일"] = r["마감일"]

        # 마감일자 기준으로 월 매칭 (우선), 없으면 타이틀 월 사용
        month_key = None
        if r["마감일"]:
            try:
                dt = pd.to_datetime(r["마감일"])
                month_key = f"{dt.year}년 {dt.month}월"
            except Exception:
                pass
        if not month_key:
            month_key = parts[0]  # 타이틀에서 추출 (예: "3월")

        if month_key not in notion_by_month:
            notion_by_month[month_key] = {}
        notion_by_month[month_key][instructor] = entry

        # 타이틀 월 키도 추가 (하위 호환)
        title_month = parts[0]
        if title_month != month_key:
            if title_month not in notion_by_month:
                notion_by_month[title_month] = {}
            if instructor not in notion_by_month[title_month]:
                notion_by_month[title_month][instructor] = entry

    return notion_by_month

def aggregate_instructor(df_month, notion_month_data, config):
    W = config["weeks_per_month"]
    results = []
    for instructor in sorted(df_month["강사"].unique()):
        idf = df_month[df_month["강사"] == instructor]
        done = idf[idf["수업완료"]]

        # 개인레슨 (정산승인 기준)
        personal = done[(done["TRNT대구분"] == "개인") & (done["TRNT중구분"] == "개인레슨")]
        personal_count = len(personal)

        # 개인OT
        personal_ot = done[(done["TRNT대구분"] == "개인") & (done["TRNT중구분"] == "개인OT")]
        ot_count = len(personal_ot)

        # 듀엣
        duet = done[done["TRNT대구분"] == "듀엣"]
        duet_lesson_count = len(duet)
        duet_members = duet["회원명"].nunique() if len(duet) > 0 else 0

        # 그룹
        group = done[done["TRNT대구분"] == "그룹"]
        group_attend = len(group)
        group_members = group["회원명"].nunique() if len(group) > 0 else 0
        group_sessions = group.groupby("수업일자").ngroups if len(group) > 0 else 0
        # 그룹수업수 = 정산승인된 전체 그룹 레코드 수 (출석+결석 모두 포함)
        group_all = idf[idf["TRNT대구분"] == "그룹"]
        group_total_records = len(group_all[group_all["수업완료"]])

        # 아카데미 (중구분별)
        academy = done[done["TRNT대구분"] == "아카데미"]
        academy_count = len(academy)
        academy_deep = len(academy[academy["TRNT중구분"] == "딥코칭"])
        academy_mock = len(academy[academy["TRNT중구분"] == "모의테스트"])
        academy_trial = len(academy[academy["TRNT중구분"] == "체험"])
        academy_group = len(academy[academy["TRNT중구분"] == "그룹"])

        # 노션 데이터
        n = notion_month_data.get(instructor, {})
        total_clients = n.get("전체고객수")
        holding = n.get("홀딩고객수")
        re_plan = n.get("재등예정수")
        re_done = n.get("재등완료수")
        trial_lesson = n.get("체험수업수")
        trial_reg = n.get("체험등록수")

        # 수식 계산
        lesson_plus_ot = personal_count + ot_count

        min_lesson = total_clients * 1.0 * W if total_clients else None
        target_lesson = total_clients * 1.5 * W if total_clients else None
        max_lesson = total_clients * 2.0 * W if total_clients else None

        personal_attend_rate = (lesson_plus_ot / total_clients) / W if total_clients and total_clients > 0 else None
        personal_achieve_rate = lesson_plus_ot / target_lesson if target_lesson and target_lesson > 0 else None

        re_reg_rate = re_done / re_plan if re_plan and re_plan > 0 and re_done is not None else None
        trial_reg_rate = trial_reg / trial_lesson if trial_lesson and trial_lesson > 0 and trial_reg is not None else None

        duet_attend_rate = (duet_lesson_count / duet_members) / W if duet_members and duet_members > 0 else None

        group_attend_rate = group_attend / group_sessions if group_sessions and group_sessions > 0 else None

        results.append({
            "강사": instructor,
            "전체고객수": total_clients,
            "홀딩고객수": holding,
            "개인레슨수": personal_count,
            "개인OT수": ot_count,
            "개인레슨+OT": lesson_plus_ot,
            "최소레슨수": round(min_lesson, 1) if min_lesson else None,
            "목표레슨수": round(target_lesson, 1) if target_lesson else None,
            "최대레슨수": round(max_lesson, 1) if max_lesson else None,
            "개인출석률": personal_attend_rate,
            "개인출석달성율": personal_achieve_rate,
            "체험수업수": trial_lesson,
            "체험등록수": trial_reg,
            "체험승률": trial_reg_rate,
            "듀엣회원수": duet_members,
            "듀엣레슨수": duet_lesson_count,
            "듀엣출석률": duet_attend_rate,
            "그룹회원수": group_members,
            "그룹수업수": group_sessions,
            "그룹출석수": group_attend,
            "그룹출석율": group_attend_rate,
            "재등예정수": re_plan,
            "재등완료수": re_done,
            "재등록율": re_reg_rate,
            "아카데미수업수": academy_count,
            "아카데미_딥코칭": academy_deep,
            "아카데미_모의테스트": academy_mock,
            "아카데미_체험": academy_trial,
            "아카데미_그룹": academy_group,
            "총수업수": len(done) - academy_count,
        })
    return pd.DataFrame(results)

def extract_month_num(month_str):
    """'2025년 3월' → '3', '3월' → '3'"""
    s = month_str.strip()
    if "년" in s:
        s = s.split("년")[-1].strip()
    return s.replace("월", "").strip()

def get_all_reports(config):
    all_data = load_excel_data(config)
    if all_data.empty:
        return {}
    notion = fetch_notion_data(config)
    months = sorted(all_data["월"].unique())
    reports = {}
    for month in months:
        month_data = all_data[all_data["월"] == month]
        month_num = extract_month_num(month)

        # 노션 매칭 우선순위:
        # 1. 정확한 "2025년 3월" 키 (마감일자 기반)
        # 2. "3월" 키 (타이틀 기반, 하위 호환)
        notion_month = notion.get(month, {})  # "2025년 3월" 정확 매칭
        if not notion_month:
            notion_month = notion.get(f"{month_num}월", notion.get(month_num, {}))
        if not notion_month:
            notion_month = {}

        reports[month] = aggregate_instructor(month_data, notion_month, config)
    return reports
