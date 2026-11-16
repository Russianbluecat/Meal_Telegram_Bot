"""식단 리마인드 봇 (GitHub Actions용)

흐름:
  1) 오늘 식단 목록 + "다 입력하셨나요?" 버튼 전송
  2) [네] -> "LDL 관점에서 분석할까요?" 버튼
        [네]     -> 오늘 + 최근 2주 분석 (Gemini)
        [괜찮아요] -> 마무리 멘트
     [아니요] -> 시트 링크 안내
  버튼 답변은 최대 WAIT_MINUTES분 동안 기다림 (그 안에 안 누르면 조용히 종료)
"""
import json
import os
import time
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import gspread
import requests

# ===== 설정 (비밀값은 전부 GitHub Secrets에서 가져옴) =====
TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
CHAT_ID = str(os.environ["CHAT_ID"])
GEMINI_API_KEY = os.environ["GEMINI_API_KEY"]
SHEET_ID = os.environ["SHEET_ID"]
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]

SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit"
MEAL_SHEET_NAME = "Meal_Plan"      # A:날짜, B:시간, C:메뉴
CHECKUP_SHEET_NAME = "Check_Up"    # A:검진일자, B:항목, C:수치, D:참고치
LDL_ITEM_KEYWORD = "LDL"

KST = ZoneInfo("Asia/Seoul")
GEMINI_MODEL = "gemini-2.0-flash"
MAX_MEALS_FOR_ANALYSIS = 300
TELEGRAM_MAX_MESSAGE_LENGTH = 4000
WAIT_MINUTES = int(os.environ.get("WAIT_MINUTES", "10"))
ANALYSIS_DAYS = 14

API = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}"

# 버튼 값
CONFIRM_YES = "confirm_yes"
CONFIRM_NO = "confirm_no"
ANALYZE_YES = "analyze_offer_yes"
ANALYZE_NO = "analyze_offer_no"

state = {"offset": None}  # getUpdates에서 어디까지 읽었는지 (엑셀의 '마지막 읽은 행 번호' 같은 것)


# =====================================================
# 1. 텔레그램
# =====================================================
def tg(method, **payload):
    r = requests.post(f"{API}/{method}", json=payload, timeout=40)
    r.raise_for_status()
    return r.json()["result"]


def send_buttons(text, rows):
    """버튼 메시지를 보내고, 그 메시지의 번호(message_id)를 돌려줌"""
    result = tg(
        "sendMessage",
        chat_id=CHAT_ID,
        text=text,
        reply_markup={"inline_keyboard": rows},
    )
    return result["message_id"]


def split_text(text, max_length):
    chunks, remaining = [], text
    while len(remaining) > max_length:
        idx = remaining.rfind("\n", 0, max_length)
        if idx < max_length * 0.5:
            idx = max_length
        chunks.append(remaining[:idx])
        remaining = remaining[idx:].lstrip("\n")
    if remaining:
        chunks.append(remaining)
    return chunks


def send_text(text):
    for chunk in split_text(str(text), TELEGRAM_MAX_MESSAGE_LENGTH):
        tg("sendMessage", chat_id=CHAT_ID, text=chunk)


def wait_for_click(message_id, deadline):
    """방금 보낸 버튼 메시지의 클릭을 기다림. 제한 시간 지나면 None"""
    while time.time() < deadline:
        remaining = deadline - time.time()
        try:
            updates = tg(
                "getUpdates",
                offset=state["offset"],
                timeout=int(min(30, max(1, remaining))),
                allowed_updates=["callback_query"],
            )
        except requests.RequestException as e:
            print(f"getUpdates 오류, 5초 후 재시도: {type(e).__name__}")
            time.sleep(5)
            continue

        for u in updates:
            state["offset"] = u["update_id"] + 1
            cq = u.get("callback_query")
            if not cq:
                continue
            msg = cq.get("message") or {}
            if str((msg.get("chat") or {}).get("id")) != CHAT_ID:
                continue
            if msg.get("message_id") != message_id:  # 예전 메시지의 버튼은 무시
                continue
            tg("answerCallbackQuery", callback_query_id=cq["id"])
            return cq["data"]
    return None


# =====================================================
# 2. 구글 시트
# =====================================================
def parse_date(value):
    s = str(value).strip().rstrip(".")
    for fmt in ("%Y-%m-%d", "%Y. %m. %d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def open_spreadsheet():
    info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    gc = gspread.service_account_from_dict(info)
    return gc.open_by_key(SHEET_ID)


def cell(row, i):
    return row[i].strip() if len(row) > i else ""


def get_today_complete_entries(sh, today):
    """날짜·시간·메뉴가 모두 채워진 오늘 행만"""
    rows = sh.worksheet(MEAL_SHEET_NAME).get_all_values()[1:]
    entries = []
    for row in rows:
        if parse_date(cell(row, 0)) == today and cell(row, 1) and cell(row, 2):
            entries.append(f"{cell(row, 1)} {cell(row, 2)}")
    return entries


def get_meal_entries(sh, today, days):
    start = today - timedelta(days=days - 1)
    rows = sh.worksheet(MEAL_SHEET_NAME).get_all_values()[1:]
    results = []
    for row in rows:
        d = parse_date(cell(row, 0))
        menu = cell(row, 2)
        if not d or not menu or d < start:
            continue
        results.append(f"{d.isoformat()} {cell(row, 1)} - {menu}")
    return results[-MAX_MEALS_FOR_ANALYSIS:]


def get_latest_ldl(sh):
    rows = sh.worksheet(CHECKUP_SHEET_NAME).get_all_values()[1:]
    latest = None
    for row in rows:
        if LDL_ITEM_KEYWORD not in cell(row, 1):
            continue
        d = parse_date(cell(row, 0))
        if not d:
            continue
        if latest is None or d > latest["date"]:
            latest = {"date": d, "value": cell(row, 2), "ref": cell(row, 3)}
    return latest


# =====================================================
# 3. Gemini 분석
# =====================================================
def build_prompt(meals, ldl):
    if ldl:
        ldl_text = (
            f"가장 최근 검진({ldl['date'].isoformat()}) LDL 콜레스테롤 수치는 "
            f"{ldl['value']}이고, 참고치는 {ldl['ref']}이다."
        )
    else:
        ldl_text = "LDL 콜레스테롤 검진 기록이 없다."

    return f"""너는 건강 코치야. 아래 정보를 보고 식습관이 LDL 콜레스테롤 관리에 어떤 영향을 줄 수 있는지 분석해줘.

{ldl_text}

최근 {ANALYSIS_DAYS}일 (오늘 포함) 동안의 식단 기록:
{chr(10).join(meals)}

분석 시 다음을 포함해줘:
1. 오늘과 최근 2주 식습관 패턴 요약
2. LDL 관점에서 주의할 점
3. 개선하면 좋을 점 1~2가지

텔레그램 메시지니까 800자 이내로 핵심만 정리해서 답해줘."""


def analyze_with_gemini(meals, ldl):
    error_message = "분석 중 오류가 발생했어요. 잠시 후 다시 시도해주세요."
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
    try:
        res = requests.post(
            url,
            headers={"x-goog-api-key": GEMINI_API_KEY},
            json={"contents": [{"parts": [{"text": build_prompt(meals, ldl)}]}]},
            timeout=60,
        )
        if res.status_code != 200:
            print(f"Gemini 오류 (상태코드 {res.status_code})")
            return error_message
        text = res.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
        return text or error_message
    except (requests.RequestException, KeyError, IndexError, ValueError) as e:
        print(f"Gemini 처리 실패: {type(e).__name__}")
        return error_message


def run_analysis(sh, today):
    meals = get_meal_entries(sh, today, ANALYSIS_DAYS)
    if not meals:
        send_text("해당 기간에 기록된 식단이 없어요.")
        return
    send_text(analyze_with_gemini(meals, get_latest_ldl(sh)))


# =====================================================
# 4. 메인 흐름
# =====================================================
def main():
    # 예전 Apps Script 웹훅이 걸려 있으면 getUpdates가 막히므로 해제
    tg("deleteWebhook", drop_pending_updates=True)

    today = datetime.now(KST).date()
    sh = open_spreadsheet()
    entries = get_today_complete_entries(sh, today)

    if entries:
        message = (
            "오늘 식단 확인할게요!\n\n"
            + "\n".join(entries)
            + f"\n\n이렇게 {len(entries)}끼 입력되어 있어요.\n오늘 식단 다 입력하신 거 맞나요??"
        )
    else:
        message = (
            "오늘 아직 완전하게 입력된 식단이 없어요 (날짜·시간·메뉴 다 채워져야 인정돼요).\n"
            "오늘 식단 입력이 없거나 더 남아있나요?"
        )

    deadline = time.time() + WAIT_MINUTES * 60

    mid = send_buttons(
        message,
        [[
            {"text": "1️⃣ 네, 맞아요", "callback_data": CONFIRM_YES},
            {"text": "2️⃣ 아니요, 더 입력할게요", "callback_data": CONFIRM_NO},
        ]],
    )
    click = wait_for_click(mid, deadline)

    if click is None:
        print("제한 시간 내 응답 없음 - 종료")
        return

    if click == CONFIRM_NO:
        send_text(f"여기서 입력하세요:\n{SHEET_URL}")
        return

    # CONFIRM_YES
    mid = send_buttons(
        "잘 했어요! 👍\n\nLDL 관점에서 식단 분석을 진행할까요?",
        [[
            {"text": "1️⃣ 네, 분석해주세요", "callback_data": ANALYZE_YES},
            {"text": "2️⃣ 괜찮아요", "callback_data": ANALYZE_NO},
        ]],
    )
    click = wait_for_click(mid, deadline)

    if click == ANALYZE_YES:
        send_text("오늘 + 최근 2주 기준으로 분석할게요...")
        run_analysis(sh, today)
    elif click == ANALYZE_NO:
        send_text("네 알겠습니다. 건강한 식단 잘 유지하시기 바랍니다. 💪")
    else:
        print("제한 시간 내 응답 없음 - 종료")


if __name__ == "__main__":
    main()
