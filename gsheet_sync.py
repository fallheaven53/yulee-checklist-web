"""
운영 체크리스트 — 구글 스프레드시트 동기화 모듈
로컬(tkinter)과 웹앱(Streamlit) 모두에서 사용
"""

import os
import time

try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GSPREAD = True
except ImportError:
    HAS_GSPREAD = False


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# 구글 시트 구조 (4개 시트)
# - "체크리스트항목": 단계 | 코드 | 항목명
# - "회차정보":       회차 | 공연일 | 출연단체 | 장르 | 공연시간 | 날씨 | 담당자
# - "회차별체크":     회차 | 코드 | 상태 | 완료시간 | 담당 | 메모
# - "운영총평":       회차 | 예상관객수 | 공연평가 | 총평 | 개선사항

MAX_RETRIES = 3


class GoogleSheetSync:
    """구글 스프레드시트 읽기/쓰기"""

    def __init__(self, credentials_path=None, credentials_dict=None,
                 spreadsheet_id=None):
        if not HAS_GSPREAD:
            raise ImportError("gspread 패키지가 필요합니다. pip install gspread")

        self.spreadsheet_id = spreadsheet_id

        if credentials_dict:
            creds = Credentials.from_service_account_info(
                credentials_dict, scopes=SCOPES)
        elif credentials_path and os.path.exists(credentials_path):
            creds = Credentials.from_service_account_file(
                credentials_path, scopes=SCOPES)
        else:
            raise FileNotFoundError("구글 서비스 계정 인증 정보가 없습니다.")

        self.service_email = credentials_dict.get("client_email", "") if credentials_dict else ""
        print(f"[GSheet] 인증 이메일: {self.service_email}")
        print(f"[GSheet] 스프레드시트 ID: {spreadsheet_id}")
        self.gc = gspread.authorize(creds)
        self.spreadsheet = self.gc.open_by_key(spreadsheet_id)
        print(f"[GSheet] 스프레드시트 연결 성공: {self.spreadsheet.title}")

    # ═══════════════════════════════════════════
    #  시트 헬퍼
    # ═══════════════════════════════════════════

    def _get_or_create_sheet(self, title, rows=200, cols=10):
        try:
            return self.spreadsheet.worksheet(title)
        except gspread.exceptions.WorksheetNotFound:
            return self.spreadsheet.add_worksheet(
                title=title, rows=rows, cols=cols)

    def _overwrite_sheet(self, ws, data):
        """clear() 없이 update()로 직접 덮어쓰기 (403 우회)"""
        if not data:
            return
        num_rows = len(data)
        num_cols = max(len(row) for row in data)
        end_col = chr(ord('A') + num_cols - 1)
        range_str = f"A1:{end_col}{num_rows}"

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                ws.update(range_notation=range_str, values=data,
                          value_input_option="RAW")
                cur_rows = ws.row_count
                if cur_rows > num_rows:
                    try:
                        ws.delete_rows(num_rows + 1, cur_rows)
                    except Exception:
                        pass
                return
            except Exception as e:
                print(f"[GSheet] {ws.title} 쓰기 실패 ({attempt}/{MAX_RETRIES}): {e}")
                if attempt < MAX_RETRIES:
                    time.sleep(1)
                else:
                    raise

    # ═══════════════════════════════════════════
    #  업로드 (ChecklistManager → 구글 시트)
    # ═══════════════════════════════════════════

    def upload_checklist(self, mgr):
        """ChecklistManager의 전체 데이터를 구글 시트에 업로드"""
        total_checks = sum(len(v) for v in mgr.checks.values())
        print(f"[업로드] 데이터: {len(mgr.round_info)}개 회차, {total_checks}건 체크")

        # 1) 체크리스트항목
        print("[업로드] 1/4 체크리스트항목...")
        ws_items = self._get_or_create_sheet("체크리스트항목")
        items_data = [["단계", "코드", "항목명"]]
        for stage, code, name in mgr.items:
            items_data.append([stage, code, name])
        self._overwrite_sheet(ws_items, items_data)

        # 2) 회차정보
        print("[업로드] 2/4 회차정보...")
        ws_info = self._get_or_create_sheet("회차정보")
        info_data = [["회차", "공연일", "출연단체", "장르", "공연시간", "날씨", "담당자"]]
        for rnd in sorted(mgr.round_info.keys()):
            info = mgr.round_info[rnd]
            info_data.append([
                rnd,
                info.get("공연일", ""),
                info.get("출연단체", ""),
                info.get("장르", ""),
                info.get("공연시간", ""),
                info.get("날씨", ""),
                info.get("담당자", ""),
            ])
        self._overwrite_sheet(ws_info, info_data)

        # 3) 회차별체크 (회차별 + 시즌 초)
        print("[업로드] 3/4 회차별체크...")
        ws_checks = self._get_or_create_sheet("회차별체크")
        checks_data = [["회차", "코드", "상태", "완료시간", "담당", "메모"]]
        for rnd in sorted(mgr.checks.keys()):
            for code, cd in mgr.checks[rnd].items():
                checks_data.append([
                    rnd, code,
                    cd.get("상태", "미완료"),
                    cd.get("완료시간", ""),
                    cd.get("담당", ""),
                    cd.get("메모", ""),
                ])
        from data_manager import CURRENT_SEASON
        for code, cd in mgr.season_checks.items():
            checks_data.append([
                CURRENT_SEASON, code,
                cd.get("상태", "미완료"),
                cd.get("완료시간", ""),
                cd.get("담당", ""),
                cd.get("메모", ""),
            ])
        self._overwrite_sheet(ws_checks, checks_data)

        # 4) 운영총평
        print("[업로드] 4/4 운영총평...")
        ws_reviews = self._get_or_create_sheet("운영총평")
        reviews_data = [["회차", "예상관객수", "공연평가", "총평", "개선사항"]]
        for rnd in sorted(mgr.reviews.keys()):
            rv = mgr.reviews[rnd]
            reviews_data.append([
                rnd,
                rv.get("예상관객수", ""),
                rv.get("공연평가", ""),
                rv.get("총평", ""),
                rv.get("개선사항", ""),
            ])
        self._overwrite_sheet(ws_reviews, reviews_data)

        print("[업로드] 전체 완료")

    # ═══════════════════════════════════════════
    #  다운로드 (구글 시트 → ChecklistManager)
    # ═══════════════════════════════════════════

    def download_checklist(self, mgr):
        """구글 시트에서 전체 데이터를 ChecklistManager로 로드"""

        # 1) 체크리스트항목 — 항상 코드(DEFAULT_ITEMS) 기준 사용
        from data_manager import DEFAULT_ITEMS
        mgr.items = list(DEFAULT_ITEMS)

        # 2) 회차정보
        try:
            ws = self.spreadsheet.worksheet("회차정보")
            rows = ws.get_all_values()
            mgr.round_info = {}
            for row in rows[1:]:
                if not row or not row[0].strip():
                    continue
                try:
                    rnd = int(row[0])
                except ValueError:
                    continue
                mgr.round_info[rnd] = {
                    "공연일":   row[1].strip() if len(row) > 1 else "",
                    "출연단체": row[2].strip() if len(row) > 2 else "",
                    "장르":     row[3].strip() if len(row) > 3 else "",
                    "공연시간": row[4].strip() if len(row) > 4 else "",
                    "날씨":     row[5].strip() if len(row) > 5 else "",
                    "담당자":   row[6].strip() if len(row) > 6 else "",
                }
        except gspread.exceptions.WorksheetNotFound:
            pass

        # 3) 회차별체크 — "2026시즌" 등 문자열 회차는 시즌 초 항목으로 분류
        try:
            ws = self.spreadsheet.worksheet("회차별체크")
            rows = ws.get_all_values()
            mgr.checks = {}
            mgr.season_checks = {}
            for row in rows[1:]:
                if not row or not row[0].strip():
                    continue
                rnd_str = row[0].strip()
                code = row[1].strip() if len(row) > 1 else ""
                if not code:
                    continue
                check_data = {
                    "상태":     row[2].strip() if len(row) > 2 else "미완료",
                    "완료시간": row[3].strip() if len(row) > 3 else "",
                    "담당":     row[4].strip() if len(row) > 4 else "",
                    "메모":     row[5].strip() if len(row) > 5 else "",
                }
                try:
                    rnd = int(rnd_str)
                    if rnd not in mgr.checks:
                        mgr.checks[rnd] = {}
                    mgr.checks[rnd][code] = check_data
                except ValueError:
                    mgr.season_checks[code] = check_data
        except gspread.exceptions.WorksheetNotFound:
            pass

        # 4) 운영총평
        try:
            ws = self.spreadsheet.worksheet("운영총평")
            rows = ws.get_all_values()
            mgr.reviews = {}
            for row in rows[1:]:
                if not row or not row[0].strip():
                    continue
                try:
                    rnd = int(row[0])
                except ValueError:
                    continue
                mgr.reviews[rnd] = {
                    "예상관객수": row[1].strip() if len(row) > 1 else "",
                    "공연평가":   row[2].strip() if len(row) > 2 else "",
                    "총평":       row[3].strip() if len(row) > 3 else "",
                    "개선사항":   row[4].strip() if len(row) > 4 else "",
                }
        except gspread.exceptions.WorksheetNotFound:
            pass

        # 진단: 로드된 데이터 요약
        for rnd in sorted(mgr.checks.keys()):
            codes = list(mgr.checks[rnd].keys())
            print(f"[GSheet] {rnd}회차 체크 로드: {len(codes)}건 ({codes[:3]}...)")
