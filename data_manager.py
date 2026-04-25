"""
운영 체크리스트 — 데이터 매니저
구글 스프레드시트 기반 읽기/쓰기
"""

import os
from datetime import datetime

# ── 기본 체크리스트 항목 ──
DEFAULT_ITEMS = [
    ("A", "A-01", "출연단체 출연 확인 연락 완료"),
    ("A", "A-02", "출연단체 셋리스트 변경사항 확인"),
    ("A", "A-03", "사회자(MC) 대본 작성·전달"),
    ("A", "A-04", "홍보물(SNS·현수막 등) 게시 완료"),
    ("A", "A-05", "음향사 당일 세팅 사전 확인"),
    ("A", "A-06", "SNS·홈페이지 공연 안내 게시 확인"),
    ("A", "A-07", "이해관계자 대상 ESG 교육(공지) 확인"),
    ("A", "A-08", "인권보호·성희롱·성폭력 방지 서약서 징구 확인"),

    ("B", "B-01", "서석당 무대 세팅 (현수막 게첨, 의자 배치)"),
    ("B", "B-02", "음향 장비 설치·테스트 및 밸런스 최종 확인"),
    ("B", "B-03", "빔프로젝터·자막 송출 테스트 및 최종 확인"),
    ("B", "B-04", "체험부스 세팅 확인"),
    ("B", "B-05", "다과·음료 준비 (일회용품 절감 확인 포함)"),
    ("B", "B-06", "안내 표지판·동선 확인"),
    ("B", "B-07", "출연단체 도착·대기실 안내 (편의시설 점검 포함)"),

    ("C", "C-01", "출연단체 리허설 완료"),
    ("C", "C-02", "사회자 동선·큐시트 최종 확인"),

    ("D", "D-01", "공연 정시 시작"),
    ("D", "D-02", "현장 사진·영상 촬영"),
    ("D", "D-03", "관객 안전 관리 이상 없음"),
    ("D", "D-04", "만족도조사 QR코드 안내 실시"),
    ("D", "D-05", "취약계층 관람 편의 확인"),
    ("D", "D-06", "저작권 침해 방지 확인 (사진·영상 기록 전 참여자 동의)"),

    ("E", "E-01", "무대·객석 철거 완료"),
    ("E", "E-02", "음향 장비 상태 점검 완료"),
    ("E", "E-03", "출연단체 출연료 지급 서류 처리"),
    ("E", "E-04", "공연 사진·영상·고정카메라 아카이빙"),
    ("E", "E-05", "체험부스 운영 완료 확인"),
    ("E", "E-06", "SNS 공연 후기 게시"),
    ("E", "E-07", "폐기물 분리배출·정리 확인"),
    ("E", "E-08", "에너지(조명·냉난방) 정상 종료 확인"),
    ("E", "E-09", "운영 ESG 점검 메모"),
    ("E", "E-10", "특이사항 기록 완료"),

    ("S", "S-01", "자동 발송 알림 확인"),
    ("S", "S-02", "발송로그 존재 확인"),
    ("S", "S-03", "GCP 콘솔 확인"),
    ("S", "S-04", "주간 발송 현황 확인"),
    ("S", "S-05", "실패 건 원인 확인"),
    ("S", "S-06", "다음 주 공연 일정 확인"),
    ("S", "S-07", "월간 헬스체크 결과 확인"),
    ("S", "S-08", "월간 발송 성공률 확인"),
    ("S", "S-09", "5일 전 안내 문자 자동 발송"),
    ("S", "S-10", "1일 전 리마인드 자동 발송"),
]

SEASON_ITEMS = {"A-07", "A-08"}
CURRENT_SEASON = "2026시즌"

STAGE_LABELS = {
    "A": "A. 공연 전 — 사전 준비",
    "B": "B. 공연 당일 — 현장 세팅",
    "C": "C. 공연 당일 — 리허설",
    "D": "D. 공연 중",
    "E": "E. 공연 후 — 마무리",
    "S": "S. SMS 자동발송 모니터링",
}

GENRE_LIST = ["전통음악", "국악", "민요", "창작국악", "판소리", "정가", "기악", "무용", "혼합장르", "기타"]
WEATHER_LIST = ["맑음", "구름조금", "흐림", "비", "눈", "안개", "강풍"]
EVAL_LIST = ["매우좋음", "좋음", "보통", "미흡", "매우미흡"]


class ChecklistManager:
    """운영 체크리스트 데이터 관리 (구글 시트 연동)"""

    def __init__(self, gsheet_sync=None):
        self.gsheet = gsheet_sync
        self.items = list(DEFAULT_ITEMS)
        self.round_info = {}   # {회차(int): {공연일, 출연단체, 장르, ...}}
        self.checks = {}       # {회차(int): {코드: {상태, 완료시간, 담당, 메모}}}
        self.reviews = {}      # {회차(int): {예상관객수, 공연평가, 총평, 개선사항}}
        self.season_checks = {}  # {코드: {상태, 완료시간, 담당, 메모}} — 시즌 초 1회성
        self.last_save_error = None
        self._loaded_ok = False
        self.load()

    @property
    def round_items(self):
        return [(s, c, n) for s, c, n in self.items if c not in SEASON_ITEMS]

    @property
    def season_item_list(self):
        return [(s, c, n) for s, c, n in self.items if c in SEASON_ITEMS]

    # ── 구글 시트에서 로드 ──
    def load(self):
        if not self.gsheet:
            return
        try:
            self.gsheet.download_checklist(self)
            self._loaded_ok = True
        except Exception as e:
            self._loaded_ok = False
            print(f"[구글시트 로드 실패] {e}")

    # ── 구글 시트에 저장 ──
    def save(self):
        if not self.gsheet:
            return
        if not self._loaded_ok:
            print("[구글시트 저장 스킵] 로드 실패 상태에서 저장하면 데이터 유실 위험")
            self.last_save_error = "초기 로드 실패 — 저장 차단 (데이터 보호)"
            return
        if not self.round_info and not self.checks and not self.season_checks:
            print("[구글시트 저장 스킵] 빈 데이터 — 시트 덮어쓰기 방지")
            return
        try:
            self.gsheet.upload_checklist(self)
            self.last_save_error = None
        except Exception as e:
            self.last_save_error = str(e)
            print(f"[구글시트 저장 실패] {e}")

    # ── 회차 관리 ──
    @property
    def round_list(self):
        return sorted(self.round_info.keys())

    def add_round(self, rnd, info):
        self.round_info[rnd] = info
        if rnd not in self.checks:
            self.checks[rnd] = {}
        if rnd not in self.reviews:
            self.reviews[rnd] = {}
        self.save()

    def delete_round(self, rnd):
        self.round_info.pop(rnd, None)
        self.checks.pop(rnd, None)
        self.reviews.pop(rnd, None)
        self.save()

    # ── 체크 동작 ──
    def set_check(self, rnd, code, status, staff="", memo=""):
        if rnd not in self.checks:
            self.checks[rnd] = {}
        now = datetime.now().strftime("%-m/%-d, %H:%M") if status == "완료" else ""
        self.checks[rnd][code] = {
            "상태": status,
            "완료시간": now,
            "담당": staff,
            "메모": memo,
        }

    def get_check(self, rnd, code):
        return self.checks.get(rnd, {}).get(code, {
            "상태": "미완료", "완료시간": "", "담당": "", "메모": ""
        })

    def save_checks(self, rnd, checks_data, review_data=None, season_data=None):
        """회차 전체 체크 데이터 저장"""
        self.checks[rnd] = checks_data
        if review_data is not None:
            self.reviews[rnd] = review_data
        if season_data is not None:
            self.season_checks = season_data
        self.save()

    def copy_prev_checks(self, rnd):
        """이전 회차 체크를 현재 회차에 복사"""
        prev = rnd - 1
        if prev not in self.checks:
            return False
        if rnd not in self.checks:
            self.checks[rnd] = {}
        for code, cd in self.checks[prev].items():
            status = cd["상태"]
            if status == "미완료" and (cd.get("담당") or cd.get("메모")):
                status = "완료"
            self.checks[rnd][code] = {
                "상태": status,
                "완료시간": cd["완료시간"],
                "담당": cd["담당"],
                "메모": cd["메모"],
            }
        self.save()
        return True

    def copy_prev_stage(self, rnd, stage):
        """이전 회차의 특정 섹션만 현재 회차에 복사"""
        prev = rnd - 1
        if prev not in self.checks:
            return False
        stage_codes = {c for s, c, _ in self.items if s == stage}
        if not stage_codes:
            return False
        if rnd not in self.checks:
            self.checks[rnd] = {}
        copied = False
        for code, cd in self.checks[prev].items():
            if code in stage_codes:
                status = cd["상태"]
                if status == "미완료" and (cd.get("담당") or cd.get("메모")):
                    status = "완료"
                self.checks[rnd][code] = {
                    "상태": status,
                    "완료시간": cd["완료시간"],
                    "담당": cd["담당"],
                    "메모": cd["메모"],
                }
                copied = True
        if copied:
            self.save()
        return copied

    def reset_checks(self, rnd):
        """현재 회차 체크 초기화"""
        self.checks[rnd] = {}
        self.save()

    # ── 통계 ──
    def get_round_rate(self, rnd):
        if rnd not in self.checks:
            return 0.0
        items = self.round_items
        total = len(items)
        if total == 0:
            return 0.0
        done = sum(
            1 for _, code, _ in items
            if self.checks[rnd].get(code, {}).get("상태") in ("완료", "해당없음")
        )
        return round(done / total * 100, 1)

    def get_stage_rate(self, rnd, stage):
        stage_items = [(c, n) for s, c, n in self.round_items if s == stage]
        if not stage_items:
            return 0.0
        done = sum(
            1 for code, _ in stage_items
            if self.checks.get(rnd, {}).get(code, {}).get("상태") in ("완료", "해당없음")
        )
        return round(done / len(stage_items) * 100, 1)

    def get_round_status(self, rnd):
        if rnd not in self.checks or not self.checks[rnd]:
            return "미착수"
        round_codes = {c for _, c, _ in self.round_items}
        statuses = [cd.get("상태", "미완료")
                     for code, cd in self.checks[rnd].items() if code in round_codes]
        if all(s in ("완료", "해당없음") for s in statuses) and len(statuses) == len(round_codes):
            return "완료"
        return "진행중"

    def get_item_stats(self):
        total_rounds = len(self.round_info)
        stats = {}
        for _, code, _ in self.items:
            stats[code] = {"완료수": 0, "미완료수": 0, "해당없음수": 0}
        for rnd, codes in self.checks.items():
            for code, cd in codes.items():
                if code not in stats:
                    continue
                s = cd.get("상태", "미완료")
                if s == "완료":
                    stats[code]["완료수"] += 1
                elif s == "해당없음":
                    stats[code]["해당없음수"] += 1
                else:
                    stats[code]["미완료수"] += 1
        for code, cd in self.season_checks.items():
            if code not in stats:
                continue
            s = cd.get("상태", "미완료")
            if s == "완료":
                stats[code]["완료수"] = 1
            elif s == "해당없음":
                stats[code]["해당없음수"] = 1
            else:
                stats[code]["미완료수"] = 1
        return stats, total_rounds

    # ── 항목 관리 ──
    def add_item(self, stage, code, name):
        self.items.append((stage, code, name))
        self.save()

    def update_item(self, old_code, stage, code, name):
        for i, (s, c, n) in enumerate(self.items):
            if c == old_code:
                self.items[i] = (stage, code, name)
                break
        if old_code != code:
            for rnd in self.checks:
                if old_code in self.checks[rnd]:
                    self.checks[rnd][code] = self.checks[rnd].pop(old_code)
        self.save()

    def delete_item(self, code):
        self.items = [(s, c, n) for s, c, n in self.items if c != code]
        for rnd in self.checks:
            self.checks[rnd].pop(code, None)
        self.save()

    def gen_code(self, stage):
        existing = [int(c.split("-")[1]) for s, c, _ in self.items if s == stage]
        next_num = max(existing, default=0) + 1
        return f"{stage}-{next_num:02d}"
