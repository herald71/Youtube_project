"""
유튜브 검색/채널/날짜(연/월/일) 정보를 입력 받아
해당 조건에 맞는 동영상 목록을 유튜브 데이터 API로 조회한 뒤
엑셀 파일로 저장하는 간단한 수집 스크립트입니다.

주요 기능
- 검색어 또는 채널 ID, 날짜 범위(옵션)를 기준으로 동영상 검색
- 각 동영상의 제목, 채널명, 채널 ID, 재생시간, 조회수, 댓글수, 태그, 썸네일, 업로드 날짜 수집
- 수집한 결과를 `입력한_파일명_오늘날짜.xlsx` 형태의 엑셀 파일로 저장
"""

# 검색어, 채널 아이디, 날짜 정보까지 입력 받아서 결과를 출력합니다.
# 활용을 잘 하면 유튜브 컨텐츠 분석/기획 등에 유용하게 사용할 수 있습니다.

import os

from dotenv import load_dotenv
import requests
import pandas as pd
from datetime import datetime
import isodate
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

# .env 파일에 정의된 환경 변수를 읽어오기 위한 설정
# (예: YOUTUBE_API_KEY 등)
load_dotenv()

def format_duration(duration):
    """
    유튜브 API에서 내려주는 ISO 8601 형식의 재생시간(duration)을
    사람이 보기 쉬운 "HH:MM:SS" 형태의 문자열로 변환하는 함수입니다.

    예) 'PT1H2M3S' -> '01:02:03'
    """
    # ISO 8601 형식의 문자열을 timedelta 형태로 파싱
    parsed_duration = isodate.parse_duration(duration)
    # 총 초(second)를 기준으로 3600초(1시간) 단위로 나누어 시/분/초 계산
    hours, remainder = divmod(parsed_duration.total_seconds(), 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}"

def fetch_youtube_data(api_key, search_query, channel_id, start_date=None, end_date=None, log_callback=None):
    """
    유튜브 데이터 API(v3)를 이용하여 동영상 검색 및 상세 정보를 조회하는 함수입니다.

    매개변수
    - api_key: 유튜브 데이터 API 키
    - search_query: 검색어(빈 문자열이면 채널 ID 기준 검색)
    - channel_id: 채널 ID(빈 문자열이면 검색어 기준 검색)
    - start_date: 조회 시작 날짜(YYYY-MM-DD, None 또는 빈 문자열이면 제한 없음)
    - end_date: 조회 종료 날짜(YYYY-MM-DD, None 또는 빈 문자열이면 제한 없음)

    반환값
    - 각 동영상의 주요 정보(제목, 채널명, 재생시간, 조회수 등)를 담은 딕셔너리의 리스트
    """
    # GUI 모드에서도 진행 상황을 볼 수 있도록 로그 콜백을 활용합니다.
    # log_callback이 지정되지 않은 경우 기본적으로 print를 사용합니다.
    if log_callback is None:
        log_callback = print

    log_callback("검색 조건으로 유튜브 검색을 시작합니다...")

    # 검색 요청의 기본 URL (검색어/채널/날짜 범위에 따라 파라미터를 추가하는 방식)
    base_search_url = f'https://www.googleapis.com/youtube/v3/search?key={api_key}&part=snippet&maxResults=50&type=video&order=date'

    # 검색어가 있으면 q 파라미터로 추가
    if search_query:
        base_search_url += f'&q={search_query}'

    # 채널 ID가 있으면 channelId 파라미터로 추가
    if channel_id:
        base_search_url += f'&channelId={channel_id}'

    # 시작/종료 날짜가 있으면 해당 기간으로 검색 기간을 제한
    # (YYYY-MM-DD 형식의 문자열을 받아, RFC3339 형식으로 변환하여 사용)
    if start_date:
        try:
            start_dt = datetime.strptime(start_date, "%Y-%m-%d")
            start_iso = start_dt.strftime("%Y-%m-%dT00:00:00Z")
            base_search_url += f"&publishedAfter={start_iso}"
        except ValueError:
            # 형식이 잘못된 경우에는 로그만 남기고 필터는 적용하지 않음
            log_callback(f"[경고] 시작 날짜 형식이 올바르지 않습니다(YYYY-MM-DD): {start_date}")

    if end_date:
        try:
            end_dt = datetime.strptime(end_date, "%Y-%m-%d")
            end_iso = end_dt.strftime("%Y-%m-%dT23:59:59Z")
            base_search_url += f"&publishedBefore={end_iso}"
        except ValueError:
            log_callback(f"[경고] 종료 날짜 형식이 올바르지 않습니다(YYYY-MM-DD): {end_date}")

    # 각 동영상의 상세 정보를 담을 리스트
    video_details = []
    # 지금까지 가져온 총 결과 개수 (페이지네이션 시 Index 컬럼에 사용)
    total_results_fetched = 0
    # 다음 페이지를 요청하기 위한 토큰 (첫 요청 시에는 None)
    next_page_token = None

    while True:
        # 매 반복마다 기본 검색 URL을 기준으로 새로 URL을 구성
        search_url = base_search_url

        # 다음 페이지 토큰이 있으면 pageToken 파라미터로 추가
        if next_page_token:
            search_url += f'&pageToken={next_page_token}'

        try:
            # 검색 API 호출
            log_callback("검색 API를 호출하는 중입니다...")
            search_response = requests.get(search_url, timeout=10)
            search_response.raise_for_status()
        except requests.exceptions.RequestException as e:
            # 네트워크 오류, 타임아웃 등 각종 예외를 로그로 남기고 반복을 종료합니다.
            log_callback(f"[에러] 검색 결과를 가져오지 못했습니다: {e}")
            break

        search_results = search_response.json()
        # 다음 페이지 호출을 위한 토큰 (마지막 페이지이면 존재하지 않을 수 있음)
        next_page_token = search_results.get('nextPageToken')
        
        # 검색 결과(`items`)를 순회하면서 각 동영상의 상세 정보를 다시 조회
        for index, item in enumerate(search_results.get('items', []), start=1):
            # 검색 결과 항목에서 동영상 ID 추출
            video_id = item['id']['videoId']
            video_url = f"https://www.youtube.com/watch?v={video_id}"
            log_callback(f"{index}번째 동영상 상세 정보 조회 중: {video_url}")

            # 동영상 상세 정보(재생시간, 통계, 제목 등)를 가져오는 API URL
            video_details_url = f'https://www.googleapis.com/youtube/v3/videos?id={video_id}&key={api_key}&part=contentDetails,statistics,snippet'
            try:
                # 동영상 상세 정보 API 호출
                video_details_response = requests.get(video_details_url, timeout=10)
                video_details_response.raise_for_status()
            except requests.exceptions.RequestException as e:
                # 개별 동영상 정보 조회 실패 시 해당 동영상만 건너뛰고 계속 진행합니다.
                log_callback(f"[경고] {index}번째 동영상 상세 정보를 가져오지 못했습니다: {e}")
                continue

            video_info = video_details_response.json()
            # items가 비어 있지 않을 때에만 처리
            if 'items' in video_info and video_info['items']:
                video_data = video_info['items'][0]
                # 동영상 제목
                title = video_data['snippet']['title']
                # 채널명
                channel_title = video_data['snippet']['channelTitle']
                # 채널 ID (각 동영상이 어떤 채널에 속하는지 식별하기 위한 값)
                channel_id_value = video_data['snippet'].get('channelId', '')
                # ISO 8601 재생시간을 "HH:MM:SS"로 변환
                duration = format_duration(video_data['contentDetails']['duration'])
                # 조회수(없을 수 있어서 기본값 0으로 처리)
                views = int(video_data['statistics'].get('viewCount', 0))
                # 댓글수(없을 수 있어서 기본값 0으로 처리)
                comments = int(video_data['statistics'].get('commentCount', 0))
                # 태그 목록 (없을 경우 빈 리스트)
                tags = video_data['snippet'].get('tags', [])
                # 썸네일 URL (high 해상도 기준)
                thumbnail_url = video_data['snippet']['thumbnails']['high']['url']
                # 업로드 날짜 (문자열 -> date 객체로 변환)
                published_date = datetime.strptime(video_data['snippet']['publishedAt'], '%Y-%m-%dT%H:%M:%SZ').date()

                # 하나의 동영상에 대한 정보를 딕셔너리로 정리해서 리스트에 추가
                video_details.append({
                    # 전체 검색 결과 기준 Index (페이지가 넘어가도 이어지도록 보정)
                    'Index': index + total_results_fetched,
                    'Title': title,
                    'Channel Title': channel_title,
                    # 채널 ID도 함께 엑셀에 저장되도록 컬럼 추가
                    'Channel ID': channel_id_value,
                    'Duration': duration,
                    'Views': views,
                    'Comments': comments,
                    'URL': video_url,
                    'Thumbnail URL': thumbnail_url,
                    'Tags': ', '.join(tags),
                    'Published Date': published_date
                })

        # 이번 페이지에서 가져온 동영상 개수를 누적
        total_results_fetched += len(search_results.get('items', []))

        # 더 이상 다음 페이지가 없거나, 임의로 정한 최대 개수(200개) 이상이면 반복 종료
        if not next_page_token or total_results_fetched >= 200:
            break

    return video_details


def run_gui():
    """
    간단한 Tkinter 기반 GUI를 통해
    - 검색어 / 채널 ID
    - 조회 기간(시작 / 종료 날짜)
    - 엑셀 파일명
    을 입력 받고,
    환경 변수(또는 .env)에 저장된 API 키를 사용하여
    버튼 클릭으로 유튜브 데이터를 조회하고
    엑셀로 저장하는 함수입니다.

    또한 하단 로그 창에서 진행 상황과 오류 메시지를
    함께 확인할 수 있습니다.
    """

    # ------------------------- 이벤트 핸들러 함수 정의 -------------------------

    def log_message(message: str):
        """
        진행 상황이나 오류 메시지를 하단 텍스트 박스에 출력하는 함수입니다.
        """
        status_text.configure(state="normal")
        status_text.insert(tk.END, message + "\n")
        status_text.see(tk.END)
        status_text.configure(state="disabled")
        # 즉시 화면에 반영되도록 업데이트
        root.update_idletasks()

    def on_run_button_click():
        """
        '검색 및 저장' 버튼 클릭 시 호출되는 콜백 함수입니다.
        입력값을 검증한 뒤, 유튜브 API를 호출하고 엑셀로 저장합니다.
        """
        # 버튼 중복 클릭 방지를 위해 비활성화
        run_button.config(state="disabled")
        # 이전 로그를 지우고 새 로그 출력 준비
        status_text.configure(state="normal")
        status_text.delete("1.0", tk.END)
        status_text.configure(state="disabled")

        log_message("설정 값을 확인하는 중입니다...")

        # .env 또는 시스템 환경 변수에서 API 키를 읽어오기
        api_key = os.getenv("YOUTUBE_API_KEY", "").strip()

        # 각 입력 필드에서 값 읽어오기
        search_query = entry_search_query.get().strip()
        channel_id = entry_channel_id.get().strip()
        start_date_str = entry_start_date.get().strip()
        end_date_str = entry_end_date.get().strip()
        file_name = entry_file_name.get().strip()

        # ------------------------- 기본 유효성 검증 -------------------------
        # API 키는 반드시 환경 변수에 설정되어 있어야 함
        if not api_key:
            messagebox.showerror(
                "환경 변수 오류",
                "YOUTUBE_API_KEY 환경 변수가 설정되어 있지 않습니다.\n"
                ".env 파일 또는 시스템 환경 변수에 YOUTUBE_API_KEY를 설정해 주세요.",
            )
            run_button.config(state="normal")
            return

        # 검색어와 채널 ID가 모두 비어 있으면 안 됨
        if not search_query and not channel_id:
            messagebox.showerror(
                "입력 오류",
                "검색어 또는 채널 ID 중 하나 이상은 반드시 입력해야 합니다.",
            )
            run_button.config(state="normal")
            return

        # 시작/종료 날짜(YYYY-MM-DD)는 선택 사항이지만,
        # 값이 있다면 올바른 형식인지 검증합니다.
        for label, value in [("시작 날짜", start_date_str), ("종료 날짜", end_date_str)]:
            if value:
                try:
                    datetime.strptime(value, "%Y-%m-%d")
                except ValueError:
                    messagebox.showerror(
                        "입력 오류",
                        f"{label}는 YYYY-MM-DD 형식으로 입력해 주세요.\n예: 2024-01-31",
                    )
                    run_button.config(state="normal")
                    return

        # 파일명이 비어 있으면 기본 파일명으로 설정
        if not file_name:
            file_name = "youtube_results"

        log_message("유튜브 데이터 조회를 시작합니다...")

        try:
            # ------------------------- 유튜브 데이터 조회 -------------------------
            # 기존에 정의된 fetch_youtube_data 함수를 호출하여 결과를 가져옵니다.
            videos = fetch_youtube_data(
                api_key,
                search_query,
                channel_id,
                start_date_str or None,
                end_date_str or None,
                log_callback=log_message,
            )

            # 수집된 동영상 정보가 하나라도 있는 경우에만 엑셀로 저장
            if videos:
                # 리스트 형태의 동영상 정보를 판다스 DataFrame으로 변환
                df = pd.DataFrame(videos)
                # 오늘 날짜를 파일명에 포함 (예: 2026-02-11)
                today_date = datetime.now().strftime("%Y-%m-%d")
                # 최종 엑셀 파일 경로/이름 구성
                file_path = f"{file_name}_{today_date}.xlsx"
                # 인덱스 컬럼 없이 저장
                df.to_excel(file_path, index=False)

                log_message(f"엑셀 파일 저장 완료: {file_path}")

                # 저장 완료 후 엑셀(또는 기본 연결 프로그램)로 파일을 자동으로 엽니다.
                # - Windows: os.startfile 사용
                # - macOS/Linux: 시스템 기본 열기 명령 사용
                try:
                    absolute_path = os.path.abspath(file_path)
                    if os.name == "nt":
                        os.startfile(absolute_path)  # type: ignore[attr-defined]
                    else:
                        import subprocess

                        opener = "open" if os.name == "posix" else None
                        # posix 계열에서 macOS는 open, Linux는 xdg-open을 우선 사용
                        if opener == "open":
                            subprocess.run(["open", absolute_path], check=False)
                        else:
                            subprocess.run(["xdg-open", absolute_path], check=False)

                    log_message("엑셀 파일을 자동으로 열었습니다.")
                except Exception as e:
                    log_message(f"[경고] 엑셀 파일 자동 열기에 실패했습니다: {e}")

                # 성공 메시지 팝업
                messagebox.showinfo(
                    "완료",
                    f"데이터를 엑셀 파일로 저장했습니다.\n\n파일명: {file_path}",
                )
            else:
                # 검색 결과가 하나도 없는 경우
                log_message("검색 결과가 없어 엑셀 파일을 생성하지 않습니다.")
                messagebox.showinfo("알림", "검색 결과가 없습니다.")

        except Exception as e:
            # 예외 발생 시 에러 메시지를 로그와 팝업으로 표시
            log_message(f"[에러] 데이터 조회 또는 저장 중 오류가 발생했습니다: {e}")
            messagebox.showerror("에러", f"데이터 조회 또는 저장 중 오류가 발생했습니다.\n\n{e}")
        finally:
            # 처리 완료 후 버튼 다시 활성화
            run_button.config(state="normal")

    # ------------------------- 메인 윈도우 생성 -------------------------
    root = tk.Tk()
    root.title("YouTube 검색 엑셀 저장 도구")

    # 윈도우 기본 여백 설정을 위한 프레임
    main_frame = ttk.Frame(root, padding="16 16 16 16")
    main_frame.grid(row=0, column=0, sticky="nsew")

    # 윈도우 리사이즈 시 프레임이 함께 늘어나도록 설정
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    # ------------------------- 각 입력 위젯 생성 -------------------------

    # 검색어
    label_search_query = ttk.Label(main_frame, text="검색어 (선택)")
    label_search_query.grid(row=0, column=0, sticky="w", pady=4)
    entry_search_query = ttk.Entry(main_frame, width=50)
    entry_search_query.grid(row=0, column=1, sticky="ew", pady=4)

    # 채널 ID
    label_channel_id = ttk.Label(main_frame, text="채널 ID (선택)")
    label_channel_id.grid(row=1, column=0, sticky="w", pady=4)
    entry_channel_id = ttk.Entry(main_frame, width=50)
    entry_channel_id.grid(row=1, column=1, sticky="ew", pady=4)

    # 검색 조건 안내 문구
    label_hint = ttk.Label(
        main_frame,
        text="※ 검색어와 채널 ID 중 하나 이상은 반드시 입력해야 합니다.",
        foreground="gray",
    )
    label_hint.grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 8))

    # 시작 날짜
    label_start_date = ttk.Label(main_frame, text="시작 날짜 (YYYY-MM-DD, 선택)")
    label_start_date.grid(row=3, column=0, sticky="w", pady=4)
    entry_start_date = ttk.Entry(main_frame, width=20)
    entry_start_date.grid(row=3, column=1, sticky="w", pady=4)

    # 종료 날짜
    label_end_date = ttk.Label(main_frame, text="종료 날짜 (YYYY-MM-DD, 선택)")
    label_end_date.grid(row=4, column=0, sticky="w", pady=4)
    entry_end_date = ttk.Entry(main_frame, width=20)
    entry_end_date.grid(row=4, column=1, sticky="w", pady=4)

    # 파일명
    label_file_name = ttk.Label(main_frame, text="엑셀 파일명 (확장자 제외)")
    label_file_name.grid(row=5, column=0, sticky="w", pady=4)
    entry_file_name = ttk.Entry(main_frame, width=50)
    entry_file_name.grid(row=5, column=1, sticky="ew", pady=4)

    # 진행 상황 로그 영역
    label_status = ttk.Label(main_frame, text="진행 상황 / 로그")
    label_status.grid(row=6, column=0, sticky="nw", pady=(8, 4))

    status_text = scrolledtext.ScrolledText(main_frame, height=10, state="disabled")
    status_text.grid(row=6, column=1, sticky="nsew", pady=(8, 4))

    # 컬럼 비율 조정 (라벨은 고정, 입력창은 가변)
    main_frame.columnconfigure(0, weight=0)
    main_frame.columnconfigure(1, weight=1)
    # 로그 텍스트 박스가 세로로도 넓게 확장되도록 설정
    main_frame.rowconfigure(6, weight=1)

    # 실행 버튼
    run_button = ttk.Button(main_frame, text="검색 및 엑셀 저장", command=on_run_button_click)
    run_button.grid(row=7, column=0, columnspan=2, pady=(12, 0))

    # 엔터 키로도 버튼이 눌리도록 바인딩
    root.bind("<Return>", lambda event: on_run_button_click())

    # 메인 이벤트 루프 시작
    root.mainloop()


if __name__ == "__main__":
    # 스크립트를 직접 실행했을 때 GUI 모드로 실행되도록 합니다.
    run_gui()
