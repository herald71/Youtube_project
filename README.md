## YouTube Project - Video Search & Excel Export Tool

이 프로젝트는 **YouTube Data API v3**를 사용해서  
검색어, 채널 ID, 날짜 범위를 기준으로 동영상을 검색하고,  
결과를 **엑셀 파일(.xlsx)** 로 저장해 주는 도구입니다.

주요 스크립트는 `Youtube_search.py` 로, 간단한 **GUI(Tkinter)** 를 통해  
검색 조건을 입력하고, 조회 결과를 바로 엑셀로 저장할 수 있습니다.

---

### Main Features

- **Flexible search options**
  - Search by **keyword** or **channel ID**
  - Optional **date range filter** (start / end date)
- **Rich video metadata collection**
  - Title, channel name, channel ID
  - Duration (normalized to `HH:MM:SS`)
  - View count, comment count
  - Tags, thumbnail URL, published date
- **Excel export**
  - Results are saved as an Excel file with a name like  
    `YOUR_FILE_NAME_YYYY-MM-DD.xlsx`

---

### Requirements

이 프로젝트는 Python 3 환경을 기준으로 하며,  
필요한 패키지는 `requirements.txt` 에 정의되어 있습니다.

```bash
pip install -r requirements.txt
```

주요 의존성:

- `pandas`
- `youtube-transcript-api`
- `openai`
- `python-dotenv`
- `openpyxl`
- `requests`
- `pyinstaller`
- `isodate`

---

### Environment Variables

API 키 등 민감한 정보는 **`.env` 파일**에 보관합니다.  
예시:

```env
YOUTUBE_API_KEY=YOUR_YOUTUBE_API_KEY_HERE
OPENAI_API_KEY=YOUR_OPENAI_API_KEY_HERE
```

`.env` 파일은 `.gitignore` 에 포함되어 있으므로  
버전 관리(GitHub)에는 업로드되지 않습니다.

---

### How to Run

1. Create and activate a virtual environment (optional but recommended).
2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Prepare your `.env` file with the necessary API keys.
4. Run the main script:

   ```bash
   python Youtube_search.py
   ```

5. Use the GUI to:
   - Enter search query or channel ID.
   - Optionally set date range.
   - Click the button to search and export results to Excel.

---

### Build (Executable)

이 프로젝트는 `pyinstaller` 를 사용해 **Windows 실행 파일(.exe)** 로 빌드할 수 있습니다.  
다만, `dist/`와 `build/` 폴더는 용량이 매우 커서 Git에는 포함하지 않습니다  
(`.gitignore` 로 무시 처리).

