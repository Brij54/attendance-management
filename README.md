# Attendance Punch-In Automation Web App

## Stack
- FastAPI backend
- Streamlit frontend
- JWT coordinator login
- msoffcrypto-tool for Excel files that ask for an opening password
- openpyxl for Excel processing

## Workflow
1. Coordinator opens the Streamlit web page.
2. Coordinator signs in.
3. Uploads Attendance Excel and its password.
4. Uploads Swipe Excel and its password.
5. Clicks Generate Updated Excel.
6. Backend matches Employee Code + date.
7. Earliest swipe on each employee/date is written into a new Punchin Time column immediately after that date.
8. If no swipe exists, the new cell is highlighted in the punch-in time column (indicating absence).
9. Updated workbook is downloaded.
10. Temporary uploaded/decrypted files are deleted after the response.

## Local Windows setup

Install Python 3.12+.

Copy `.env.example` to `.env` and change:
- JWT_SECRET
- COORDINATOR_USERNAME
- COORDINATOR_PASSWORD

### Backend
Open Terminal 1:
```powershell
cd backend
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

### Frontend
Open Terminal 2:
```powershell
cd frontend
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python -m streamlit run app.py --server.port 8502
```

Open http://localhost:8502

## Docker

Copy `.env.example` to `.env`, edit the secrets, then:
`docker compose up --build -d`

Open http://localhost:8502

Stop:
`docker compose down`

## Production

Put HTTPS/Nginx in front of Streamlit. Keep FastAPI internal rather than exposing port 8000 publicly. Change all default secrets before deployment.

## Notes

This version expects `.xlsx` or `.xlsm` files where Excel requests an opening password. The passwords are used only during processing and are not stored in the generated workbook.
