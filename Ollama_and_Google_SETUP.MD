# AIPO Attendance Digitizer — Setup Guide

This project is a Flask-based automation utility designed to process scanned attendance sheets. It utilizes a local vision language model (Ollama) to extract tabular timestamp rows and automatically exports formatted records into localized Excel sheets or cloud-hosted Google Sheets instances.

---

## 1. Prerequisites & Environment Setup

I developed this with Python 3.10+. Ensure you have Python 3.10+ installed on your system, or kindly adjust the settings.

- **Step 1**: Clone the repo and navigate into the project folder
- **Step 2**: Create and activate a isolated virtual environment

**On macOS/Linux:**
```bash
python3 -m venv venv
source venv/bin/activate
```

**On Windows:**
```bash
python -m venv venv
.\venv\Scripts\activate
```

- **Step 3**: Install the required dependencies

```bash
pip install -r requirements.txt
```

Ensure your `requirements.txt` contains:
- `flask`
- `ollama`
- `openpyxl`
- `google-api-python-client`
- `google-auth-oauthlib`
- `google-auth-httplib2`

I'm not entirely sure with google-api but pls keep it muna.
---

## 2. Setting Up the Vision Engine (Ollama)

This application runs structural OCR and image inference **locally** using Ollama. It does not send image data to external cloud APIs (except compiled text data to Google Drive for sheet creation).

### Step 1: Download and Install Ollama

Go to [ollama.com](https://ollama.com) and download the client optimized for your operating system (macOS, Windows, or Linux).

### Step 2: Launch the Ollama Desktop App

Ensure the Ollama application is actively running in your system tray.

### Step 3: Pull the Required Vision Model

Open your terminal window and pull the multi-modal model utilized by this app:

```bash
ollama pull qwen2.5vl:7b
```

**Note:** This model is approximately 6 GB for me. Ensure you have a stable network connection and enough storage space.

---

## 3. Google Cloud Project & OAuth Credentials Setup

To write directly to Google Sheets, you need to authenticate your application via the Google Cloud Console.

### Part A: Create a Google Cloud Project

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Log in with your preferred Google Account
3. Click **Create Project**
4. Name your project (e.g., `AIPO-Attendance-Digitizer`)
5. Click **Create**

### Part B: Enable the Required APIs

1. In the sidebar, navigate to **APIs & Services** > **Library**
2. Search for **Google Sheets API** and click **Enable**
3. Go back to the Library
4. Search for **Google Drive API** and click **Enable**

### Part C: Configure the OAuth Consent Screen

1. Go to **APIs & Services** > **OAuth consent screen**
2. Choose **External** (unless you are using an enterprise Google Workspace account that allows *Internal*)
3. Click **Create**
4. Fill out the required basic information:
   - **App name:** `AIPO Attendance Digitizer`
   - **User support email:** Your email address
   - **Developer contact information:** Your email address
5. Click **Save and Continue** through the *Scopes* section
6. **Important Step (Test Users):** Because the app is in "Testing" mode, you must add the Google accounts of anyone testing this app
   - Under **Test Users**, click **+ Add Users**
   - Enter your groupmates' Google email addresses
   - If an email isn't on this list, Google will block them with an error screen

### Part D: Generate credentials.json

1. Navigate to **APIs & Services** > **Credentials**
2. Click **+ Create Credentials** at the top
3. Select **OAuth client ID**
4. Set the *Application type* to **Desktop app**
5. Name it (e.g., `DTR Digitizer Client`)
6. Click **Create**
7. A dialog box will appear — click **Download JSON**
8. **Rename the downloaded file exactly to `credentials.json`** and place it in the root of your project directory (alongside `app.py`)

⚠️ **SECURITY WARNING:** Do **NOT** commit `credentials.json` or local token files to GitHub. They are listed in `.gitignore` to prevent leaking private credentials.

---

## 4. Running the Application

1. Ensure Ollama is running:
   ```bash
   ollama list
   ```
   You should see `qwen2.5vl:7b` in the output.

2. Make sure your virtual environment is active:
   ```bash
   source venv/bin/activate
   ```

3. Start the Flask development server:
   ```bash
   python3 app.py
   ```

4. Open your browser and navigate to:
   ```
   http://localhost:5000
   ```

5. On the first export to Google Sheets, your browser will open a sign-in prompt:
   - Click **Advanced**
   - Click **Go to AIPO Attendance Digitizer (unsafe)**
   - This approves local script authorization for Google Sheets access

---

## 5. Project Structure

```
attendance-digitizer/
├── app.py                 # Flask backend with Ollama integration
├── requirements.txt       # Python dependencies
├── credentials.json       # Google OAuth credentials (PUT IN GITIGNORE)
├── .gitignore            # Git ignore rules
├── Ollama_and_Google_SETUP.MD             # This file
├── README.md             # Project overview
└── templates/
    └── index.html        # Web interface (minimal, black & white design)
```

---

## 6. Troubleshooting

### Ollama not connecting
- Ensure `ollama serve` is running or the Ollama app is open
- Check that `qwen2.5vl:7b` is pulled: `ollama list`

### Browser shows "Access Denied"
- Try `http://localhost:5000` instead of `http://127.0.0.1:5000`
- Hard refresh: `Cmd + Shift + R` (Mac) or `Ctrl + Shift + R` (Windows/Linux)

### Google Sheets export fails
- Verify `credentials.json` is in the project root
- Check that your Google account is added to "Test Users" in the OAuth consent screen
- Ensure Google Sheets API and Google Drive API are enabled in Cloud Console

### Model takes too long to process
- First scan loads the 4.7GB model into RAM — may take 30-60 seconds
- Subsequent scans are faster (~10-20 seconds per sheet)

---

## 7. Development Notes

- **Data Privacy:** All image processing happens locally with Ollama. No images are sent to external APIs.
- **Model:** `qwen2.5vl:7b` is a 8.3B parameter vision-language model optimized for handwriting recognition
- **Export Formats:** CSV and Excel (.xlsx) with proper DTR formatting matching AIPO's template
- **Browser Compatibility:** Works on Chrome, Safari, Firefox (recommend Chrome for best experience)

---

## 8. Support & Next Steps

For issues or questions:
1. Check the Troubleshooting section above
2. Verify all prerequisites are installed
3. Review terminal output for specific error messages
4. Consult the main README.md for project overview