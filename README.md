# Resume Roast

A no-login Streamlit demo for resume feedback, powered by Ollama Cloud Beam when an API key is configured.

## Run on Windows

Open PowerShell in this project folder and run:

```powershell
.\run_app.ps1
```

The script creates the local virtual environment if needed, installs packages from `requirements.txt`, and starts Streamlit. It uses the local URL printed in the terminal.

## Enable Ollama Cloud

The app loads the local `.env` file when it starts and calls Ollama's hosted chat endpoint directly; a local Ollama installation is not required. Create an Ollama API key and add it to `.env`:

```text
OLLAMA_API_KEY=your_key_here
OLLAMA_MODEL=beam:cloud
OLLAMA_API_URL=https://ollama.com/api/chat
OLLAMA_TIMEOUT_SECONDS=180
WORKSHOP_URL=https://nxtwave.tech/
```

Keep the key private and out of source control. Without `OLLAMA_API_KEY`, the app starts in a clearly labeled rules-based preview mode. The feedback styles are Gentle nudge, Talk directly, and Full roast (Gen Z). The app requests concise JSON in the prompt and retries once if the model returns incomplete or invalid JSON.

## Privacy

The app has no login, database, or file-writing code for resumes. Uploaded files and extracted text are handled in temporary app memory. After explicit consent, extracted resume text is sent to Ollama Cloud for feedback. Review Ollama and hosting data terms before using real resumes; this app cannot guarantee how an external provider or hosting platform retains request data.

The tool gives coaching feedback, not an ATS score or a hiring recommendation.
