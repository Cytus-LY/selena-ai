![ezgif-7b176ba4f61dde85](https://github.com/user-attachments/assets/c0de0ff0-cebe-4deb-8fc7-a3bb04f9d099)
# Selena AI

An AI-powered all-in-one assistant focused on intelligent PPT generation and real-world usability.

---

## ✨ Features

* 🧠 AI-powered PPT generation (with automatic image integration)
* 💬 Multi-mode interaction: Chat / Code / Image / PPT
* 📂 File upload support (PDF / DOCX / TXT analysis)
* ⚡ Streaming response with stop control
* 🎨 Clean and modern UI design (dark mode supported)
* 🛠 Fully local deployment with Flask backend

---

## 🚀 Demo

> Below are real screenshots from the system:

* AI-generated PPT preview
* Interactive UI with multi-mode switching
* File upload and content parsing

![ezgif-7b176ba4f61dde85](https://github.com/user-attachments/assets/5babf380-b7a8-4188-85b0-157b1ad0870d)![Uploading ezgif-7b176ba4f61dde85.gif…]()
<img width="1260" height="857" alt="屏幕截图 2026-04-11 015231" src="https://github.com/user-attachments/assets/1b46f613-b305-41f6-a66a-56db12eccae8" />
<img width="1267" height="863" alt="屏幕截图 2026-04-11 015222" src="https://github.com/user-attachments/assets/607aa7a8-ef30-413d-b51d-9411882cc153" />
<img width="1153" height="653" alt="屏幕截图 2026-04-10 235808" src="https://github.com/user-attachments/assets/f6c15c73-1f35-4a3e-83c8-a747534600f2" />



---

## 🛠 Tech Stack

* **Backend:** Flask (Python)
* **Frontend:** HTML / CSS / JavaScript
* **AI API:** OpenRouter (GPT-based models)
* **Other:** Markdown rendering, streaming response, file parsing

---
## ⚠️ Configuration

This project requires an API key from [OpenRouter](https://openrouter.ai/).

1. **The Easy Way (Recommended):**
   - Copy `.env.example` to a new file named `.env`.
   - Open `.env` and paste your key: `OPENROUTER_API_KEY=your_api_key_here`

2. **The Manual Way (Environment Variable):**
   - **Windows (CMD):** `set OPENROUTER_API_KEY=your_api_key_here`
   - **PowerShell:** `$env:OPENROUTER_API_KEY="your_api_key_here"`
   - **Mac / Linux:** `export OPENROUTER_API_KEY=your_api_key_here`

## ⚡ Run Locally

```bash
pip install -r requirements.txt
python app.py
```

Then open your browser:

```
http://127.0.0.1:5000
```

---

## 🎯 Project Motivation

This project was built to explore practical applications of AI in productivity tools, especially in automated presentation generation.

Instead of being a simple demo, the system focuses on:

* Real usability
* Full-stack integration
* Clean UI/UX experience
* Practical AI workflow design

---

## 📌 Future Improvements

* Better layout engine for PPT generation
* Smarter image selection and positioning
* Online deployment (cloud-based access)
* More advanced AI memory system

---

## 👨‍💻 Author

* GitHub: https://github.com/Cytus-LY

---

## ⭐ Notes

This project is designed as a practical AI application prototype and can be extended into a production-ready system.

