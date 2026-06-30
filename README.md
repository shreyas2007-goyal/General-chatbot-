# SG General Chatbot

A modern AI-powered chatbot built with **Python** and **Streamlit** that brings multiple leading AI providers into a single application. The chatbot supports multi-document conversations, vision-enabled AI models, persistent chat history, and automatic provider switching through an intuitive interface.

Developed by **Shreyas Goyal (SG)**.

---

## Overview

SG General Chatbot provides a unified interface for interacting with multiple AI models while allowing users to upload and chat with documents, images, presentations, spreadsheets, and source code.

The application is designed to be lightweight, easy to extend, and simple to deploy without requiring complex frontend frameworks.

---

## Features

### AI Providers

- OpenAI
- Google Gemini
- Anthropic Claude
- Groq
- Ollama (Local Models)

Switch between providers at any time without restarting the application.

---

### Multi-Document Chat

Upload multiple files simultaneously and ask questions across all uploaded content.

Supported formats include:

- PDF
- DOCX
- PPTX
- XLSX
- TXT
- Markdown
- Python files
- PNG
- JPG
- JPEG

Documents remain active until removed, allowing continuous conversations across multiple files.

---

### Vision Support

Compatible vision models can analyze:

- Images
- PowerPoint slides
- Excel sheets
- Charts
- Layouts
- Diagrams

When LibreOffice is installed, presentations and spreadsheets are automatically rendered into images for improved visual understanding.

---

### Persistent Chat History

- Conversations are automatically saved
- SQLite-based storage
- Automatic chat titles
- Resume previous conversations
- Delete unwanted chats
- Start new conversations anytime

---

### Modern Interface

- Custom Streamlit interface
- Provider-specific chat styling
- Timestamped messages
- Sidebar navigation
- Organized settings

---

### Automatic Web Search

The application can detect queries that require recent information and perform a live web search before generating a response.

If live search is unavailable, it automatically falls back to the model's internal knowledge.

---

## Project Structure

```
general_chatbot/
│
├── app.py
├── requirements.txt
├── README.md
└── chat_history.db
```

---

## Installation

### Clone the Repository

```bash
git clone https://github.com/yourusername/general_chatbot.git

cd general_chatbot
```

### Install Dependencies

```bash
pip install -r requirements.txt
```

---

## Optional: Install LibreOffice

LibreOffice enables rendering of PowerPoint and Excel files into images.

Download from:

https://www.libreoffice.org/download

Verify installation:

```bash
soffice --version
```

---

## Optional: Install Ollama

For local AI models:

```bash
ollama serve

ollama pull llama3.2

ollama pull llava
```

---

## API Keys

Cloud providers require their respective API keys.

| Provider | API Key |
|----------|----------|
| OpenAI | OpenAI API Key |
| Gemini | Google AI Studio API Key |
| Anthropic | Claude API Key |
| Groq | Groq API Key |

Keys are entered through the sidebar and remain only for the active session.

---

## Run the Application

```bash
streamlit run app.py
```

The application starts at:

```
http://localhost:8501
```

---

## Supported File Types

| Category | Formats |
|-----------|----------|
| Documents | PDF, DOCX, TXT, MD |
| Presentations | PPTX |
| Spreadsheets | XLSX |
| Images | PNG, JPG, JPEG |
| Source Code | PY |

---

## How It Works

1. Select an AI provider.
2. Choose a compatible model.
3. Enter the API key if required.
4. Upload one or more files.
5. Ask questions naturally.
6. Continue conversations with persistent chat history.

---

## Technologies Used

### Frontend

- Streamlit

### Backend

- Python

### AI Providers

- OpenAI
- Google Gemini
- Anthropic
- Groq
- Ollama

### Database

- SQLite

### Libraries

- python-docx
- PyPDF2
- openpyxl
- python-pptx
- Pillow
- requests

---

## Current Limitations

- Uses direct document context instead of vector databases (RAG).
- Very large documents are truncated for efficiency.
- Vision rendering for PPTX and XLSX requires LibreOffice.
- Ollama image support depends on the selected model.
- API keys are not stored permanently.

---

## Future Improvements

- Streaming responses
- RAG with embeddings
- Audio transcription
- Video understanding
- Full-text search across conversations
- User authentication
- Cloud synchronization
- Per-document selection
- Export conversations
- Dark/Light themes

---

## Contributing

Contributions are welcome.

1. Fork the repository.
2. Create a feature branch.
3. Commit your changes.
4. Submit a Pull Request.

---

## License

This project is released under the MIT License.

---

## Author

**Shreyas Goyal (SG)**

GitHub: https://github.com/shreyas2007-goyal

LinkedIn: https://linkedin.com/in/shreyas-goyalofficial

---

If you find this project useful, consider giving it a star.