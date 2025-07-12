# 📅 C# Excel DateTime Importer

> **A tiny demo to prove that handling date values from Excel in C# is not rocket science.**

---

## 📖 Overview

This repository was created solely to demonstrate that **handling various date formats from Excel files in C#** is entirely feasible — despite what certain backend developers might have you believe.  
After a long and passionate debate about how "impossible", "tricky", and "unsafe" this simple task might be, we decided to settle it the proper way: with code.

---

## 🛠️ Tech Stack

- **Backend:** C# (.NET 6.0) + Entity Framework Core  
- **Frontend:** React + Ant Design + TypeScript + Vite + TailwindCSS  
- **Bonus:** Monaco Editor for displaying the C# implementation right in the browser (because why not?).

---

## 🎯 Features

- ✅ Import Excel files containing:
  - Standard date formats
  - Serial numbers (e.g., `45123` = `2023-07-15`)
  - Plain text dates (like `15/07/2023`)
  
- ✅ Convert and handle those values cleanly in C#

- ✅ Frontend allows:
  - Uploading an Excel file
  - Viewing the C# code responsible for this magical feat right in a Monaco Editor window

---

## 💬 Why This Exists

Because someone on the backend team kept listing an endless array of reasons why this couldn't (or shouldn't) be done.  
Rather than argue, we chose violence… in the form of a proof-of-concept repo.

---

## 📦 How to Run

### 🐳 Docker Compose (Recommended - Because We're Professionals)

**Option 1: Production deployment with pre-built images**
```bash
# Copy and configure environment
cp .env.example .env
# Edit .env with your Docker Hub username

# Deploy like a boss
docker-compose -f docker-compose.prod.yml up -d
```

**Option 2: Local development build**
```bash
# Build and run everything locally
docker-compose up -d

# Watch the magic happen
docker-compose logs -f
```

**Option 3: Development with hot reload (for the impatient)**
```bash
# Start with hot reload for rapid development
docker-compose -f docker-compose.dev.yml up -d
```

### 🔧 Manual Setup (For the Brave)

**Backend (.NET Core 6.0)**
```bash
cd ExcelDateImporter.Api
dotnet restore
dotnet run
```

**Frontend (React + TypeScript)**
```bash
cd ExcelDateImporter.Api/excel-date-frontend
npm install
npm run dev
```

### 🌐 Access URLs

- **Frontend:** http://localhost:3000 (Upload your Excel files here)
- **Backend API:** http://localhost:5000 (The magic happens here)
- **Backend Health:** http://localhost:5000/health (Proof it's alive)
- **Backend Swagger:** http://localhost:5000/swagger (Development only)

### 📋 Supported Excel Date Formats

Because apparently this was "impossible":

**Standard Formats**
- `dd/MM/yyyy` (25/12/2023)
- `MM/dd/yyyy` (12/25/2023) 
- `yyyy-MM-dd` (2023-12-25)
- `dd-MM-yyyy` (25-12-2023)

**Excel Serial Dates (The "Scary" Ones)**
- Excel DateTime objects
- Excel serial numbers (e.g., 44561 = 2022-01-01)
- `dd.MM.yyyy` (25.12.2023)
- `MM.dd.yyyy` (12.25.2023)

**Text Formats (Yes, Even These)**
- `dd MMM yyyy` (25 Dec 2023)
- `MMM dd, yyyy` (Dec 25, 2023)
- `d/M/yyyy` (5/3/2023)
- `M/d/yyyy` (3/5/2023)

### 🛠️ Docker Management Commands

```bash
# Start the show
docker-compose up -d

# End the show
docker-compose down

# Watch the action
docker-compose logs -f [service-name]

# Rebuild (when you break things)
docker-compose build

# Get latest images
docker-compose -f docker-compose.prod.yml pull
```

### 🏗️ Architecture

```
┌─────────────────┐    HTTP     ┌─────────────────┐
│   React SPA     │   Requests   │   .NET Core     │
│   (Port 3000)   │ ──────────► │   Web API       │
│   Ant Design +  │             │   + EF Core     │
│   Monaco Editor │             │   (Port 5000)   │
└─────────────────┘             └─────────────────┘
        │                              │
        │                              │
        ▼                              ▼
┌─────────────────┐             ┌─────────────────┐
│   Nginx         │             │   In-Memory     │
│   (Static Files)│             │   Database      │
└─────────────────┘             └─────────────────┘
```

### 🚀 CI/CD Features

- Multi-architecture Docker builds (AMD64/ARM64)
- GitHub Actions workflows
- Security scanning with Trivy
- Automated deployment artifacts

### 🐛 Troubleshooting

**"It doesn't work!" - Debug Edition**

```bash
# Check what's actually running
docker-compose ps

# See what went wrong
docker-compose logs backend
docker-compose logs frontend

# Try turning it off and on again
docker-compose restart backend
```

**Common Issues & Solutions:**
1. **CORS Errors**: The backend is probably not running or CORS isn't configured
2. **File Upload Issues**: Check file size limits and ensure you're using .xlsx/.xls files  
3. **Connection Refused**: Services might not be started or ports are blocked

---

## ✌️ License

MIT — because even petty projects deserve freedom.
