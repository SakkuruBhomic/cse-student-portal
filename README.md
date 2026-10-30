# CSE (Sections 6, 7, 8) Student Academic Details Portal

A mobile-first, high-performance Google Apps Script Web App designed for collecting and managing student academic details (SSC Marks, Intermediate Marks, EAMCET Rank, Rural/Urban status) with real-time two-way synchronization to Google Sheets.

**Created by Bhomic (CSE-7)**

---

## 🚀 Live Web App

* **Live Deployment URL:** [CSE 6,7,8 Student Portal](https://script.google.com/macros/s/AKfycbwNUOvVIn_WXNDmy2M7nc9827YwA3RvJsvGDSGX1cQE9IH17X4jIwQMJtY-xiJMmDRjew/exec)
* **Direct Account Access (Multi-account bypass):** [Portal via authuser](https://script.google.com/macros/s/AKfycbwNUOvVIn_WXNDmy2M7nc9827YwA3RvJsvGDSGX1cQE9IH17X4jIwQMJtY-xiJMmDRjew/exec?authuser=sbhomic11208@gmail.com)
* **Connected Google Sheet:** `CSE_Master` (Tab: `CSE_6,7,8`)

---

## ✨ Features

- **Mobile-First Student Form**: Fast, responsive layout with intuitive auto-completion by student name or admission number.
- **Direct Google Sheet Sync**: Updates are written directly to columns E (SSC), F (Inter), G (EAMCET), and H (Rural/Urban) of worksheet `CSE_6,7,8`.
- **Concurrency Locking**: Built-in Google Apps Script `LockService` prevents race conditions when multiple students submit simultaneously.
- **Admin Dashboard**:
  - Live progress monitoring (Total students, Completed, Pending, and Completion Rate %).
  - Real-time search and filtering (All, Completed, Pending).
  - Offline Excel export (`.xlsx`) generated client-side via SheetJS.
- **Pre-configured Manifest**: Automated OAuth scopes for Google Drive and Google Sheets with anonymous public access.

---

## 📁 Repository Structure

```
cse-student-portal/
├── Code.gs             # Apps Script backend (Direct sheet sync, LockService, API handlers)
├── Index.html          # Frontend (Mobile-first form, autocomplete, Admin dashboard)
├── appsscript.json     # Apps Script project manifest (timezone, scopes, webapp config)
├── .clasp.json         # Google Clasp project configuration
└── .gitignore          # Git ignore rules
```

---

## 🛠️ Deployment & Maintenance via Clasp

This project is managed locally using [`@google/clasp`](https://github.com/google/clasp).

### 1. Authenticate Clasp
```bash
npx @google/clasp login
```

### 2. Push Code Updates
```bash
npx @google/clasp push -f
```

### 3. Deploy a New Version
```bash
npx @google/clasp deploy --description "Production Release"
```

---

## 📄 License

MIT License. Designed and developed by **Bhomic (CSE-7)**.
