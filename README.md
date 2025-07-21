# Therapist Scheduling System

> **Custom web-based ABA therapy scheduling & mapping suite**

---

## Table of Contents
1. Project Overview  
2. Key Features  
3. Tech Stack  
4. Repository Layout  
5. Google Cloud / Sheets Setup  
6. Deployment  
7. Roadmap  
8. Contributing  
9. License  
10. Acknowledgements  

---

## 1. Project Overview
The **Therapist Scheduling System** is custom scheduling software purpose-built for an ABA therapy business.  
It provides two synchronized views—a spreadsheet-style **Table** and an interactive **Map**—so staff can easily add new clients, assign therapists to sessions, and adjust schedules on the fly.

Built entirely on Google’s platform (Apps Script, Sheets, and Maps), all data stays inside Google Cloud, supporting the organization’s HIPAA-compliance requirements without extra infrastructure.

This repository is a **redacted portfolio snapshot** of the original codebase.  Sensitive data and proprietary business logic have been removed, leaving just enough structure and sample code to demonstrate the architecture and coding style.

<p align="center">
  <img src="docs/demo_dashboard.png" width="700" alt="Dashboard screenshot"/>
</p>

---

## 2. Key Features
* 🔄 **Real-time Scheduling Table** – drag-and-drop grid with week, day, and time-slot granularity.
* 🗺️ **Interactive Map View** – geospatial view of therapists & clients with assignment lines, clustering, and search.
* 👥 **Multi-Therapist Assignments** – assign multiple therapists to the same client & session effortlessly.
* 🚗 **Distance & Travel Time Engine** – Google Maps API-driven calculations with offline fallbacks.
* 📋 **Google Sheets Synchronisation** – spreadsheet as single source of truth; instant bi-directional updates.
* 🔄 **One-click Schedule Copying** – current ⇄ future, therapist-level, or time-slot-level copy flows.
* 📝 **Rich Notes & Metadata** – per-assignment notes, colour-coding, indicators, and hover tooltips.
* ✅ **Conflict & Validation Rules** – double-booking prevention, availability checks, and data sanity guards.

---

## 3. Tech Stack
| Layer | Technology |
|-------|------------|
| Front-end | HTML, Vanilla JS, CSS (no framework) |
| Mapping | Google Maps JavaScript API |
| Back-end | Google Apps Script (serverless) & Velo by Wix (\*.jsw REST stubs) |
| Data | Google Sheets (multiple tabs) |
| Distance Calc | Google Maps Directions API + haversine fallback |
| DevOps | GitHub Actions (lint / deploy) – _optional sample_ |

---

## 4. Repository Layout
```
├── backend/
│   ├── distanceCalculator.jsw
│   └── geocode.jsw
├── Therapist_Scheduling_System_User_Guide.txt
└── README.md
```

---

## 5. Google Cloud / Sheets Setup
1. **Create a Google Sheet** using the structure outlined in the [User Guide](Therapist_Scheduling_System_User_Guide.txt#L110-L150).  
   * Tabs: `Therapists`, `Clients`, `Assignments`, `Sessions`, `Notes`, `ClientDistances`.
2. **Enable Google Maps APIs** in Google Cloud Console and obtain an API key.  
   * _Directions API_, _Maps JavaScript API_, _Geocoding API_.
3. **Deploy Google Apps Script** (included in `backend/`) and connect it to the Sheet.  
   * Set script properties for **SHEET_ID** and **MAPS_API_KEY**.
4. **Update `config.js`** (redacted sample) with your keys & Sheet ID.

---

## 6. Deployment
### Option A – Google Apps Script Web App
```text
1. Open the script in Apps Script editor.
2. Deploy → New deployment → Web app.
3. Execute as: Me.  Access: Anyone with link.
4. Copy the web-app URL → share with users.
```

### Option B – Wix Velo Site
1. Create a new Wix site and enable _Dev Mode_.  
2. Paste contents of `backend/*.jsw` into the Velo backend.  
3. Add the provided HTML, CSS, and JS into a page; wire up API endpoints.  
4. Publish.

---

## 7. Roadmap
- [ ] Full test suite with Jest & Playwright  
- [ ] Role-based auth & granular permissions  
- [ ] Offline-first Progressive Web App (PWA)  
- [ ] Automated route optimisation (client-to-client distance matrix)  
- [ ] Dark-mode & WCAG AA colour palette

---

## 8. Contributing
Pull requests are welcome!  For major changes, please open an issue first to discuss what you would like to change.

To contribute:
1. Fork ➡️ create feature branch ➡️ commit ➡️ open PR.  
2. Follow the existing code style (Prettier & ESLint configs).  
3. Write or update tests where relevant.

---

## 9. License
Distributed under the **MIT License**. See `LICENSE` for more information.

---

## 10. Acknowledgements
* PTLT the ABA business
* ABA therapists & BCBAs who provided domain expertise.  
* Google Maps & Google Apps Script teams.  

> _“The secret of getting ahead is getting started.”_ – Mark Twain 
