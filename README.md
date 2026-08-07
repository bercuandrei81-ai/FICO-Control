# FICO Control Mobile – Starter

Acest pachet conține:

- `mobile/` – aplicația iPhone + Android în Expo / React Native
- `backend/` – versiunea FastAPI care expune și API JSON pentru aplicația mobilă

## Arhitectură

Șofer:
FICO Control App -> API online -> baza de date

Admin:
Browser -> Admin Dashboard -> aceeași bază de date

## Ce este deja pregătit

- Lista șoferilor programați din ziua curentă
- Selectarea șoferului în aplicația mobilă
- Introducerea scorului FICO
- Trimiterea scorului prin API
- Detectarea trimiterilor duplicate
- Compatibilitate cu dashboard-ul existent
- Import Excel Cortex `.xlsx`

## Ce facem în continuare

1. Publicăm backend-ul pe internet.
2. În `mobile/app/index.js` schimbăm `API_BASE` cu adresa serverului online.
3. Instalăm Node.js + Expo pe MacBook.
4. Testăm aplicația reală pe iPhone și Android.
5. Adăugăm autentificarea/codul personal pentru șofer.
6. Adăugăm upload fotografie FICO.
7. Pregătim icon, splash screen, privacy policy și build-urile pentru App Store / Google Play.

## Pornire mobilă după instalarea Node.js

În Terminal:

cd mobile
npm install
npx expo start

Apoi aplicația poate fi testată cu Expo Go înainte de publicarea în store.
