# 🏍️ Kesri Nandan Spares — Live Website Setup Guide
## अपनी दुकान की Website Live करें — Step by Step

---

## ✅ What You Will Get
- A live website accessible from **any phone, tablet, or PC**
- Your data synced **automatically** across all devices
- **Free** hosting and free database
- URL like: `https://kesri-nandan-spares.vercel.app`

---

## 📋 What You Need
- Gmail account (Google account)
- Computer with internet (for one-time setup)
- 20-30 minutes

---

# STEP 1: Create Firebase Database (Free)
### Firebase = Google ka free online database

1. Go to: **https://firebase.google.com**
2. Click **"Get started"** → Sign in with Gmail
3. Click **"Create a project"**
4. Project name: `kesri-nandan-spares` → Click Continue
5. **Disable Google Analytics** → Click Continue → Create Project
6. Wait 30 seconds → Click **Continue**

### Enable Realtime Database:
7. In left menu → Click **"Build"** → Click **"Realtime Database"**
8. Click **"Create Database"**
9. Select location: **Asia-southeast1 (Singapore)** → Click Next
10. Select **"Start in test mode"** → Click **Enable**
11. ✅ Your database is created!

### Get Your Config:
12. Click ⚙️ gear icon (top left) → **"Project settings"**
13. Scroll down → Click **"Add app"** → Click **Web icon** `</>`
14. App nickname: `KNS Web` → Click **Register app**
15. You will see code like this:
```javascript
const firebaseConfig = {
  apiKey: "AIzaSy...",
  authDomain: "kesri-nandan...",
  databaseURL: "https://kesri-nandan...firebaseio.com",
  projectId: "kesri-nandan...",
  ...
};
```
16. **COPY this entire config object** (you will need it in Step 3)

---

# STEP 2: Install Node.js (One-time only)

1. Go to: **https://nodejs.org**
2. Download **"LTS"** version (green button)
3. Install it (keep clicking Next/Install)
4. ✅ Done

---

# STEP 3: Setup the App Code

1. Download this project folder (kesri-nandan-live)
2. Open file: **`src/firebase.js`**
3. Find this section:
```javascript
const firebaseConfig = {
  apiKey:            "PASTE_YOUR_API_KEY_HERE",
  ...
```
4. **Replace ALL the placeholder values** with your Firebase config from Step 1
5. Save the file

---

# STEP 4: Deploy to Vercel (Free Hosting)

### Option A: Easiest — Vercel CLI

Open **Command Prompt** (Windows) or **Terminal** (Mac):

```bash
# Go to project folder
cd kesri-nandan-live

# Install dependencies
npm install

# Install Vercel tool
npm install -g vercel

# Build and deploy
npm run build
vercel --prod
```

Follow the prompts:
- Create Vercel account (free) with Gmail
- Press Enter for all questions
- ✅ You will get a URL like: `https://kesri-nandan-spares.vercel.app`

---

### Option B: Netlify Drag & Drop (Even Easier!)

1. Run in Command Prompt:
```bash
cd kesri-nandan-live
npm install
npm run build
```

2. Go to: **https://app.netlify.com/drop**
3. **Drag and drop** the `dist` folder onto that page
4. ✅ Your site is live instantly!
5. You can customize the URL in Netlify settings

---

# STEP 5: Access from Any Device

After deployment, you get a URL like:
`https://your-app-name.vercel.app`

- **Mobile**: Open URL in browser → Tap "Add to Home Screen" → Works like an app!
- **PC**: Open URL in any browser
- **Multiple devices**: All data syncs automatically via Firebase

---

## 🔒 Security (Important!)

After testing, secure your Firebase database:

1. Go to Firebase Console → Realtime Database → **Rules**
2. Change rules to:
```json
{
  "rules": {
    ".read": true,
    ".write": true
  }
}
```
This is fine for a private business app. For extra security, add Firebase Authentication later.

---

## 📱 Add to Phone Home Screen

**Android:**
1. Open your website URL in Chrome
2. Tap ⋮ menu → "Add to Home screen"
3. Tap "Add"
4. App icon appears on your home screen!

**iPhone:**
1. Open URL in Safari
2. Tap Share button (square with arrow)
3. Tap "Add to Home Screen"
4. Tap "Add"

---

## 🆘 Troubleshooting

**"npm is not recognized"** → Install Node.js from nodejs.org first

**"Firebase permission denied"** → Make sure you selected "Test mode" in Step 1

**Data not syncing** → Check your firebaseConfig values are correct in firebase.js

**Build fails** → Run `npm install` again, then `npm run build`

---

## 📞 Quick Reference

| Service | Link | Cost |
|---------|------|------|
| Firebase (Database) | firebase.google.com | FREE up to 1GB |
| Vercel (Hosting) | vercel.com | FREE forever |
| Netlify (Hosting) | netlify.com | FREE forever |
| Node.js | nodejs.org | FREE |

---

## 🎉 Congratulations!

Aapki dukaan ki website live ho gayi! Ab aap kisi bhi phone ya computer se apna business manage kar sakte hain.

For any help: Show this guide to a local computer technician — they can set this up in 30 minutes.

---
*Kesri Nandan Spares AI Business Manager v1.0*
