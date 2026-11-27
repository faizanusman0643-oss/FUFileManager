# FU File Manager Pro (Android)

Gallery (photos/videos/audio/docs), file browser, batch move/copy/delete/share,
duplicate finder, sort by format/date/month, consolidate, clean junk,
hidden vault with PIN, built-in video/audio player (Media3, all formats).

## APK banane ka tareeqa (PC pe kuch install nahi)
1. github.com pe naya repo banao (e.g. FUFileManager).
2. Is folder ki sab files (hidden `.github` folder samet) repo mai upload karo.
3. Repo > Actions tab > "Build APK" > Run workflow (push pe khud bhi chalta hai).
4. 5-8 minute baad run khol ke neeche Artifacts se `FUFileManager-APK` download karo.
5. Zip kholo, `app-debug.apk` phone mai bhejo, install karo (Unknown sources allow).
6. App pehli baar kholo > "Allow access" > All files access ON.

## Android Studio se
File > Open > is folder ko kholo > Build > Build APK(s).

## WhatsApp Recovery (Tools tab)
- "Auto backup" ON karo: WhatsApp ki har nayi photo/video/voice/document/status ki copy
  `.FURecovery` folder mai foran ban jati hai (background service + notification).
- WhatsApp pe koi file delete ho ("Delete for everyone" ya khud) to wo "Deleted" tab mai milegi.
  Select > Recover = `Download/WhatsApp Recovered` mai save.
- Dekhe hue status 24 ghante baad bhi "Status" tab mai rehte hain.
- Backup ON karne se PEHLE delete hui files root ke baghair wapas nahi aa saktin
  (sirf Android trash `.trashed-*` wali milti hain).
- Backup storage leta hai: ⋮ > "Purana backup saaf karo" se 30 din+ purani copies hatao.
- View-once media support nahi hai (WhatsApp use encrypted rakhta hai, aur bhejne wale ki privacy).

## Notes
- Vault sirf chhupata hai (`.FUVault` folder + PIN), encrypt nahi karta.
- Delete permanent hai (recycle bin nahi) — delete se pehle confirm aata hai.
- PIN bhool gaye to files phone ke `.FUVault` folder mai hain.
