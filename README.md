# FU File Manager (Android)

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

## Notes
- Vault sirf chhupata hai (`.FUVault` folder + PIN), encrypt nahi karta.
- Delete permanent hai (recycle bin nahi) — delete se pehle confirm aata hai.
- PIN bhool gaye to files phone ke `.FUVault` folder mai hain.
