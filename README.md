# Dashboard Slip Gaji - PT Aksara Perdana Multimedia

Tampilan (GitHub Pages) + data (Google Apps Script web app) yang membaca Google Spreadsheet.

```
src/Code.gs          backend Apps Script (login, slip, ganti password, API JSON)
src/index.html       tampilan dashboard (terang, responsif, PDF + watermark)
src/appsscript.json  izin & pengaturan web app
.clasp.json          Script ID Apps Script (isi sendiri)
.github/workflows/
  deploy.yml         push src/ ke Apps Script (clasp) setiap commit ke main
  pages.yml          publikasi src/index.html ke GitHub Pages
```

## Secrets (Settings > Secrets and variables > Actions)
| Nama | Isi |
|---|---|
| `CLASPRC_JSON` | isi file `~/.clasprc.json` hasil `clasp login` |
| `DEPLOYMENT_ID` | ID deployment web app (`AKfycb...`) |
| `WEBAPP_URL` | URL web app berakhiran `/exec` |

## Langkah singkat
1. Buat project Apps Script, isi `scriptId` di `.clasp.json`, deploy web app (Execute as: Me, Access: Anyone), jalankan fungsi `pub` sekali untuk otorisasi.
2. Aktifkan Google Apps Script API di https://script.google.com/home/usersettings
3. Isi 3 secret di atas.
4. Settings > Pages > Source: **GitHub Actions** (repo harus Public untuk akun gratis).
5. Push ke `main`. Alamat situs: `https://NAMA-GITHUB.github.io/NAMA-REPO/`

Jangan pernah meng-commit `.clasprc.json`.
