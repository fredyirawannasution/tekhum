# tekhum.pages.dev - Cloudflare Pages Package

Paket ini untuk menampilkan aplikasi Google Apps Script di domain Cloudflare Pages: `https://tekhum.pages.dev`.

## Isi folder

- `index.html` - halaman Cloudflare Pages yang menampilkan Web App Apps Script di dalam iframe.
- `config.js` - tempat mengisi URL Web App Apps Script.
- `_headers` - header dasar Cloudflare Pages.
- `_redirects` - fallback routing ke `index.html`.
- `apps-script/Code.gs` - backend Apps Script dari kode aplikasi terbaru.
- `apps-script/index.html` - HTML Apps Script dari kode aplikasi terbaru.

## Deploy via GitHub ke Cloudflare Pages

1. Buat repository GitHub bernama `tekhum`.
2. Upload semua file dalam folder ini ke repository tersebut.
3. Edit `config.js`, ganti `PASTE_APPS_SCRIPT_WEB_APP_URL_HERE` dengan URL Web App Apps Script yang berakhiran `/exec`.
4. Cloudflare Dashboard > Workers & Pages > Create application > Pages > Connect to Git.
5. Pilih repo `tekhum`.
6. Project name: `tekhum`.
7. Build settings:
   - Framework preset: `None`
   - Build command: kosong
   - Build output directory: `/`
8. Deploy.
9. Domain akan menjadi `https://tekhum.pages.dev`.

## Catatan

Aplikasi utama masih memakai Google Apps Script, jadi URL Apps Script harus sudah dideploy sebagai Web App. Kode `Code.gs` sudah memakai:

```js
.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
```

Jika frame kosong, pakai URL deployment `/exec`, bukan `/dev`, lalu redeploy Apps Script.
