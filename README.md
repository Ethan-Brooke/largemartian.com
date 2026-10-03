# largemartian.com

Static site for Large Martian, hosted on GitHub Pages.

## Updating the public site

Edit on github.com (pencil icon), commit, and the site updates in about a minute.

- `shows.json` — upcoming and past shows. `date` as `YYYY-MM-DD`, `venue`, `city`; optional `time`, `tickets`, `note`, `poster` (image in `img/`), `recording` (link to the full-show recording, shown on past shows), `map`. Past shows move to "Past shows" automatically.
- `listen.json` — public recordings and social links.
- `img/` — artwork. `hero.jpg` is the homepage art.

## Band room (`/band/`)

Password-protected. The page's data (calendar, session archive, tracks, set lists) lives in `band/band.enc.json`, encrypted with the band password (AES-256-GCM, key from PBKDF2) and decrypted in the browser. Without the password the file is unreadable. It is not linked from the public site and is marked no-index.

**To change anything in the band room** (add a session, set the calendar, change the password): open `largemartian.com/band/lock.html`, unlock with the current password, edit the JSON, lock, download `band.enc.json`, and upload it over the old one in `band/` here on GitHub.

The plain `band/band.json` is deliberately not in this repo.

## Audio

MP3s live in a separate repo, `largemartian-audio`, served at `https://ethan-brooke.github.io/largemartian-audio/`. Full session recordings stay in the band's shared Google Drive.
