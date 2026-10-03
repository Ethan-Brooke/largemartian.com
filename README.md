# largemartian.com

Static site for Large Martian, hosted on GitHub Pages.

## Updating

Everything editable lives in three JSON files. Edit them on github.com (pencil icon), commit, and the site updates in about a minute.

- `shows.json` — upcoming and past shows (homepage). Add a show with `date` as `YYYY-MM-DD`, `venue`, `city`, optional `time`, `tickets` link, `note`. Past shows move to the "Past shows" list automatically.
- `listen.json` — public recordings and social links (homepage).
- `band/band.json` — the band room: Google Calendar embed URL, practice recordings by session, set lists.

## Band room

`largemartian.com/band/` is not linked from the public site and is marked no-index, but it is not password-protected — GitHub Pages can't do that. Treat the link as unlisted.

## Audio

MP3s live in a separate repo, `largemartian-audio`, served at `https://ethan-brooke.github.io/largemartian-audio/`. Folders match `folder` in `band/band.json`; the public set is under `public/`.
