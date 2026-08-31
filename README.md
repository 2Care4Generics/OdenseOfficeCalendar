# Room Status Dashboard (GitHub Pages version)

Shows live free/busy status for beC, beJava, and beSwift, hosted for free on
GitHub Pages. A scheduled GitHub Action re-fetches the calendars and
regenerates the page automatically.

## How it works

- `scripts/generate.js` fetches the 3 room ICS feeds, works out each room's
  current status, and writes `docs/index.html`.
- `.github/workflows/update-dashboard.yml` runs that script every 10 minutes
  (and whenever you push to `main`) and commits the regenerated page.
- GitHub Pages serves `docs/index.html` as your live site.

## Setup

1. Create a new **public** GitHub repo (private repos work too, but public
   repos get unlimited free Actions minutes; private repos get ~2,000
   minutes/month free, which is still plenty for a 10-minute schedule).
2. Upload this whole folder (`scripts/`, `.github/`, `docs/`, `README.md`)
   into the repo, keeping the folder structure exactly as-is.
3. Go to **Settings → Pages** in your repo.
   - Source: **Deploy from a branch**
   - Branch: **main**, folder: **/docs**
   - Save.
4. Go to the **Actions** tab, find "Update Room Dashboard", and click
   **Run workflow** once to generate the first version of the page.
5. After a minute or two, your site will be live at:
   `https://<your-username>.github.io/<repo-name>/`

From then on, the Action keeps refreshing the page automatically every
10 minutes. Share that URL with your team.

## Adjusting things

- **Refresh frequency**: change the `cron` line in
  `.github/workflows/update-dashboard.yml`. GitHub's practical minimum is
  about every 5 minutes; scheduled runs can also be delayed a few minutes
  under GitHub's own load, so treat this as "near real-time," not
  second-by-second.
- **Timezone**: change `TIMEZONE` at the top of `scripts/generate.js` if your
  office isn't in `Europe/Copenhagen`.
- **Room colors**: change the `color` values in the `ROOMS` array in
  `scripts/generate.js`.
