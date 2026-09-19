# bcschedule

College timetable from a Google Sheet, living in your calendar.
Live at https://schdl.eu.cc/

## Motivation

Google Docs are great for collaborating, but terrible for information that changes
rapidly: silent edits, no notifications, and good luck finding anything in a hurry
on your phone. Therefore, this thing:

- Google Sheet timetable → `.ics` in your calendar app
- Auto-updates every 6 hours via GitHub Actions
- Technically serverless (static files plus a cron job)

## Use

Pick your group on the [site](https://schdl.eu.cc/) and subscribe:
Google Calendar or any `webcal` client. Clicking a card shows a week preview.

## Run it yourself

```sh
python3 schedule_to_ics.py --group all --weeks 3   # everything
python3 schedule_to_ics.py --group 1КІ-25 --weeks 4 -o out.ics
```

Stdlib only, no dependencies. `--flip-weeks` swaps чисельник/знаменник if the
college starts the semester on the other foot.

## Layout

- `schedule_to_ics.py` — fetcher, parser, `.ics` writer, page generator (`app.js` inlined at build, `style.css` linked)
- `guide.html`, `info.html` — static docs (how-to, about/data sources)
- `index.html` — generated, do not edit by hand
- `calendars/` — generated per-group files
- `.github/workflows/schedule.yml` — the 6-hour regen
