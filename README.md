# Job Finder Scanner

This project scans company career pages from `IT_Companies_with_Career_Pages_updated.xlsx` and creates daily Excel reports for Java/backend roles matching your profile.

## First Time Setup

Run:

```bat
install_dependencies.bat
```

## Daily Run

Run:

```bat
run_daily_scan.bat
```

Outputs are saved in:

```text
outputs/job_matches_YYYY-MM-DD.xlsx
outputs/job_history.xlsx
outputs/scan_log_YYYY-MM-DD.json
```

## GitHub Automatic Run

The project includes a GitHub Actions workflow:

```text
.github/workflows/daily_job_scan.yml
```

It runs every day at `09:00 India time` and `18:00 India time`, and can also be started manually from GitHub.

To use it:

1. Create a private GitHub repository.
2. Upload or push this project folder to that repository.
3. Open the repository on GitHub.
4. Go to `Actions`.
5. Select `Daily Job Scan`.
6. Click `Run workflow` for a manual test.

After the workflow finishes, download the report from the workflow run's `Artifacts` section. The artifact contains:

```text
outputs/job_matches_YYYY-MM-DD.xlsx
outputs/job_history.xlsx
outputs/scan_log_YYYY-MM-DD.json
```

Keep `IT_Companies_with_Career_Pages_updated.xlsx`, `scan_jobs.py`, `job_search_config.json`, and `requirements.txt` committed in GitHub. Daily reports and logs are uploaded as artifacts. `outputs/job_history.xlsx` is committed back to the repo by the workflow so the scanner can mark jobs as `New`, `Still Open`, or `Not Seen Today` across days.

## Telegram Setup For GitHub

The workflow sends the report to Telegram after every scheduled or manual run. Add these repository secrets in GitHub:

```text
TELEGRAM_BOT_TOKEN
TELEGRAM_CHAT_ID
```

To create the bot token:

1. Open Telegram.
2. Search for `@BotFather`.
3. Send `/newbot`.
4. Follow the prompts and copy the bot token.
5. Start a chat with your new bot and send any message, such as `hi`.

To find your chat ID:

1. Open this URL in your browser, replacing the token:

```text
https://api.telegram.org/botYOUR_BOT_TOKEN/getUpdates
```

2. Look for `"chat":{"id":...}`.
3. Copy that number as `TELEGRAM_CHAT_ID`.

To add secrets:

1. Open your GitHub repository.
2. Go to `Settings`.
3. Go to `Secrets and variables`.
4. Open `Actions`.
5. Click `New repository secret`.
6. Add `TELEGRAM_BOT_TOKEN`.
7. Add `TELEGRAM_CHAT_ID`.

## How Matching Works

The scanner prioritizes Java-related roles for 1-2 years or less than 3 years experience.

It boosts jobs mentioning:

- Java
- Spring Boot
- Spring MVC
- REST APIs
- Microservices
- SQL
- MySQL

It skips jobs that clearly look too senior, such as:

- Senior
- Lead
- Principal
- Architect
- Manager
- 3+ years or higher

## Change Your Search

Edit `job_search_config.json` to change:

- job titles
- skills
- locations
- experience filters
- scan speed

Useful scan speed settings:

```json
"worker_count": 8,
"batch_size": 40,
"delay_between_companies_seconds": 0.2
```

Use fewer workers, such as `4`, if many sites fail or your internet feels slow. Use more workers, such as `12`, if you want faster scans and the sites are responding well.

## Notes

Some career pages load jobs using JavaScript or block automated requests. Those companies will appear in the scan log as errors or with zero jobs found. Later, we can add Playwright browser automation and special handlers for Workday, Greenhouse, Lever, Ashby, and SmartRecruiters to improve coverage.
