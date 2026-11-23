# Daily Attendance Email Automation

Python example for sending a dated attendance report email through Microsoft Outlook on Windows.

## What it does

Calculates the previous reporting date, prepares an HTML email, adds a local attachment and sends it through Outlook. On Monday, the date helper selects the previous Friday. Other days select the preceding calendar day.

## Files and setup

- `Daily_Email ATTD.ipynb`: notebook example.
- `.env.example`: local configuration names with non-working example values.
- `.gitignore`: excludes local configuration, report data and outputs.

Use Windows, a configured Outlook desktop installation, Python, Jupyter and pywin32. Set `ATTENDANCE_RECIPIENT`, `ATTENDANCE_DASHBOARD_URL` and `ATTENDANCE_ATTACHMENT_PATH` in the local process environment. The example file is not loaded automatically.

## Safety

Running the notebook sends email, not a draft. Review the recipient, dashboard link, attachment and reporting date before execution. Use synthetic data and a test mailbox first. Do not run it against a real recipient merely to preview the example.

## Limitations

The date rule does not account for public holidays or all workweek patterns. No scheduler configuration is included. The notebook uses a local file and Outlook COM, so it is not cross-platform. No measured time-saving claim is made. This update has been statically checked, not executed against Outlook.
