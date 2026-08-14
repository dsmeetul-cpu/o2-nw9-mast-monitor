# O2 NW9 0RY Mast Monitor

Automated monitoring and evidence collection for the O2 network status affecting:

**NW9 0RY**

## O2 Complaint

Complaint reference:

C-1308267357

## Purpose

This project automatically checks the O2 network status for NW9 0RY every hour.

Each check records:

- Date
- Time
- Postcode
- O2 status
- O2 displayed message
- Expected resolution
- Check result
- Status changes
- Screenshot evidence
- Complaint reference

## Evidence

Screenshots are stored in:

evidence/

Structured data is stored in:

data/o2-nw9-0ry-monitor.csv

Excel evidence is stored in:

data/O2-NW9-0RY-Evidence.xlsx

## Source

https://status.o2.co.uk/

## Monitoring

The GitHub Actions workflow runs approximately once per hour.

The workflow can also be manually triggered from:

GitHub → Actions → O2 NW9 0RY Network Monitor → Run workflow

## Important

This is an evidence collection tool.

It does not independently establish that O2's network was unavailable. It records what the O2 status website displayed at each check together with a timestamped screenshot.

Screenshots should be retained as original evidence.
