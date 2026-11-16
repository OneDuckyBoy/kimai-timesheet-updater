# A tiny image that turns a Kimai CSV + a styled Excel template into the next
# month's timesheet. It is meant to be *called* when needed (a batch job, the
# ai-automation worker) rather than run as a long-lived service:
#
#   docker run --rm -v "$PWD/input:/in" -v "$PWD/output:/out" \
#     kimai-timesheet-updater \
#     /in/template.xlsx /in/export.csv -o /out/next-month.xlsx
#
# Files travel over mounted volumes: the script takes <template> <csv> [-o out].
FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY update_timesheet.py ./

ENTRYPOINT ["python", "/app/update_timesheet.py"]