from __future__ import annotations

from pathlib import Path
from datetime import datetime
import json
import subprocess
import sys
import traceback

PROJECT = Path(__file__).resolve().parent
LOG_DIR = PROJECT / 'official_calendar_cache'
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / 'daily_refresh.log'


def log(msg: str):
    line = f'[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}'
    print(line)
    with LOG_FILE.open('a', encoding='utf-8') as f:
        f.write(line + '\n')


def run(script: str):
    log(f'RUN {script}')
    proc = subprocess.run([sys.executable, str(PROJECT / script)], cwd=str(PROJECT), capture_output=True, text=True, encoding='utf-8', errors='replace')
    if proc.stdout.strip():
        log(proc.stdout.strip())
    if proc.stderr.strip():
        log(proc.stderr.strip())
    if proc.returncode != 0:
        raise RuntimeError(f'{script} failed with code {proc.returncode}')


def main():
    try:
        log('START daily official calendar refresh')
        run('fetch_official_calendar_sources.py')
        run('update_calendar_workbook.py')
        log('DONE daily official calendar refresh')
    except Exception as e:
        log('FAILED ' + str(e))
        log(traceback.format_exc())
        raise


if __name__ == '__main__':
    main()
