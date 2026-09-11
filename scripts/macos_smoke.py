"""Check packaged startup using temporary data, never a real user's database."""
import os
from pathlib import Path
import sqlite3
import subprocess
import sys
import tempfile
import time


def main():
    executable = Path(sys.argv[1]).resolve()
    with tempfile.TemporaryDirectory(prefix='campus-smoke-') as directory:
        data = Path(directory) / 'data'
        env = dict(os.environ, CAMPUS_JOB_TRACKER_HOME=str(data))
        with (Path(directory) / 'startup.log').open('w+') as log:
            process = subprocess.Popen([str(executable)], env=env, stdout=log, stderr=log)
            try:
                for _ in range(60):
                    if process.poll() is not None:
                        raise RuntimeError('Packaged app exited before initialization')
                    if (data / 'jobs.db').exists() and (data / 'documents').is_dir():
                        break
                    time.sleep(1)
                else:
                    raise RuntimeError('Data directory initialization timed out')
                time.sleep(3)
                if process.poll() is not None:
                    raise RuntimeError('Packaged app exited after initialization')
                with sqlite3.connect(data / 'jobs.db') as conn:
                    tables = {row[0] for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                    assert {'jobs', 'documents'} <= tables, tables
                print('Packaged startup and database initialization passed')
            except Exception:
                log.flush()
                log.seek(0)
                print(log.read()[-8000:])
                raise
            finally:
                if process.poll() is None:
                    process.terminate()
                    try:
                        process.wait(timeout=10)
                    except subprocess.TimeoutExpired:
                        process.kill()
                        process.wait()


if __name__ == '__main__':
    main()
