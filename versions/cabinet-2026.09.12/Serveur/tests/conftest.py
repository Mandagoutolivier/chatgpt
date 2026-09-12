"""L'adaptateur PGlite ne participe jamais aux tests PostgreSQL natifs."""
import os
import time
import psycopg

if os.getenv('CABINET_TEST_PGLITE')=='1':
    _connect=psycopg.connect
    def _connect_pglite(*args,**kwargs):
        for attempt in range(4):
            try:return _connect(*args,**kwargs)
            except psycopg.OperationalError:
                if attempt==3:raise
                time.sleep(0.025)
    psycopg.connect=_connect_pglite
