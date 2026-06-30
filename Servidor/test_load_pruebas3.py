"""Prueba de carga contra solidpruebas3 (sin tocar solid)."""
import os
import sys
import time

basedir = os.path.dirname(os.path.abspath(__file__))
os.chdir(basedir)
sys.path.insert(0, basedir)
os.environ["POLARS_SKIP_CPU_CHECK"] = "1"

# Override de BD antes de importar el modulo de carga
os.environ["DB_URI"] = "postgresql://postgres:AdminSolid2025@192.168.0.21:5432/solidpruebas3"

import logging
from sqlalchemy import create_engine, text

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("TestLoadPruebas3")

from upload_buffer_polars import (
    DB_URI,
    TABLE_NAME_PG,
    TABLE_CATALOGO,
    set_logger,
    upload_buffer_with_merge,
    upload_catalogo,
)

set_logger(logger)


def verify_target_db():
    logger.info("Verificando destino: %s", DB_URI.split("@")[-1])
    engine = create_engine(DB_URI)
    with engine.connect() as conn:
        tables = conn.execute(
            text(
                """
                SELECT table_name
                FROM information_schema.tables
                WHERE table_schema = 'public'
                  AND table_name IN (:t1, :t2)
                ORDER BY table_name
                """
            ),
            {"t1": TABLE_NAME_PG, "t2": TABLE_CATALOGO},
        ).fetchall()
    found = [r[0] for r in tables]
    logger.info("Tablas encontradas: %s", found)
    missing = {TABLE_NAME_PG, TABLE_CATALOGO} - set(found)
    if missing:
        raise RuntimeError(f"Faltan tablas en solidpruebas3: {sorted(missing)}")
    return engine


def count_rows(engine, table_name: str) -> int:
    with engine.connect() as conn:
        return conn.execute(text(f"SELECT COUNT(*) FROM {table_name}")).scalar_one()


def main():
    what = (sys.argv[1] if len(sys.argv) > 1 else "all").lower()
    logger.info("=" * 70)
    logger.info("PRUEBA DE CARGA -> solidpruebas3 (modo: %s)", what)
    logger.info("=" * 70)

    engine = verify_target_db()
    before_ot = count_rows(engine, TABLE_NAME_PG)
    before_cat = count_rows(engine, TABLE_CATALOGO)
    logger.info("Antes - %s: %s | %s: %s", TABLE_NAME_PG, before_ot, TABLE_CATALOGO, before_cat)

    start = time.time()

    if what in ("ot", "all"):
        logger.info("Iniciando carga OT (basegeneralcostos)...")
        upload_buffer_with_merge(include_catalogo=False)

    if what in ("catalogo", "cat", "all"):
        logger.info("Iniciando carga catalogo...")
        registros = upload_catalogo()
        logger.info("Catalogo subido: %s registros", registros)

    elapsed = time.time() - start
    after_ot = count_rows(engine, TABLE_NAME_PG)
    after_cat = count_rows(engine, TABLE_CATALOGO)

    logger.info("=" * 70)
    logger.info("PRUEBA COMPLETADA en %.2fs", elapsed)
    logger.info(
        "Despues - %s: %s | %s: %s",
        TABLE_NAME_PG,
        after_ot,
        TABLE_CATALOGO,
        after_cat,
    )
    logger.info("=" * 70)


if __name__ == "__main__":
    main()
