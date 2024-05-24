import logging
from fastapi import FastAPI
from app.api.endpoints import apprenants, groupes, absences, uploads

# Configurer le logger
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = FastAPI()

try:
    app.include_router(apprenants.router, prefix="/apprenants", tags=["apprenants"])
    # logger.debug("Le routeur apprenants est inclus.")
except Exception as e:
    logger.exception("Erreur lors de l'inclusion du routeur apprenants: %s", e)

try:
    app.include_router(groupes.router, prefix="/groupes", tags=["groupes"])
    # logger.debug("Le routeur groupes est inclus.")
except Exception as e:
    logger.exception("Erreur lors de l'inclusion du routeur groupes: %s", e)

try:
    app.include_router(absences.router, prefix="/absences", tags=["absences"])
    # logger.debug("Le routeur absences est inclus.")
except Exception as e:
    logger.exception("Erreur lors de l'inclusion du routeur absences: %s", e)

try:
    app.include_router(uploads.router, prefix="", tags=["uploads"])
    # logger.debug("Le routeur uploads est inclus.")
except Exception as e:
    logger.exception("Erreur lors de l'inclusion du routeur uploads: %s", e)

#uvicorn app.main:app --reload --log-level debug
#uvicorn app.main:app --reload