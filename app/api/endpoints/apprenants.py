from fastapi import APIRouter, HTTPException
from app.services.api_service import fetch_api_data
from app.core.config import settings

router = APIRouter()

@router.get("/")
async def get_apprenants():
    endpoint_apprenant = "/r/v1/formation-longue/apprenants?codesPeriode=2"
    url_apprenant = f"{settings.YPAERO_BASE_URL}{endpoint_apprenant}"
    headers = {
        "X-Auth-Token": settings.YPAERO_API_TOKEN,
        "Content-Type": "application/json"
    }
    try:
        # Utiliser un service pour appeler l'API et récupérer les données
        apprenants_data = await fetch_api_data(url_apprenant, headers)
        return apprenants_data
    except HTTPException as http_exc:
        # Vous pouvez gérer les erreurs spécifiques ici si nécessaire
        raise http_exc
    except Exception as exc:
        # Gestion des erreurs non spécifiques
        raise HTTPException(status_code=500, detail=str(exc))
