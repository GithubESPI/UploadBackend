from fastapi import APIRouter, HTTPException
from app.services.api_service import fetch_api_data
from app.core.config import settings

router = APIRouter()

@router.get("/")
async def get_groupes():
    endpoint_groupe = "/r/v1/formation-longue/groupes"
    url_groupe = f"{settings.YPAERO_BASE_URL}{endpoint_groupe}"
    headers = {
        "X-Auth-Token": settings.YPAERO_API_TOKEN,
        "Content-Type": "application/json"
    }
    try:
        groupes_data = await fetch_api_data(url_groupe, headers)
        return groupes_data
    except HTTPException as http_exc:
        raise http_exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))
