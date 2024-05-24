from fastapi import APIRouter, HTTPException, Query
from app.services.api_service import fetch_api_data
from app.core.config import settings

router = APIRouter()

@router.get("/")
async def get_absences(date_deb: str = Query(..., regex="^\\d{2}-\\d{2}-\\d{4}$"), date_fin: str = Query(..., regex="^\\d{2}-\\d{2}-\\d{4}$")):
    endpoint_abs = f"/r/v1/absences/{date_deb}/{date_fin}"
    url_abs = f"{settings.YPAERO_BASE_URL}{endpoint_abs}"
    headers = {
        "X-Auth-Token": settings.YPAERO_API_TOKEN,
        "Content-Type": "application/json"
    }
    try:
        absences_data = await fetch_api_data(url_abs, headers)
        return absences_data
    except HTTPException as http_exc:
        raise http_exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))
