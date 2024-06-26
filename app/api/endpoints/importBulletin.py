import base64
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from app.core.config import settings
import requests

router = APIRouter()

class DocumentModel(BaseModel):
    contenu: str
    nomDocument: str
    mimeType: str
    extension: str

@router.post("/import")
async def import_document(document: DocumentModel):
    endpoint = f"/r/v1/document/apprenant/69992/document?codeRepertoire=1000011"
    url = f"{settings.YPAERO_BASE_URL}{endpoint}"
    headers = {
        "X-Auth-Token": settings.YPAERO_API_TOKEN,
        "Content-Type": "application/json"
    }

    try:
        # Create the JSON payload
        payload = {
            "contenu": document.contenu,
            "nomDocument": document.nomDocument,
            "typeMime": document.mimeType,  # Assurez-vous que ce champ est correctement orthographié et accepté par l'API
            "extension": document.extension,
        }

        # Using the requests library to perform the POST request
        response = requests.post(url, headers=headers, json=payload)
        
        if response.status_code == 200:
            return {"message": "Document imported successfully."}
        else:
            raise HTTPException(status_code=response.status_code, detail=response.text)

    except HTTPException as http_exc:
        raise http_exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))
