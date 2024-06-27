import base64
from fastapi import APIRouter, HTTPException, UploadFile, File, Form
from pydantic import BaseModel
from app.core.config import settings
import requests

router = APIRouter()

@router.post("/import")
async def import_document(file: UploadFile = File(...), nomDocument: str = Form(...), mimeType: str = Form(...), extension: str = Form(...)):
    endpoint = f"/r/v1/document/apprenant/69992/document?codeRepertoire=1000011"
    url = f"{settings.YPAERO_BASE_URL}{endpoint}"
    headers = {
        "X-Auth-Token": settings.YPAERO_API_TOKEN,
        "Content-Type": "application/json"
    }

    try:
        # Read the file content and encode it in base64
        file_content = await file.read()
        encoded_content = base64.b64encode(file_content).decode('utf-8')

        # Create the JSON payload
        payload = {
            "contenu": encoded_content,
            "nomDocument": nomDocument,
            "typeMime": mimeType,
            "extension": extension,
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
