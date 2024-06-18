import httpx
from fastapi import HTTPException
import logging

# Configure the logger
logger = logging.getLogger(__name__)

async def fetch_api_data(url: str, headers: dict):
    logger.debug(f"Fetching data from {url} with headers {headers}")
    async with httpx.AsyncClient(follow_redirects=True) as client:  # Enable following redirects
        response = await client.get(url, headers=headers, timeout=60.0)
        if response.status_code != 200:
            logger.error(f"Failed to fetch data: {response.status_code} - {response.text}")
            raise HTTPException(status_code=response.status_code, detail=f"API call failed with status {response.status_code}")
        try:
            data = response.json()
            logger.debug(f"Fetched data: {data}")
            if isinstance(data, (list, dict)):
                return data
            else:
                logger.error("Data is not a list or dict")
                return None
        except ValueError as e:
            logger.error(f"Error parsing JSON: {str(e)}")
            return None
