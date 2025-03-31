import os
import json
import logging
from dataclasses import dataclass
from dotenv import load_dotenv

logger = logging.getLogger(__name__)

@dataclass
class Config:
    client_secret_path: str
    token_path: str
    folder_id: str

    @classmethod
    def from_env(cls) -> 'Config':
        load_dotenv()
        
        client_secret_path = os.getenv('MCPGD_CLIENT_SECRET_PATH')
        if not client_secret_path:
            raise ValueError("MCPGD_CLIENT_SECRET_PATH environment variable is required")
            
        token_path = os.getenv('MCPGD_TOKEN_PATH')
        if not token_path:
            home_dir = os.path.expanduser('~')
            token_path = os.path.join(home_dir, '.mcp_google_spreadsheet.json')
            
        folder_id = os.getenv('MCPGD_FOLDER_ID')
        if not folder_id:
            raise ValueError("MCPGD_FOLDER_ID environment variable is required")
            
        logger.info(json.dumps({
            "event": "config_loaded",
            "config": {
                "client_secret_path": client_secret_path,
                "token_path": token_path,
                "folder_id": folder_id,
                "client_secret_exists": os.path.exists(client_secret_path)
            }
        }))
        
        return cls(
            client_secret_path=client_secret_path,
            token_path=token_path,
            folder_id=folder_id
        ) 