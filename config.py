# config.py
import os
from dotenv import load_dotenv

# Carga .env si existe (opcional)
load_dotenv()

class Config:
    # Ejemplo de URL: postgresql://postgres:mi_pass@localhost:5432/Danvilo
    SQLALCHEMY_DATABASE_URI = os.getenv(
        "DATABASE_URL",
        "postgresql://postgres:postgres123@localhost:5432/Danvilo"
    )
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SECRET_KEY = os.getenv("SECRET_KEY", "4404")
