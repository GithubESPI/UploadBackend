from pydantic_settings import BaseSettings  # type: ignore
import os

class Settings(BaseSettings):
    PROJECT_NAME: str = "Upload de Bulletins"
    BASE_DIR: str = os.getcwd()
    UPLOAD_DIR: str = os.path.join(BASE_DIR, "uploads")
    OUTPUT_DIR: str = os.path.join(BASE_DIR, "template", "S1")
    TEMPLATE_FILE: str = os.path.join(BASE_DIR, "template", "modeleM1S1.docx")
    
    # M1-S1 excel empty
    M1_S1_MAPI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M1-S1", "M1-S1-MAPI.xlsx")
    M1_S1_MAGI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M1-S1", "M1-S1-MAGI.xlsx")
    M1_S1_MEFIM_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M1-S1", "M1-S1-MEFIM.xlsx")
    # M1-S1 excel not empty
    
    M1_S1_MAPI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M1-S1-MAPI.xlsx")
    M1_S1_MAGI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M1-S1-MAGI.xlsx")
    M1_S1_MEFIM_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M1-S1-MEFIM.xlsx")
    
    # M1-S1 excel bulletin
    M1_S1_MAPI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM1S1.docx")
    M1_S1_MAGI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM1S1.docx")
    M1_S1_MEFIM_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM1S1.docx")

    # M1-S2
    M1_S2_MAPI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M1-S2", "M1-S2-MAPI.xlsx")
    M1_S2_MAGI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M1-S2", "M1-S2-MAGI.xlsx")
    M1_S2_MEFIM_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M1-S2", "M1-S2-MEFIM.xlsx")
    
    # M1-S1 excel not empty
    
    M1_S2_MAPI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M1-S2-MAPI.xlsx")
    M1_S2_MAGI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M1-S2-MAGI.xlsx")
    M1_S2_MEFIM_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M1-S2-MEFIM.xlsx")
    
    # M1-S1 excel bulletin
    
    M1_S2_MAPI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM1S2.docx")
    M1_S2_MAGI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM1S2.docx")
    M1_S2_MEFIM_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM1S2.docx")

    # M2-S3
    M2_S3_MAPI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M2-S3", "M2-S3-MAPI.xlsx")
    M2_S3_MAGI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M2-S3", "M2-S3-MAGI.xlsx")
    M2_S3_MEFIM_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M2-S3", "M2-S3-MEFIM.xlsx")
    
    # M1-S1 excel not empty
    
    M2_S3_MAPI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M2-S3-MAPI.xlsx")
    M2_S3_MAGI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M2-S3-MAGI.xlsx")
    M2_S3_MEFIM_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M2-S3-MEFIM.xlsx")
    
    # M1-S1 excel bulletin
    
    M2_S3_MAPI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM2S3MAPI.docx")
    M2_S3_MAGI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM2S3.docx")
    M2_S3_MEFIM_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM2S3.docx")

    # M2-S4
    M2_S4_MAPI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M2-S4", "M2-S4-MAPI.xlsx")
    M2_S4_MAGI_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M2-S4", "M2-S4-MAGI.xlsx")
    M2_S4_MEFIM_TEMPLATE: str = os.path.join(BASE_DIR, "excel", "M2-S4", "M2-S4-MEFIM.xlsx")
    
    # M1-S1 excel not empty
    
    M2_S4_MAPI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M2-S4-MAPI.xlsx")
    M2_S4_MAGI_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M2-S4-MAGI.xlsx")
    M2_S4_MEFIM_TEMPLATE_NOT_EMPTY: str = os.path.join(BASE_DIR, "template", "S1", "M2-S4-MEFIM.xlsx")
    
    # M1-S1 excel bulletin
    
    M2_S4_MAPI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM2S4.docx")
    M2_S4_MAGI_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM2S4.docx")
    M2_S4_MEFIM_TEMPLATE_WORD: str = os.path.join(BASE_DIR, "template", "modeleM2S4.docx")

    # ECTS
    ECTS_JSON_PATH: str = os.path.join(BASE_DIR, "json", "ects.json")

    RELEVANT_GROUPS: list = [
        "N-M1 MAPI ALT 1", "P-M1 MAPI ALT 2", "L-M1 MAPI ALT 2", "MP-M1 MAPI ALT",
        "P-M1 MAPI ALT 5", "L-M1 MAPI ALT 1", "P-M1 MAPI ALT 1", "P-M1 MAPI ALT 3",
        "B-M1 MAPI ALT 1", "M-M1 MAPI ALT 1", "LI-M1 MAPI ALT", "N-M1 MAPI ALT 2",
        "M-M1 MAPI ALT 2", "P-M1 MAPI ALT 4", "B-M1 MAPI ALT 2", "MP-M1 MAPI ALT",
        "L-M1 MAPI ALT 3", "P-M1 MAGI ALT 1", "N-M1 MAGI ALT", "M-M1 MAGI ALT",
        "LI-M1 MAGI ALT", "B-M1 MAGI ALT", "MP-M1 MAGI ALT", "L-M1 MAGI ALT",
        "P-M1 MAGI ALT 2", "LI-M1 MAGI ALT", "P-M1 MAGI ALT 2", "M-M1 MIFIM ALT",
        "N-M1 MIFIM ALT", "P-M1 MIFIM ALT 1", "P-M1 MIFIM ALT 2", "P-M1 MIFIM ALT 3",
        "LI-M1 MIFIM ALT", "B-M1 MIFIM ALT", "MP-M1 MIFIM ALT", "L-M1 MIFIM ALT"
    ]
    RELEVANT_GROUPS_M2: list = [
        "L-M2 MAPI ALT 1", "N-M2 MAGI ALT", "P-M2 MAPI ALT 3", "B-M2 MAGI ALT",
        "P-M2 MAPI ALT 5", "N-M2 MAPI ALT 2", "B-M2 MAPI ALT 1", "P-M2 MAGI ALT 2",
        "M-M2 MAPI ALT 2", "P-M2 MAPI ALT 1", "M-M2 MAPI ALT 1", "L-M2 MAPI ALT 2",
        "P-M2 MAPI ALT 2", "P-M2 MAPI ALT 4", "N-M2 MAPI ALT 1", "L-M2 MAGI ALT",
        "P-M2 MAGI ALT 1", "M-M2 MAGI ALT", "LI-M2 MAPI ALT", "M-M2 MAPI ALT 3",
        "M-M2 2ESI ALT", "N-M2 2ESI ALT", "N-M2 MIFIM ALT", "P-M2 2ESI ALT",
        "P-M2 MIFIM ALT 1", "P-M2 MIFIM ALT 2", "P-M2 MIFIM ALT 3", "M-M2 MIFIM ALT",
        "MP-M2 MAGI ALT", "MP-M2 MAPI ALT 1", "MP-M2 MAPI ALT 2", "B-M2 2ESI ALT",
        "B-M2 MIFIM ALT", "B-M2 MAPI ALT 2", "L-M2 MIFIM ALT", "L-M2 2ESI ALT",
        "P-M2 MAPI RP", "P-M2 MIFIM RP", "P-M2 MAGI RP", "CA-M2 MIFIM TP", "CA-M2 MAPI TP",
    ]

    # Paramètres d'API externe
    YPAERO_BASE_URL: str
    YPAERO_API_TOKEN: str

    class Config:
        # Chargez les variables d'environnement à partir d'un fichier .env situé à la racine du projet.
        env_file = ".env"

# Instanciez les paramètres pour qu'ils soient importés et utilisés dans d'autres fichiers
settings = Settings()

