class Configuracion:
    
    EXCEL_PATH = "Docs/Bitacora.xlsx"
    WORD_TEMPLATE = "Docs/Formato.docx"
    OUTPUT_DIR = "SALIDA"
    MES="05"
    FECHA_INICIO = "23-03-2026"
    FECHA_FIN = "27-03-2026"

    MAP_TITULOS = {
        "Administración de Servidores, Dominios, Sistemas Operativos": "Administración de Servidores, Dominios, Sistemas Operativos",
        "Administración Infraestructura Cloud": "Administración Infraestructura Cloud",
        "Monitoreo y Respaldos": "Monitoreo y Respaldos",
        "Administración Middleware": "Administración Middleware"
    }


    EQUIPOS_ORDEN = [
        "Administración de Servidores, Dominios, Sistemas Operativos",
        "Administración Infraestructura Cloud",
        "Monitoreo y Respaldos",
        "Administración Middleware"  # ← siempre último
    ]

    EQUIPOS = {
        "Administración de Servidores, Dominios, Sistemas Operativos": [
            "RodrigoLorca", "PabloSalazar"
        ],
        "Administración Infraestructura Cloud": [
            "GonzaloRodriguez", "CesarGonzalez"
        ],
            "Monitoreo y Respaldos": [
            "JPBuzeta", "CarlosJara", "CristiánVásquez"
        ],
        "Administración Middleware": [
            "OscarMora", "AmericoSaravia", "GabrielNuñez"
        ],
    }

    NOMBRES_PROFESIONALES = {
        "RodrigoLorca": "Rodrigo Lorca",
        "PabloSalazar": "Pablo Salazar",
        "GonzaloRodriguez": "Gonzalo Rodríguez",
        "CesarGonzalez": "César González",
        "OscarMora": "Oscar Mora",
        "AmericoSaravia": "Américo Saravia",
        "GabrielNuñez": "Gabriel Núñez",
        "JPBuzeta": "Juan Pablo Buzeta",
        "CarlosJara": "Carlos Jara",
        "CristiánVásquez": "Cristián Vásquez"
    }

    PROFESIONALES = [p for lista in EQUIPOS.values() for p in lista]

    MAPEO_EQUIPOS = {
        profesional: equipo
        for equipo, lista in EQUIPOS.items()
        for profesional in lista
    }
   
   