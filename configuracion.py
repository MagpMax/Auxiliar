class Configuracion:
    
   
    EXCEL_PATH = "Docs/Bitacora.xlsx"
    WORD_TEMPLATE = "Docs/Formato.docx"
    OUTPUT_DIR = "SALIDA"
    MES="05"
    
   
    FECHA_INICIO = "10-04-2026"
    FECHA_FIN = "08-05-2026"

    MAP_TITULOS2 = {
        "Administración de Servidores, Dominios, Sistemas Operativos": "Administración de Servidores, Dominios, Sistemas Operativos",
        "Administración Infraestructura Cloud": "Administración Infraestructura Cloud",
        "Administración Middleware": "Administración Middleware"
   }

    MAP_TITULOS = { "Monitoreo y Respaldos": "Monitoreo y Respaldos" }

    EQUIPOS_ORDEN2 = [
        "Administración de Servidores, Dominios, Sistemas Operativos",
        "Administración Infraestructura Cloud",
        "Administración Middleware"  # ← siempre último
    ]
 
    EQUIPOS_ORDEN = [     "Monitoreo y Respaldos"    ]   


    EQUIPOS2 = {
        "Administración de Servidores, Dominios, Sistemas Operativos": [
            "RodrigoLorca", "PabloSalazar"
        ],
        "Administración Infraestructura Cloud": [
            "GonzaloRodriguez", "CesarGonzalez"
        ],
        "Administración Middleware": [
            "OscarMora", "AmericoSaravia", "GabrielNuñez"
        ],
    }
    
    EQUIPOS = {
                    "Monitoreo y Respaldos": [
            "JPBuzeta", "CarlosJara", "CristiánVásquez","JCarlos"
        ],
    }
    

    NOMBRES_PROFESIONALES2 = {
       "RodrigoLorca": "Rodrigo Lorca",
        "PabloSalazar": "Pablo Salazar",
        "GonzaloRodriguez": "Gonzalo Rodríguez",
        "CesarGonzalez": "César González",
        "OscarMora": "Oscar Mora",
        "AmericoSaravia": "Américo Saravia",
        "GabrielNuñez": "Gabriel Núñez"
    }
    
    NOMBRES_PROFESIONALES = {
        "JPBuzeta": "Juan Pablo Buzeta",
        "CarlosJara": "Carlos Jara",
        "CristiánVásquez": "Cristián Vásquez",
        "JCarlos": "Juan Carlos Fuentes"
    }



    PROFESIONALES = [p for lista in EQUIPOS.values() for p in lista]
    

    MAPEO_EQUIPOS = {
        profesional: equipo
        for equipo, lista in EQUIPOS.items()
        for profesional in lista
    }
    

   
   