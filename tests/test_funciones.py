import sys
import os

import pandas as pd

sys.path.append(
    os.path.abspath(
        os.path.join(
            os.path.dirname(__file__),
            ".."
        )
    )
)

from funciones import Funciones


def test_filtrar_por_fecha():

    df = pd.DataFrame({

        "FECHA SOLICITUD": [

            pd.Timestamp("2026-04-01"),
            pd.Timestamp("2026-04-15"),
            pd.Timestamp("2026-05-01"),
            pd.Timestamp("2026-06-01")
        ]
    })

    resultado = Funciones.filtrar_por_fecha(df)

    assert len(resultado) == 2