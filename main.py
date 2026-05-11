from ejecutor import ejecutor
import pytest
import sys


def ejecutar_tests():

    print("\n==============================")
    print("EJECUTANDO TESTS")
    print("==============================\n")

    resultado = pytest.main([

        "tests",
        "-v"
    ])

    # =====================================================
    # pytest devuelve:
    #
    # 0 = OK
    # 1 = tests fallaron
    # 2+ = errores ejecución
    # =====================================================

    if resultado != 0:

        print("\n❌ TESTS FALLARON")
        sys.exit(1)

    print("\n✅ TESTS OK\n")
    

def main():
    ejecutor.ejecutar()
    pass
      
if __name__ == "__main__":
    ejecutar_tests()
    main()