#!/usr/bin/env python3
"""
main.py
───────
Punto de entrada principal. Ejecuta:
  1. Extractor de datos de PDFs (genera Reporte_Consolidado.xlsx)
  2. Consolidador (actualiza el maestro anterior con los datos extraídos)

Uso:
  python main.py              # Ejecuta ambos pasos
  python main.py --solo-extraer   # Solo extrae datos de PDFs
  python main.py --solo-consolidar # Solo consolida con el maestro
"""

import sys
import os
import argparse

# Asegurar que el directorio de trabajo sea el del script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)


def main():
    parser = argparse.ArgumentParser(description="Automatizacion de Estados de Cuenta")
    parser.add_argument("--solo-extraer", action="store_true",
                        help="Solo extraer datos de PDFs")
    parser.add_argument("--solo-consolidar", action="store_true",
                        help="Solo consolidar con el maestro anterior")
    args = parser.parse_args()

    if args.solo_extraer:
        print("=" * 52)
        print("  PASO 1: Extrayendo datos de PDFs...")
        print("=" * 52)
        import extractor_gbm
        return

    if args.solo_consolidar:
        print("=" * 52)
        print("  PASO 2: Consolidando con maestro anterior...")
        print("=" * 52)
        import consolidador
        consolidador.main()
        return

    # Ambos pasos
    print("=" * 52)
    print("  PASO 1 de 2: Extrayendo datos de PDFs...")
    print("=" * 52)
    print()
    import extractor_gbm

    print()
    print("=" * 52)
    print("  PASO 2 de 2: Consolidando con maestro anterior...")
    print("=" * 52)
    print()
    import consolidador
    consolidador.main()

    print()
    print("=" * 52)
    print("  Proceso completo finalizado.")
    print("=" * 52)


if __name__ == "__main__":
    main()
