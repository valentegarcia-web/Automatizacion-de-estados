# Automatizacion de Estados de Cuenta

Herramienta para extraer datos de estados de cuenta en PDF (GBM / Prestadero) y consolidarlos con un archivo maestro anterior.

## Estructura del proyecto

```
├── main.py                  # Punto de entrada principal
├── extractor_gbm.py         # Extrae datos de PDFs → Reporte_Consolidado.xlsx
├── consolidador.py           # Actualiza el maestro anterior con datos extraidos
├── requirements.txt          # Dependencias de Python
├── PDFs_Origen/              # Colocar aqui los PDFs del mes
├── 00_Maestro_Anterior/      # Colocar aqui el Excel del mes anterior
├── Resultados_Excel/         # Aqui se genera el Reporte_Consolidado.xlsx
└── Estados de cuenta del mes/ # Aqui se genera el Excel consolidado final
```

## Uso local

```bash
# Instalar dependencias
pip install -r requirements.txt

# Ejecutar todo (extraccion + consolidacion)
python main.py

# Solo extraer datos de PDFs
python main.py --solo-extraer

# Solo consolidar con el maestro
python main.py --solo-consolidar
```

## Uso en GitHub (en linea)

1. Sube los PDFs a la carpeta `PDFs_Origen/`
2. Sube el Excel maestro anterior a `00_Maestro_Anterior/`
3. Haz push al repositorio
4. Ve a la pestana **Actions** en GitHub
5. Selecciona **"Procesar Estados de Cuenta"**
6. Click en **"Run workflow"** y elige el modo
7. Descarga los archivos generados desde **Artifacts**

## Que hace cada paso

### Paso 1: Extraccion (`extractor_gbm.py`)
- Lee cada PDF de `PDFs_Origen/`
- Detecta si es GBM Regular, Smart Cash o Prestadero
- Extrae: resumen, portafolio, deuda, movimientos
- Genera `Resultados_Excel/Reporte_Consolidado.xlsx` con una hoja por cliente

### Paso 2: Consolidacion (`consolidador.py`)
- Lee el archivo maestro de `00_Maestro_Anterior/`
- Cruza los datos extraidos con cada hoja del maestro
- Actualiza saldos, ganancias, compras, ventas, deuda
- Genera el nuevo Excel en `Estados de cuenta del mes/`
