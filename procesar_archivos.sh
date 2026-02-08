#!/bin/bash

main() {
    # 1. Tu directorio objetivo
    local subdirectorio="./12_Licenciaturas_BIS" 
    local archivos=()

    # 2. Verificamos y llenamos el arreglo
    if [ -d "$subdirectorio" ]; then
        for ruta_archivo in "$subdirectorio"/*; do
            if [ -f "$ruta_archivo" ]; then
                # Guardamos solo el nombre (ej: "Sistemas.txt")
                archivos+=("$(basename "$ruta_archivo")")
            fi
        done
    else
        echo "Error: El directorio $subdirectorio no existe."
        return 1
    fi

    # 3. ¡AQUÍ ESTÁ LA MAGIA!
    # Llamamos a Python y le "lanzamos" toda la lista de archivos
    echo "Ejecutando Python con ${#archivos[@]} archivos..."
    python3 generar_mapas.py "${archivos[@]}"
}

main
