def filtrar_archivo_por_fecha(archivo_entrada, archivo_salida, anio, mes):
    """
    Filtra un archivo de texto eliminando las líneas con una fecha anterior
    a la especificada por el usuario.

    Args:
        archivo_entrada (str): La ruta del archivo de texto de entrada.
        archivo_salida (str): La ruta donde se guardará el nuevo archivo depurado.
        anio (int): El año límite.
        mes (int): El mes límite.
    """
    try:
        # Combinamos el año y el mes del usuario para una fácil comparación numérica.
        fecha_limite_int = int(f"{anio}{mes:02d}")

        with open(archivo_entrada, 'r') as f_entrada, open(archivo_salida, 'w', encoding='utf-8') as f_salida:
            for linea in f_entrada:
                try:
                    # Dividimos la línea en columnas.
                    columnas = linea.split()

                    # Nos aseguramos de que la línea tiene al menos dos columnas.
                    if len(columnas) > 1:
                        # Extraemos la fecha de la segunda columna y eliminamos las comillas.
                        fecha_str = columnas[1].strip('"')
                        fecha_linea_int = int(fecha_str)

                        # Si la fecha de la línea es mayor o igual a la fecha límite, la escribimos en el nuevo archivo.
                        if fecha_linea_int >= fecha_limite_int:
                            f_salida.write(linea)
                    else:
                        # Si una línea no tiene el formato esperado, la escribimos tal cual en el archivo de salida.
                        f_salida.write(linea)
                except (ValueError, IndexError):
                    # Si ocurre un error al procesar una línea, la escribimos igualmente para no perder datos.
                    f_salida.write(linea)

        print(f"\n¡Proceso completado! El archivo depurado ha sido guardado como '{archivo_salida}'")

    except FileNotFoundError:
        print(f"\nError: El archivo de entrada '{archivo_entrada}' no fue encontrado.")
    except Exception as e:
        print(f"\nHa ocurrido un error inesperado: {e}")

if __name__ == "__main__":
    print("--- Aplicación para Depurar Archivo de Texto por Fecha ---")
    try:
        nombre_archivo_entrada = input("Introduce el nombre del archivo de texto de entrada (ej: input.txt): ")
        nombre_archivo_salida = input("Introduce el nombre para el nuevo archivo depurado (ej: output.txt): ")
        anio_usuario = int(input("Introduce el año de corte (formato yyyy): "))
        mes_usuario = int(input("Introduce el mes de corte (formato mm): "))

        if not (1 <= mes_usuario <= 12):
            print("\nError: El mes debe ser un número entre 1 y 12.")
        else:
            filtrar_archivo_por_fecha(nombre_archivo_entrada, nombre_archivo_salida, anio_usuario, mes_usuario)

    except ValueError:
        print("\nError: Por favor, introduce un año y un mes válidos (deben ser números enteros).")
