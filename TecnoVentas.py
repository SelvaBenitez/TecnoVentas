from retos import (
    cargar_datos_excel,
    ver_catalogo,
    filtrar_por_precio,
    buscar_por_id,
    exportar_informe_pdf
)


def menu_interactivo():
    datos = cargar_datos_excel()
    if not datos:
        return

    while True:
        print("MENÚ PRINCIPAL - SISTEMA TECNOVENTAS")
        print("1. Ver catálogo completo \n2. Filtrar por precio (> $100)\n3. Buscar por ID\n4. Exportar informe PDF\n5. Salir")
        opcion = input("Seleccione una opción (1-5): ").strip()

        if opcion == "1":
            ver_catalogo(datos)
        elif opcion == "2":
            filtrar_por_precio(datos, 100)
        elif opcion == "3":
            buscar_por_id(datos)
        elif opcion == "4":
            exportar_informe_pdf(datos)
        elif opcion == "5":
            print("Saliendo del sistema...")
            break
        else:
            print("\nOpción no válida. Ingrese un número entre 1 y 5.")
menu_interactivo()