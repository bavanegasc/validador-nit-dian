"""
Ejemplo de uso avanzado del Validador de NITs DIAN
"""

from validador_nit_dian import ValidadorNITDIAN
import os

def ejemplo_basico():
    """Ejemplo básico de validación"""
    print("\n" + "="*60)
    print("EJEMPLO 1: Validación Básica")
    print("="*60)
    
    validador = ValidadorNITDIAN(
        archivo_entrada="NITs_a_validar.xlsx",
        archivo_salida="Resultados_Basico.xlsx"
    )
    validador.ejecutar()


def ejemplo_con_archivos_multiples():
    """Ejemplo procesando múltiples archivos"""
    print("\n" + "="*60)
    print("EJEMPLO 2: Procesar Múltiples Archivos")
    print("="*60)
    
    archivos = [
        "NITs_grupo1.xlsx",
        "NITs_grupo2.xlsx",
        "NITs_grupo3.xlsx"
    ]
    
    for archivo in archivos:
        if os.path.exists(archivo):
            print(f"\nProcesando: {archivo}")
            validador = ValidadorNITDIAN(
                archivo_entrada=archivo,
                archivo_salida=f"Resultados_{archivo.replace('.xlsx', '')}.xlsx"
            )
            validador.ejecutar()
        else:
            print(f"⚠️ Archivo no encontrado: {archivo}")


def ejemplo_procesamiento_personalizado():
    """Ejemplo con procesamiento personalizado de resultados"""
    print("\n" + "="*60)
    print("EJEMPLO 3: Procesamiento Personalizado")
    print("="*60)
    
    validador = ValidadorNITDIAN(
        archivo_entrada="NITs_a_validar.xlsx",
        archivo_salida="Resultados_Personalizado.xlsx"
    )
    
    # Ejecutar validación
    validador.procesar_nits()
    
    # Procesar resultados personalizados
    encontrados = [r for r in validador.resultados if "Encontrado" in r['Estado']]
    no_encontrados = [r for r in validador.resultados if "Encontrado" not in r['Estado']]
    
    print(f"\n✓ NITs encontrados: {len(encontrados)}")
    print(f"✗ NITs no encontrados: {len(no_encontrados)}")
    
    # Mostrar primeros 5 encontrados
    print("\nPrimeros 5 NITs encontrados:")
    for resultado in encontrados[:5]:
        print(f"  - {resultado['NIT']}: {resultado['Nombre_DIAN']}")
    
    # Generar Excel con resultados
    validador.generar_excel_resultados()


def ejemplo_con_manejo_errores():
    """Ejemplo con manejo de errores mejorado"""
    print("\n" + "="*60)
    print("EJEMPLO 4: Manejo de Errores")
    print("="*60)
    
    archivo_entrada = "NITs_a_validar.xlsx"
    
    # Verificar que el archivo existe
    if not os.path.exists(archivo_entrada):
        print(f"❌ Error: El archivo '{archivo_entrada}' no existe")
        print("Por favor, crea un Excel con los NITs a validar")
        return
    
    try:
        validador = ValidadorNITDIAN(
            archivo_entrada=archivo_entrada,
            archivo_salida="Resultados_ConErrores.xlsx"
        )
        
        if validador.ejecutar():
            print("\n✅ Validación completada exitosamente")
        else:
            print("\n❌ Error durante la validación")
            
    except Exception as e:
        print(f"\n❌ Error inesperado: {str(e)}")


def generar_reporte_consolidado():
    """Genera un reporte consolidado de múltiples validaciones"""
    print("\n" + "="*60)
    print("EJEMPLO 5: Reporte Consolidado")
    print("="*60)
    
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    
    # Simular validación de múltiples archivos
    validador = ValidadorNITDIAN(
        archivo_entrada="NITs_a_validar.xlsx",
        archivo_salida="Reporte_Consolidado.xlsx"
    )
    
    validador.procesar_nits()
    
    # Crear libro con múltiples hojas
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Eliminar hoja vacía
    
    # Hoja 1: Todos los resultados
    ws_todos = wb.create_sheet("Todos los NITs", 0)
    ws_todos.append(['NIT', 'Nombre DIAN', 'Estado', 'Fila Original'])
    
    for resultado in validador.resultados:
        ws_todos.append([
            resultado['NIT'],
            resultado['Nombre_DIAN'],
            resultado['Estado'],
            resultado['Fila_Original']
        ])
    
    # Hoja 2: Solo encontrados
    ws_encontrados = wb.create_sheet("Encontrados", 1)
    ws_encontrados.append(['NIT', 'Nombre DIAN'])
    
    encontrados = [r for r in validador.resultados if "Encontrado" in r['Estado']]
    for resultado in encontrados:
        ws_encontrados.append([resultado['NIT'], resultado['Nombre_DIAN']])
    
    # Hoja 3: Solo no encontrados
    ws_no_encontrados = wb.create_sheet("No Encontrados", 2)
    ws_no_encontrados.append(['NIT', 'Estado'])
    
    no_encontrados = [r for r in validador.resultados if "Encontrado" not in r['Estado']]
    for resultado in no_encontrados:
        ws_no_encontrados.append([resultado['NIT'], resultado['Estado']])
    
    # Hoja 4: Resumen estadístico
    ws_resumen = wb.create_sheet("Resumen", 3)
    
    resumen_data = [
        ['Métrica', 'Cantidad'],
        ['Total de NITs', len(validador.resultados)],
        ['Encontrados', len(encontrados)],
        ['No Encontrados', len(no_encontrados)],
        ['Porcentaje Exitoso', f"{(len(encontrados)/len(validador.resultados)*100):.2f}%"]
    ]
    
    for row in resumen_data:
        ws_resumen.append(row)
    
    # Estilo para la hoja de resumen
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    for cell in ws_resumen[1]:
        cell.fill = header_fill
        cell.font = header_font
    
    wb.save("Reporte_Consolidado.xlsx")
    print("✓ Reporte consolidado generado: Reporte_Consolidado.xlsx")


if __name__ == "__main__":
    print("\n" + "="*60)
    print("EJEMPLOS DE USO - VALIDADOR NIT DIAN")
    print("="*60)
    
    print("\nSelecciona un ejemplo para ejecutar:")
    print("1. Validación Básica")
    print("2. Procesar Múltiples Archivos")
    print("3. Procesamiento Personalizado")
    print("4. Manejo de Errores")
    print("5. Reporte Consolidado")
    print("0. Salir")
    
    opcion = input("\nIngresa el número de la opción: ").strip()
    
    if opcion == "1":
        ejemplo_basico()
    elif opcion == "2":
        ejemplo_con_archivos_multiples()
    elif opcion == "3":
        ejemplo_procesamiento_personalizado()
    elif opcion == "4":
        ejemplo_con_manejo_errores()
    elif opcion == "5":
        generar_reporte_consolidado()
    elif opcion == "0":
        print("Saliendo...")
    else:
        print("Opción no válida")
