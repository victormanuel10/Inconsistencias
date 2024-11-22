import pandas as pd
from tkinter import messagebox
from ValidacionesRPH.fichasrph import FichasRPH

class FichasRPHProcesador:
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        self.resultados_generales = []

    def agregar_resultados(self, resultados):
        """ Agrega los resultados de validaciones a la lista general. """
        if isinstance(resultados, list):
            for resultado in resultados:
                self.resultados_generales.append(resultado)
        elif isinstance(resultados, pd.DataFrame):
            self.resultados_generales.extend(resultados.to_dict(orient='records'))

    def validar_npn_caracteristica(self, row):
        """ Valida una fila para ver si el carácter 22 de Npn es '9' y CaracteristicaPredio es '2|RPH'. """
        npn = str(row.get('Npn', ''))  # Convertir Npn a string por si acaso
        caracteristica_predio = row.get('CaracteristicaPredio', '')

        # Verificar las condiciones
        if len(npn) >= 22 and npn[21] == '9' and caracteristica_predio == '2|RPH':
            return True
        return False

    def procesar_errores_rph(self):
        """ Verifica si hay registros RPH antes de ejecutar validaciones y genera un archivo consolidado. """
        
        # Cargar la hoja Fichas Prediales
        try:
            archivo_excel = self.archivo_entry if isinstance(self.archivo_entry, str) else self.archivo_entry.get()
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo: {str(e)}")
            return

        # Verificar si existen registros que cumplen con validar_npn_caracteristica
        registros_rph = df_fichas[df_fichas.apply(self.validar_npn_caracteristica, axis=1)]
        
        if registros_rph.empty:
            # Si no hay registros RPH, mostrar mensaje y terminar
            messagebox.showinfo("Sin registros RPH", "No se encontraron registros RPH.")
            return
        
        # Si existen registros válidos, continuar con las validaciones
        fichas_rph = FichasRPH(self.archivo_entry)
        
        # Validación de coeficiente de copropiedad
        '''
        
        self.agregar_resultados(fichas_rph.validar_coeficiente_copropiedad_por_npn())
        self.agregar_resultados(fichas_rph.validar_duplicados_npn())
        self.agregar_resultados(fichas_rph.edificio_en_cero_rph())
        self.agregar_resultados(fichas_rph.unidad_predial_en_cero())
        self.agregar_resultados(fichas_rph.validar_destino_economico())
        '''
        self.agregar_resultados(fichas_rph.validar_npn_y_caracteristica())
        self.agregar_resultados(fichas_rph.validar_npn_num_cedula())
        self.generar_archivo_errores()

    def generar_archivo_errores(self):
        """ Genera un archivo Excel con los errores recopilados. """
        errores_por_hoja = {}

        if self.resultados_generales:
            for resultado in self.resultados_generales:
                nombre_hoja = resultado.get('Nombre Hoja', 'Sin Nombre')  # Obtener el nombre de la hoja
                if nombre_hoja not in errores_por_hoja:
                    errores_por_hoja[nombre_hoja] = []  # Inicializa la lista para esa hoja
                errores_por_hoja[nombre_hoja].append(resultado)

            # Crear un archivo Excel con múltiples hojas
            with pd.ExcelWriter('ERRORES_FICHAS_RPH_CONSOLIDADOS.xlsx') as writer:
                for hoja, errores in errores_por_hoja.items():
                    df_resultado = pd.DataFrame(errores)
                    df_resultado.to_excel(writer, sheet_name=hoja, index=False)
                    print(f"Errores guardados en la hoja: {hoja}")

            messagebox.showinfo("Éxito", "Proceso completado. Se ha creado el archivo 'ERRORES_FICHAS_RPH_CONSOLIDADOS.xlsx'.")
        else:
            messagebox.showinfo("Sin errores", "No se encontraron errores en los archivos procesados.")