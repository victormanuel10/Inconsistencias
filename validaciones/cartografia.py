import pandas as pd
from tkinter import messagebox

class Cartografia:
    
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        self.resultados_generales = []
        
        
    def validar_fichas_faltantes(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return []
        
        try:
            # Leer las hojas FICHAS y CARTOGRAFIA
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')
            df_cartografia = pd.read_excel(archivo_excel, sheet_name='CartografiaInformacionGrafica')

            # Convertir a string y eliminar espacios en blanco
            df_fichas['NroFicha'] = df_fichas['NroFicha'].astype(str).str.strip()
            df_cartografia['NroFicha'] = df_cartografia['NroFicha'].astype(str).str.strip()

            # Obtener los números de ficha únicos
            nro_ficha_fichas = set(df_fichas['NroFicha'].dropna().unique())
            nro_ficha_cartografia = set(df_cartografia['NroFicha'].dropna().unique())

            # Fichas que están en FICHAS pero no en CARTOGRAFIA
            fichas_faltantes_en_cartografia = nro_ficha_fichas - nro_ficha_cartografia

            resultados = []

            # Crear resultados para las fichas que faltan
            for nro_ficha in fichas_faltantes_en_cartografia:
                resultado = {
                    'NroFicha': nro_ficha,
                    'Observacion': 'NroFicha en FICHAS no está en CARTOGRAFIA',
                    'Nombre Hoja': 'CartografiaInformacionGrafica'
                }
                resultados.append(resultado)

            # Solo crear y guardar el DataFrame si hay resultados
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Fichas_Faltantes.xlsx'
                sheet_name = 'Fichas Faltantes'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"NroFicha en FICHAS no está en CARTOGRAFIA: {len(resultados)} registros.")
            else:
                messagebox.showinfo("Información", "No faltan fichas en Cartografía.")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
    
    def validar_cartografia_faltantes(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return []
        
        try:
            # Leer las hojas FICHAS y CARTOGRAFIA
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')
            df_cartografia = pd.read_excel(archivo_excel, sheet_name='CartografiaInformacionGrafica')

            # Convertir a string y eliminar espacios en blanco
            df_fichas['NroFicha'] = df_fichas['NroFicha'].astype(str).str.strip()
            df_cartografia['NroFicha'] = df_cartografia['NroFicha'].astype(str).str.strip()

            # Obtener los números de ficha únicos
            nro_ficha_fichas = set(df_fichas['NroFicha'].dropna().unique())
            nro_ficha_cartografia = set(df_cartografia['NroFicha'].dropna().unique())

            # Fichas que están en CARTOGRAFIA pero no en FICHAS
            fichas_faltantes_en_fichas = nro_ficha_cartografia - nro_ficha_fichas

            resultados = []

            # Crear resultados para las fichas que faltan
            for nro_ficha in fichas_faltantes_en_fichas:
                resultado = {
                    'NroFicha': nro_ficha,
                    'Observacion': 'NroFicha en CARTOGRAFIA no está en FICHAS',
                    'Nombre Hoja': 'Fichas'
                }
                resultados.append(resultado)

            # Solo crear y guardar el DataFrame si hay resultados
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Fichas_Faltantes.xlsx'
                sheet_name = 'Fichas Faltantes'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"NroFicha en CARTOGRAFIA no está en FICHAS: {len(resultados)} registros.")
            else:
                messagebox.showinfo("Información", "No faltan fichas de cartografia en Fichas.")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
        
    
    def validar_cartografia_columnas(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return []

        try:
            # Leer la hoja 'CartografiaInformacionGrafica'
            df_cartografia = pd.read_excel(archivo_excel, sheet_name='CartografiaInformacionGrafica')
            
            # Columnas a validar
            columnas_requeridas = [
                'IndicePlancha', 'Ventana', 'Escala', 'Vigencia', 'IndiceVuelo', 
                'Faja', 'Foto', 'VigenciaAero', 'ampliacion'
            ]
            
            # Lista para almacenar errores
            resultados = []

            # Validar columnas no nulas
            for columna in columnas_requeridas:
                if df_cartografia[columna].isnull().any():
                    resultados.append({
                        'Columna': columna,
                        'Observacion': f"La columna '{columna}' contiene valores nulos.",
                        'Nombre Hoja': 'CartografiaInformacionGrafica'
                    })

            # Validar que 'Vigencia' no sea menor a 1995
            if (df_cartografia['Vigencia'].astype(float) < 1995).any():
                resultados.append({
                    'Columna': 'Vigencia',
                    'Observacion': "La Vigencia es menor a 1995.",
                    'Nombre Hoja': 'CartografiaInformacionGrafica'
                })

            # Solo crear y guardar el DataFrame si hay errores
            if resultados:
                df_resultados = pd.DataFrame(resultados)
                output_file = 'Errores_Cartografia.xlsx'
                sheet_name = 'Errores Cartografia'
                df_resultados.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(resultados)} errores en la hoja 'CartografiaInformacionGrafica'.")
            else:
                messagebox.showinfo("Validación completada", "No se encontraron errores en la hoja 'CartografiaInformacionGrafica'.")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []