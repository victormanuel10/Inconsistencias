import pandas as pd
from tkinter import messagebox

class FichasRPH:
    def __init__(self, archivo_entry):
        # archivo_entry can be either a string (file path) or tkinter.Entry
        self.archivo_entry = archivo_entry
        self.resultados_generales = []

    def obtener_archivo(self):
        """ Helper function to get the file path from either a tkinter.Entry or a string. """
        if isinstance(self.archivo_entry, str):
            return self.archivo_entry
        elif hasattr(self.archivo_entry, 'get'):
            return self.archivo_entry.get()
        else:
            return None

    def validar_coeficiente_copropiedad_por_npn(self):
        """ 
        Valida que la suma de CoeficienteCopropiedad para los primeros 22 dígitos de Npn 
        sea igual a 100 en la hoja 'Fichas Prediales'; si no, genera un error.
        """
        archivo_excel = self.obtener_archivo()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja específica "Fichas Prediales"
            df_fichas = pd.read_excel(archivo_excel, sheet_name='FichasPrediales')

            # Crear una columna 'Npn_22' con los primeros 22 caracteres de Npn
            df_fichas['Npn_22'] = df_fichas['Npn'].astype(str).str[:22]

            # Agrupar por 'Npn_22' y sumar 'CoeficienteCopropiedad'
            suma_coeficientes = df_fichas.groupby('Npn_22')['CoeficienteCopropiedad'].sum().reset_index()

            # Filtrar donde la suma no es 100
            errores = suma_coeficientes[suma_coeficientes['CoeficienteCopropiedad'] != 100]
            resultados = []

            # Para cada error, buscar el valor completo de 'Npn' original y agregarlo al resultado
            for _, row in errores.iterrows():
                npn_22 = row['Npn_22']
                coeficiente_suma = row['CoeficienteCopropiedad']
                
                # Obtener todos los valores 'Npn' completos que corresponden al 'Npn_22'
                npn_completos = df_fichas[df_fichas['Npn_22'] == npn_22]['Npn'].unique()
                
                for npn in npn_completos:
                    resultado = {
                        'Npn': npn,
                        'Suma CoeficienteCopropiedad': coeficiente_suma,
                        'Observacion': 'La suma de CoeficienteCopropiedad no es 100',
                        'Nombre Hoja': 'FichasPrediales'
                    }
                    resultados.append(resultado)

            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_CoeficienteCopropiedad_Npn_22_FichasPrediales.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros.")
            else:
                messagebox.showinfo("Sin errores", "Todos los coeficientes de copropiedad suman 100 en 'Fichas Prediales'.")
            
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
        
    def validar_duplicados_npn(self):
        """
        Verifica que en la hoja 'FichasPrediales' existan duplicados en los primeros 22 caracteres de Npn.
        Si no hay duplicados, genera un error.
        """
        archivo_excel = self.obtener_archivo()
        nombre_hoja = 'FichasPrediales'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido y especifica la hoja.")
            return []

        try:
            # Leer la hoja específica 'FichasPrediales'
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            # Crear una columna 'Npn_22' con los primeros 22 caracteres de Npn
            df['Npn_22'] = df['Npn'].astype(str).str[:22]

            # Contar ocurrencias de cada valor en 'Npn_22'
            conteo_npn = df['Npn_22'].value_counts()

            # Filtrar los valores de 'Npn_22' que no tienen duplicados
            sin_duplicados = conteo_npn[conteo_npn == 1].index.tolist()

            # Generar una lista de errores para los registros sin duplicados
            resultados = []
            if sin_duplicados:
                for npn_22 in sin_duplicados:
                    filas_error = df[df['Npn_22'] == npn_22]
                    for _, fila in filas_error.iterrows():
                        resultado = {
                            'NroFicha': fila['NroFicha'],
                            'Npn': fila['Npn'],
                            'Observacion': 'No existe ficha resumen 2 para predio con característica RPH  y parcelación',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Error agregado: {resultado}")

                # Guardar resultados en archivo si existen errores
                if resultados:
                    df_resultado = pd.DataFrame(resultados)
                    output_file = 'Errores_Duplicados_Npn_FichasPrediales.xlsx'
                    df_resultado.to_excel(output_file, index=False)
                    print(f"Archivo de errores guardado: {output_file}")
                    messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros sin duplicados.")
                else:
                    messagebox.showinfo("Sin errores", "Todos los Npn tienen duplicados en los primeros 22 dígitos.")
            
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    def edificio_en_cero_rph(self):
        """
        Valida que en la columna 'Npn' de la hoja 'FichasPrediales', el dígito 22 sea 8 o 9, y los dígitos 23 y 24 sean 00.
        Si se cumple esta condición, genera un error.
        """
        archivo_excel = self.obtener_archivo()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja específica 'FichasPrediales'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='FichasPrediales')
            
            resultados = []

            # Iterar sobre las filas para validar la condición en la columna 'Npn'
            for index, row in df_fichas.iterrows():
                npn = str(row['Npn']).zfill(24)  # Rellenar con ceros a la izquierda para asegurar longitud de 24
                
                # Verificar longitud mínima de 24 caracteres antes de aplicar la validación
                if len(npn) >= 24 and npn[21] in ['8', '9'] and npn[22:24] == '00':
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Npn': row['Npn'],
                        'Observacion': 'Edificio en cero para caracteristica RPH, parcelación.',
                        'Nombre Hoja': 'FichasPrediales'
                    }
                    resultados.append(resultado)

            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Digitos_Npn_FichasPrediales.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron registros que cumplan con la condición especificada en 'Npn'.")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    def unidad_predial_en_cero(self):
        """
        Valida que en la columna 'Npn' de la hoja 'FichasPrediales', si el dígito 22 es '8' o '9' y el último
        dígito es '0', entonces genera un error.
        """
        archivo_excel = self.obtener_archivo()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja específica 'FichasPrediales'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='FichasPrediales')
            
            resultados = []

            # Iterar sobre las filas para validar la condición en la columna 'Npn'
            for index, row in df_fichas.iterrows():
                npn = str(row['Npn'])  # Convertir a cadena para asegurar el acceso a los dígitos específicos
                
                # Verificar que tenga al menos 22 dígitos antes de acceder al índice 21 y al último
                if len(npn) >= 22:
                    # Condiciones: el dígito 22 (índice 21) es '8' o '9', y el último dígito es '0'
                    if npn[21] in ['8', '9'] and npn.endswith('0'):
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Npn': row['Npn'],
                            'Observacion': 'Unidad Predial en cero para RPH, parcelación',
                            'Nombre Hoja': 'FichasPrediales'
                        }
                        resultados.append(resultado)

            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'unidad_predial_en_cero.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros que cumplen con las condiciones.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron registros en 'Npn' que cumplan con las condiciones en 'FichasPrediales'.")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []    
        
    def validar_destino_economico(self):
        """
        Verifica que en la hoja 'FichasPrediales', si el 'Npn' tiene '00' en las posiciones 6 y 7
        y el 'DestinoEconomico' es uno de los valores especificados, se genera un error.
        """
        archivo_excel = self.obtener_archivo()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'FichasPrediales'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='FichasPrediales')

            # Lista de valores de DestinoEconomico para verificar
            destinos_invalidos = [
                "12|LOTE URBANIZADO NO CONSTRUIDO",
                "13|LOTE URBANIZABLE NO URBANIZADO",
                "14|LOTE NO URBANIZABLE"
            ]
            
            resultados = []

            # Iterar sobre cada fila para validar las condiciones
            for index, row in df_fichas.iterrows():
                npn = str(row['Npn'])  # Convertir 'Npn' a cadena para acceder a posiciones específicas
                destino_economico = row.get('DestinoEconomico', '')

                # Verificar que Npn tenga al menos 7 caracteres y cumpla con las condiciones
                if len(npn) >= 7 and npn[5:7] == "00" and destino_economico in destinos_invalidos:
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Npn': row['Npn'],
                        'DestinoEconomico': destino_economico,
                        'Observacion': 'El Npn tiene "00" en las posiciones 6 y 7 y DestinoEconomico es un lote no desarrollado',
                        'Nombre Hoja': 'FichasPrediales'
                    }
                    resultados.append(resultado)

            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Posiciones_Npn_y_DestinoEconomico_FichasPrediales.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros que cumplen con las condiciones.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron registros que cumplan con las condiciones en 'FichasPrediales'.")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    
    
    def validar_npn_y_caracteristica(self):
        """
        Valida en la hoja 'Fichas' los registros donde:
        - El 22.º dígito de 'Npn' es '9'.
        - Los primeros 22 dígitos de 'Npn' están duplicados.
        - La suma de los últimos cuatro dígitos (27.º a 30.º) es mayor a cero.
        - 'CaracteristicaPredio' debe coincidir con la del registro con suma cero.
        """
        archivo_excel = self.obtener_archivo()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Fichas'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')
            
            resultados = []

            # Filtrar los registros donde el 22.º dígito de 'Npn' es '9'
            df_fichas = df_fichas[df_fichas['Npn'].astype(str).str[21] == '9']
            
            # Agrupar por los primeros 22 dígitos de 'Npn' para encontrar duplicados
            df_fichas['Npn_22_digitos'] = df_fichas['Npn'].astype(str).str[:22]
            grupos_npn = df_fichas.groupby('Npn_22_digitos')

            for npn_22, grupo in grupos_npn:
                # Convertir los últimos 4 dígitos a enteros y calcular la suma
                grupo['Ultimos_4_digitos'] = grupo['Npn'].astype(str).str[-4:].astype(int)
                grupo['Suma_ultimos_4'] = grupo['Ultimos_4_digitos'].apply(lambda x: sum(int(d) for d in str(x)))

                # Identificar el registro con suma igual a cero, si existe
                referencia = grupo[grupo['Suma_ultimos_4'] == 0]
                if not referencia.empty:
                    caracteristica_referencia = referencia.iloc[0]['CaracteristicaPredio']

                    # Validar los otros registros en el grupo
                    for _, row in grupo.iterrows():
                        if row['Suma_ultimos_4'] > 0 and row['CaracteristicaPredio'] != caracteristica_referencia:
                            resultado = {
                                'NroFicha': row['NroFicha'],
                                'Npn': row['Npn'],
                                'CaracteristicaPredio': row['CaracteristicaPredio'],
                                'CaracteristicaPredioEsperada': caracteristica_referencia,
                                'Observacion': 'CaracteristicaPredio no coincide con la ficha resumen',
                                'Nombre Hoja': 'Fichas'
                            }
                            resultados.append(resultado)
            '''
            
            # Guardar los errores en un archivo Excel si existen
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Npn_CaracteristicaPredio_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(resultados)} errores en la validación de 'Npn' y 'CaracteristicaPredio'.")
            else:
                messagebox.showinfo("Validación completada", "No se encontraron errores en 'Npn' y 'CaracteristicaPredio'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    def validar_npn_num_cedula(self):
        """
        Valida en la hoja 'Fichas' los registros donde:
        - El 22.º dígito de 'Npn' es '9'.
        - Los primeros 22 dígitos de 'Npn' están duplicados.
        - La suma de los últimos cuatro dígitos (27.º a 30.º) es mayor a cero.
        - 'NumCedulaCatastral' de los registros con suma mayor a cero no debe coincidir con el del registro con suma cero.
        """
        archivo_excel = self.obtener_archivo()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Fichas'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')
            
            resultados = []

            # Filtrar los registros donde el 22.º dígito de 'Npn' es '9'
            df_fichas = df_fichas[df_fichas['Npn'].astype(str).str[21] == '9']
            
            # Agrupar por los primeros 22 dígitos de 'Npn' para encontrar duplicados
            df_fichas['Npn_22_digitos'] = df_fichas['Npn'].astype(str).str[:22]
            grupos_npn = df_fichas.groupby('Npn_22_digitos')

            for npn_22, grupo in grupos_npn:
                # Convertir los últimos 4 dígitos a enteros y calcular la suma
                grupo['Ultimos_4_digitos'] = grupo['Npn'].astype(str).str[-4:].astype(int)
                grupo['Suma_ultimos_4'] = grupo['Ultimos_4_digitos'].apply(lambda x: sum(int(d) for d in str(x)))

                # Identificar el registro con suma igual a cero, si existe
                referencia = grupo[grupo['Suma_ultimos_4'] == 0]
                if not referencia.empty:
                    referencia_cedula = referencia.iloc[0]['NumCedulaCatastral']
                    
                    # Validar los otros registros en el grupo
                    for _, row in grupo.iterrows():
                        if row['Suma_ultimos_4'] > 0 and row['NumCedulaCatastral'] == referencia_cedula:
                            resultado = {
                                'NroFicha': row['NroFicha'],
                                'Npn': row['Npn'],
                                'NumCedulaCatastral': row['NumCedulaCatastral'],
                                'Observacion': 'NumCedulaCatastral ya existe en ficha resumen',
                                'Nombre Hoja': 'Fichas'
                            }
                            resultados.append(resultado)

            # Guardar los errores en un archivo Excel si existen
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_NumCedulaCatastral_Npn_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(resultados)} errores en la validación de 'NumCedulaCatastral' y 'Npn'.")
            else:
                messagebox.showinfo("Validación completada", "No se encontraron errores en 'NumCedulaCatastral' y 'Npn'.")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []