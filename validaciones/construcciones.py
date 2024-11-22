import pandas as pd
from tkinter import messagebox
from datetime import datetime

class Construcciones:
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        
        
    def validar_construcciones_No_convencionales(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Construcciones'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_construcciones")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            
            resultados = []

            for index, row in df.iterrows():
                conv = row['ConvencionalNoConvencional']
                calificacion = row['calificacionNoConvencional']

                if conv == 'No Convencional' and pd.isna(calificacion) or calificacion=='':
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'secuencia':row['secuencia'],
                        'Tipo Contruccion':row['TipoConstruccion'],
                        'ConvencionalNoConvencional': row['ConvencionalNoConvencional'],
                        'calificacionNoConvencional': row['calificacionNoConvencional'],
                        'Observacion': 'Calificación no convencional es nula para Noconvencional',
                        'Nombre Hoja': 'Construcciones'
                    }
                    resultados.append(resultado)
                    
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                
                
            if resultados:
                df_resultado = pd.DataFrame(resultados)

                '''
                output_file = 'CONSTRUCCIONES_VALIDACION.xlsx'
                sheet_name = 'CONSTRUCCIONES_VALIDACION'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")

                
                
                messagebox.showinfo("Éxito", f"No convencionales. {len(resultados)} registros.")
                '''
            else:
                messagebox.showinfo("Información", "No se encontraron registros No convencionales.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def tipo_construccion_noconvencionales(self):
        
        archivo_excel=self.archivo_entry.get()
        nombre_hoja='Construcciones'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_construcciones")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            
            resultados = []

            for index, row in df.iterrows():
                TipoConstruccion = row['Tipo Construccion']
                NoConvensional = row['ConvencionalNoConvencional']

                if NoConvensional == 'No Convencional' and TipoConstruccion != '' :
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'secuencia':row['secuencia'],
                        'Tipo Contruccion':row['TipoConstruccion'],
                        'ConvencionalNoConvencional': row['ConvencionalNoConvencional'],
                        'calificacionNoConvencional': row['calificacionNoConvencional'],
                        'Observacion': 'TipoConstruccion debe ser vacio si es No convencional',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                
                
            if resultados:
                df_resultado = pd.DataFrame(resultados)

                '''
                output_file = 'TipoConstruccion.xlsx'
                sheet_name = 'TipoConstruccion'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")

                
                
                messagebox.showinfo("Éxito", f"TipoConstruccion no vacio en No convencionales. {len(resultados)} registros.")
                '''
            else:
                messagebox.showinfo("Información", "No se encontraron registros tipo construccion.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    
    
    def areaconstruida_mayora1000(self):
        
        archivo_excel=self.archivo_entry.get()
        nombre_hoja='Construcciones'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_construcciones")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            
            resultados = []

            for index, row in df.iterrows():
                AreaConstruida = row['AreaConstruida']
                

                if AreaConstruida >= 1000 :
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'secuencia':row['secuencia'],
                        'AreaConstruida': row['AreaConstruida'],
                        'Observacion': 'El area construida es mayor a 1000 mts',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                
                
            if resultados:
                df_resultado = pd.DataFrame(resultados)

                '''
                output_file = 'Areamayor1000.xlsx'
                sheet_name = 'Areamayor1000'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"AreaMayor a 1000. {len(resultados)} registros.")
                '''
            else:
                messagebox.showinfo("Información", "No se encontraron registros Areamayor1000.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def validar_secuencia_construcciones_vs_calificaciones(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return
        
        try:
            # Leer las hojas Construcciones y CalificacionesConstrucciones
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_calificaciones = pd.read_excel(archivo_excel, sheet_name='CalificacionesConstrucciones')

            print(f"Leyendo archivo: {archivo_excel}")
            print(f"Dimensiones de Construcciones: {df_construcciones.shape}")
            print(f"Dimensiones de CalificacionesConstrucciones: {df_calificaciones.shape}")

            # Extraer las secuencias de cada hoja
            secuencia_construcciones = set(df_construcciones['secuencia'].dropna())
            secuencia_calificaciones = set(df_calificaciones['Secuencia'].dropna())

            # Encontrar secuencias en Construcciones que no están en CalificacionesConstrucciones
            secuencias_faltantes = secuencia_construcciones - secuencia_calificaciones

            resultados = []
            for secuencia in secuencias_faltantes:
                resultado = {
                    'secuencia': secuencia,
                    'Observacion': 'Secuencia está en Construcciones pero no en CalificacionesConstrucciones',
                    'Nombre Hoja': 'Construcciones'
                }
                resultados.append(resultado)
                print(f"Secuencia faltante: {resultado}")
            '''
            # Si se encuentran errores, guardar los resultados en un archivo Excel
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'ERRORES_SECUENCIAS_CONSTRUCCIONES.xlsx'
                df_resultado.to_excel(output_file, sheet_name='ErroresSecuencias', index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} errores.")
            else:
                messagebox.showinfo("Sin errores", "Todas las secuencias en Construcciones están presentes en CalificacionesConstrucciones.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    
    def validar_edad_construccion(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Construcciones'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                edad_construccion = row.get('EdadConstruccion', None)

                # Verificar si 'EdadConstruccion' es <= 0
                if edad_construccion is not None and edad_construccion <= 0:
                    resultado = {
                        'NroFicha': row.get('NroFicha', 'Sin NroFicha'),
                        'Secuencia': row.get('Secuencia', 'Sin Secuencia'),
                        'EdadConstruccion': edad_construccion,
                        'Observacion': 'Edad de construcción inválida (<= 0)',
                    }
                    resultados.append(resultado)
                    print(f"Fila {index}: Agregado a resultados: {resultado}")

            print(f"Total de errores encontrados: {len(resultados)}")

            # Agregar resultados a la lista general
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def validar_porcentaje_construido(self):
        """
        Verifica que en la hoja 'Construcciones' no haya valores en la columna 'PorcentajeConstruido' 
        que sean iguales o menores a 0. Si los hay, genera un error por cada registro que cumple la condición.
        """
        archivo_excel = self.archivo_entry.get()
        
        
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Construcciones'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            
            # Filtrar registros donde 'PorcentajeConstruido' es menor o igual a 0
            errores = df_construcciones[df_construcciones['PorcentajeConstruido'] <= 0]
            
            resultados = []

            # Generar lista de errores
            for _, row in errores.iterrows():
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'PorcentajeConstruido': row['PorcentajeConstruido'],
                    'Observacion': 'El valor de PorcentajeConstruido es igual o menor a 0',
                    'Nombre Hoja': 'Construcciones'
                }
                resultados.append(resultado)
            '''
            
            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_PorcentajeConstruido_Construcciones.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros con PorcentajeConstruido igual o menor a 0.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron valores de PorcentajeConstruido igual o menor a 0 en la hoja 'Construcciones'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    
    def validar_construcciones_puntos(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return []

        try:
            # Leer la hoja 'Construcciones'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            
            # Lista para almacenar errores
            errores = []

            # Validar cada fila en la columna 'Puntos'
            for index, row in df_construcciones.iterrows():
                # Validar si 'Puntos' es nulo
                if pd.isnull(row['Puntos']):
                    errores.append({
                        'Puntos':row ['Puntos'],
                        'Observacion': "La columna 'Puntos' contiene valores nulos.",
                        'NroFicha': row['NroFicha'],
                        'Nombre Hoja': 'Construcciones'
                    })
                
                # Validar si 'Puntos' es menor a 1
                elif row['Puntos'] < 1:
                    errores.append({
                        'Puntos':row ['Puntos'],
                        'Observacion': "La columna 'Puntos' contiene valores menores a 1.",
                        'NroFicha': row['NroFicha'],
                        'Nombre Hoja': 'Construcciones'
                    })

            # Solo crear y guardar el DataFrame si hay errores
            '''
            
            if errores:
                df_errores = pd.DataFrame(errores)
                output_file = 'Errores_Construcciones_Puntos.xlsx'
                sheet_name = 'Errores Construcciones'
                df_errores.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(errores)} errores en la columna 'Puntos' de la hoja 'Construcciones'.")
            else:
                messagebox.showinfo("Validación completada", "No se encontraron errores en la columna 'Puntos' de la hoja 'Construcciones'.")
            '''
            return errores

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
        
    def validar_secuencia_calificaciones_vs_construcciones(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return
        
        try:
            # Leer las hojas Construcciones y CalificacionesConstrucciones
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_calificaciones = pd.read_excel(archivo_excel, sheet_name='CalificacionesConstrucciones')

            print(f"Leyendo archivo: {archivo_excel}")
            print(f"Dimensiones de Construcciones: {df_construcciones.shape}")
            print(f"Dimensiones de CalificacionesConstrucciones: {df_calificaciones.shape}")

            # Extraer las secuencias de cada hoja
            secuencia_construcciones = set(df_construcciones['secuencia'].dropna())
            secuencia_calificaciones = set(df_calificaciones['Secuencia'].dropna())

            # Encontrar secuencias en CalificacionesConstrucciones que no están en Construcciones
            secuencias_faltantes = secuencia_calificaciones - secuencia_construcciones

            resultados = []
            for secuencia in secuencias_faltantes:
                resultado = {
                    'secuencia': secuencia,
                    'Observacion': 'Secuencia está en CalificacionesConstrucciones pero no en Construcciones',
                    'Nombre Hoja': 'CalificacionesConstrucciones'
                }
                resultados.append(resultado)
                print(f"Secuencia faltante: {resultado}")
            '''
            
            # Si se encuentran errores, guardar los resultados en un archivo Excel
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'ERRORES_SECUENCIAS_CALIFICACIONES.xlsx'
                df_resultado.to_excel(output_file, sheet_name='ErroresSecuencias', index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} errores.")
            else:
                messagebox.showinfo("Sin errores", "Todas las secuencias en CalificacionesConstrucciones están presentes en Construcciones.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    
    def validar_secuencia_convencional(self):
        """
        Valida que si en la columna 'ConvencionalNoConvencional' de la hoja 'Construcciones' el valor es 'Convencional',
        al menos una secuencia de la hoja 'Construcciones' exista en el campo 'Secuencia' de la hoja 'ConstruccionesGenerales'.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer las hojas 'Construcciones' y 'ConstruccionesGenerales'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_generales = pd.read_excel(archivo_excel, sheet_name='ConstruccionesGenerales')

            # Filtrar las secuencias de 'Construcciones' donde 'ConvencionalNoConvencional' es 'Convencional'
            convencionales = df_construcciones[df_construcciones['ConvencionalNoConvencional'] == 'Convencional']
            
            # Obtener las secuencias únicas en 'ConstruccionesGenerales'
            secuencias_generales = set(df_generales['Secuencia'].dropna().unique())

            # Lista para almacenar los errores
            errores = []

            # Validar cada secuencia en 'convencionales' para asegurarse de que exista en 'ConstruccionesGenerales'
            for index, row in convencionales.iterrows():
                secuencia = row['secuencia']
                if secuencia not in secuencias_generales:
                    errores.append({
                        'NroFicha': row['NroFicha'],
                        'secuencia': secuencia,
                        'Observacion': "Secuencia 'Convencional' no encontrada en 'ConstruccionesGenerales'",
                        'Nombre Hoja': 'Construcciones'
                    })
            '''
            
            # Guardar los resultados en un archivo Excel si hay errores
            if errores:
                df_errores = pd.DataFrame(errores)
                output_file = 'Errores_Convencional_Secuencia.xlsx'
                df_errores.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(errores)} secuencias sin correspondencia en 'ConstruccionesGenerales'.")
            else:
                messagebox.showinfo("Sin errores", "Todas las secuencias 'Convencional' están presentes en 'ConstruccionesGenerales'.")
            '''
            return errores

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        

    def validar_secuencia_convencional_calificaciones(self):
        """
        Valida que si en la columna 'ConvencionalNoConvencional' de la hoja 'Construcciones' el valor es 'Convencional',
        al menos una secuencia de la hoja 'Construcciones' exista en el campo 'Secuencia' de la hoja 'ConstruccionesGenerales'.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer las hojas 'Construcciones' y 'ConstruccionesGenerales'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_generales = pd.read_excel(archivo_excel, sheet_name='CalificacionesConstrucciones')

            # Filtrar las secuencias de 'Construcciones' donde 'ConvencionalNoConvencional' es 'Convencional'
            convencionales = df_construcciones[df_construcciones['ConvencionalNoConvencional'] == 'Convencional']
            
            # Obtener las secuencias únicas en 'ConstruccionesGenerales'
            secuencias_generales = set(df_generales['Secuencia'].dropna().unique())

            # Lista para almacenar los errores
            errores = []

            # Validar cada secuencia en 'convencionales' para asegurarse de que exista en 'ConstruccionesGenerales'
            for index, row in convencionales.iterrows():
                secuencia = row['secuencia']
                if secuencia not in secuencias_generales:
                    errores.append({
                        'NroFicha': row['NroFicha'],
                        'secuencia': secuencia,
                        'Observacion': "Secuencia 'Convencional' no encontrada en 'CalificacionesConstrucciones'",
                        'Nombre Hoja': 'Construcciones'
                    })
            '''
            
            # Guardar los resultados en un archivo Excel si hay errores
            if errores:
                df_errores = pd.DataFrame(errores)
                output_file = 'Errores_Convencional_Secuencia.xlsx'
                df_errores.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(errores)} secuencias sin correspondencia en 'ConstruccionesGenerales'.")
            else:
                messagebox.showinfo("Sin errores", "Todas las secuencias 'Convencional' están presentes en 'ConstruccionesGenerales'.")
            '''
            return errores

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
        
    def validar_no_convencional_secuencia(self):
        """
        Valida que si en la columna 'ConvencionalNoConvencional' de la hoja 'Construcciones' el valor es 'No Convencional',
        la secuencia de esa fila no esté presente en la columna 'Secuencia' de la hoja 'CalificacionesConstrucciones'.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer las hojas 'Construcciones' y 'CalificacionesConstrucciones'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_calificaciones = pd.read_excel(archivo_excel, sheet_name='CalificacionesConstrucciones')

            # Filtrar las secuencias de 'Construcciones' donde 'ConvencionalNoConvencional' es 'No Convencional'
            no_convencionales = df_construcciones[df_construcciones['ConvencionalNoConvencional'] == 'No Convencional']

            # Obtener las secuencias únicas en 'CalificacionesConstrucciones'
            secuencias_calificaciones = set(df_calificaciones['Secuencia'].dropna().unique())

            # Lista para almacenar los errores
            errores = []

            # Validar cada secuencia en 'no_convencionales' para asegurarse de que no exista en 'CalificacionesConstrucciones'
            for index, row in no_convencionales.iterrows():
                secuencia = row['secuencia']
                if secuencia in secuencias_calificaciones:
                    errores.append({
                        'NroFicha': row['NroFicha'],
                        'secuencia': secuencia,
                        'Observacion': "Secuencia 'No Convencional' encontrada en 'CalificacionesConstrucciones'",
                        'Nombre Hoja': 'Construcciones'
                    })

            # Guardar los resultados en un archivo Excel si hay errores
            if errores:
                df_errores = pd.DataFrame(errores)
                output_file = 'Errores_No_Convencional_Secuencia.xlsx'
                df_errores.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(errores)} secuencias 'No Convencional' que no cumplen la condición.")
            else:
                messagebox.showinfo("Sin errores", "Todas las secuencias 'No Convencional' cumplen la condición en 'CalificacionesConstrucciones'.")

            return errores

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
        
    def validar_secuencias_repetidas(self):
        """
        Verifica si existen secuencias repetidas en las hojas 'Construcciones' y 'CalificacionesConstrucciones'.
        Genera un error si una secuencia se encuentra en ambas hojas.
        """
        archivo_excel = self.archivo_entry.get()
        
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer las hojas 'Construcciones' y 'CalificacionesConstrucciones'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_calificaciones = pd.read_excel(archivo_excel, sheet_name='CalificacionesConstrucciones')

            # Extraer las secuencias únicas de cada hoja
            secuencias_construcciones = df_construcciones['Secuencia'].dropna().unique()
            secuencias_calificaciones = df_calificaciones['Secuencia'].dropna().unique()

            # Encontrar secuencias repetidas en ambas hojas
            secuencias_repetidas = set(secuencias_construcciones) & set(secuencias_calificaciones)

            resultados = []
            for secuencia in secuencias_repetidas:
                resultado = {
                    'Secuencia': secuencia,
                    'Observacion': 'Secuencia duplicada en ambas hojas',
                    'Nombre Hoja': 'Construcciones y CalificacionesConstrucciones'
                }
                resultados.append(resultado)
                print(f"Secuencia duplicada encontrada: {resultado}")

            # Si se encuentran duplicados, guardarlos en un archivo Excel
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Secuencias_Duplicadas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(resultados)} secuencias duplicadas en ambas hojas.")
            else:
                messagebox.showinfo("Validación completada", "No se encontraron secuencias duplicadas en ambas hojas.")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []