# -*- coding: utf-8 -*-
import pandas as pd
from tkinter import messagebox

class Ficha:
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        
        
    def terreno_cero(self):
        
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            
            print("Leyendo archivo: {}, Hoja: {}".format(archivo_excel, nombre_hoja))
            print(f"funcion: terreno_cero")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                valor_b = row['Npn']
                valor_p = row['AreaTotalTerreno']

                # Verificar si valor_b no es nulo o vacío, y si tiene al menos 22 caracteres
                if pd.notna(valor_b) and len(str(valor_b)) > 21:
                    valor_b_str = str(valor_b)  # Convertir el valor a cadena
                    print(f"Fila {index}: Valor B = '{valor_b_str}', condicion: {valor_b_str[21]}, Valor P = '{valor_p}'")

                    # Verificar las condiciones
                    if valor_b_str[21] == '0' and (valor_p == '0' or valor_p == 0):
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Observacion': 'Terreno en ceros para ficha que no es mejora',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                else:
                    print(f"Fila {index}: NumCedulaCatastral no tiene suficientes caracteres o es nulo.")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            
            df_resultado = pd.DataFrame(resultados)
            
            #messagebox.showinfo("Éxito", f"Proceso completado Terreno cero. con {len(resultados)} registros.")
            return resultados

        except Exception as e:
                print(f"Error: {str(e)}")
                messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    
    def terreno_null(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return []  # Devolver lista vacía si no se encuentra el archivo
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: terreno_null")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                valor_a = str(row['Npn'])  # Convertir a cadena por si acaso
                valor_b = row['AreaTotalTerreno']

            print(f"Fila {index}: Valor A = '{valor_a}'")

            # Verificar que valor_a tenga al menos 22 caracteres
            if len(valor_a) > 21:
                print(f"CARACTER22B : {valor_a[21]}")

                if valor_a[21] == '2' and pd.isna(valor_b):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'NumCedulaCatastral': row['NumCedulaCatastral'],
                        'Condicion de predio': valor_a[21],
                        'AreaTotalTerreno': valor_b,
                        'Observacion': 'Terreno nulo para condición de predio',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
            else:
                print(f"Fila {index}: Valor A tiene menos de 22 caracteres, se omite.")
                    
            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados) if resultados else pd.DataFrame(columns=[
                'NroFicha', 'NumCedulaCatastral', 'Condicion de predio', 'AreaTotalTerreno', 'Observacion'])

            # Guardar el resultado en un nuevo archivo Excel
            '''
            
            output_file = 'TERRENO_NULL.xlsx'
            sheet_name = 'TERRENO_NULL'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")
            messagebox.showinfo("Éxito",
                                f"Proceso completado Terreno null. con {len(resultados)} registros.")
         
            '''
            
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []  # Devolver lista vacía en caso de error
            
    def informal_matricula(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: informal_matricula")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                valor_a = str(row['MatriculaInmobiliaria']) if pd.notna(row['MatriculaInmobiliaria']) else ''
                valor_b = row['ModoAdquisicion']

                print(f"Fila {index}: Valor A = '{valor_a}', condicion: {valor_b}")

                # Verificar las condiciones: valor_b es '2|POSESIÓN' y valor_a no está vacío
                if valor_b == '2|POSESIÓN' and (valor_a != '' or pd.notna(valor_a)):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'ModoAdquisicion': row['ModoAdquisicion'],
                        'MatriculaInmobiliaria': valor_a,
                        'Observacion': 'Matricula invalida para posesión',
                        'Nombre Hoja': nombre_hoja
                         
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")
            
            
       
            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)
            '''
            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'INFORMA_MATRICULA.xlsx'
            sheet_name = 'INFORMAL_MATRICULA'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito", f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
            
            '''
            return resultados   
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def matricula_mejora(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: matricula_mejora")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                valor_a = str(row['Npn']) if pd.notna(row['Npn']) else ''
                valor_b = row['MatriculaInmobiliaria']

                print(f"Fila {index}: Valor A = '{valor_a}'")
                print(f"Fila {index}: Valor A = '{valor_b}'")
                # Verificar si valor_a tiene al menos 22 caracteres
                if len(valor_a) >= 22:
                    print(f"CARACTER22B : {valor_a[21]}")

                    # Verificar las condiciones
                    if valor_a[21] == '2' and pd.notna(valor_b) or valor_b=='':
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'NumCedulaCatastral': valor_a,
                            'Condicion de predio': valor_a[21],
                            'ModoAdquisicion': row['ModoAdquisicion'],
                            'MatriculaInmobiliaria': row['MatriculaInmobiliaria'],
                            'Observacion': 'Informalidad con matrícula',
                            'Nombre Hoja': nombre_hoja
                            
                        }
                        
                        resultados.append(resultado)
                        print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                else:
                    print(f"Fila {index} tiene una NumCedulaCatastral demasiado corta: '{valor_a}'")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)
            '''
            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'MATRICULA_MEJORA.xlsx'
            sheet_name = 'MATRICULA_MEJORA'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
            
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")  
            
    def circulo_mejora(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: circulo_mejora")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iloc[0:].iterrows():
                valor_a = row['Npn']
                valor_b = row['circulo']

            print(f"Fila {index}: Valor A = '{valor_a}'")

            # Verificar que 'valor_a' tiene al menos 22 caracteres
            if len(valor_a) > 21:
                print(f"CARACTER22B : {valor_a[21]}")

                # Verificar las condiciones
                if valor_a[21] == '2' and pd.notna(valor_b):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'NumCedulaCatastral': row['NumCedulaCatastral'],
                        'Condicion de predio': valor_a[21],
                        'ModoAdquisicion': row['ModoAdquisicion'],
                        'circulo': row['circulo'],
                        'Observacion': 'Informalidad con matrícula',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
            else:
                print(f"El valor de 'NumCedulaCatastral' en la fila {index} no tiene suficientes caracteres.")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)
            '''
            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'CIRCULO_MEJORA.xlsx'
            sheet_name = 'CIRCULO_MEJORA'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
            
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def tomo_mejora(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: tomo_mejora")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame, comenzando desde la fila 2
            for index, row in df.iloc[0:].iterrows():
                valor_a = row['Npn']
                valor_b = row['Tomo']

                print(f"Fila {index}: Valor A = '{valor_a}'")

                # Verificar que 'valor_a' tiene al menos 22 caracteres antes de acceder al índice 21
                if len(valor_a) > 21:
                    # Verificar las condiciones
                    if valor_a[21] == '2' and valor_b != 0:
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'NumCedulaCatastral': row['NumCedulaCatastral'],
                            'Condicion de predio': valor_a[21],
                            'ModoAdquisicion': row['ModoAdquisicion'],
                            'Tomo': row['Tomo'],
                            'Observacion': 'Informalidad con Tomo',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                else:
                    print(f"El valor de 'NumCedulaCatastral' en la fila {index} no tiene suficientes caracteres.")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)
            '''
            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'TOMO_MEJORA.xlsx'
            sheet_name = 'TOMO_MEJORA'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
            
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")         
    
    
    def modo_adquisicion_informal(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: modo_adquisicion_informal")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame, comenzando desde la fila 2
            for index, row in df.iloc[0:].iterrows():
                valor_a = row['Npn']
                valor_b = row['ModoAdquisicion']

                print(f"Fila {index}: Valor A = '{valor_a}'")

                # Verificar que 'valor_a' tiene al menos 22 caracteres antes de acceder al índice 21
                if len(valor_a) > 21:
                    # Verificar las condiciones
                    if valor_a[21] == '2' and valor_b != '2|POSESIÓN':
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Npn': row['Npn'],
                            'Condicion de predio': valor_a[21],
                            'ModoAdquisicion': row['ModoAdquisicion'],
                            'Observacion': 'La informalidad no puede tener modo de adquisición diferente a posesión',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                else:
                    print(f"El valor de 'NumCedulaCatastral' en la fila {index} no tiene suficientes caracteres.")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)
            '''
            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'MODO_ADQUISICION_INFORMAL.xlsx'
            sheet_name = 'MODO_ADQUISICION_INFORMAL'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
            
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}") 
            
            
            
    def ficha_repetida(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: ficha_repetida")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Verificar si hay duplicados en la columna 'NroFicha'
            duplicados = df[df.duplicated(subset='NroFicha', keep=False)]  # Detectar duplicados
            
            print(f"Total de registros duplicados encontrados: {duplicados.shape[0]}")

            if duplicados.shape[0] > 0:
                
                # Guardar los resultados en un nuevo archivo Excel
                output_file = 'FICHAS_REPETIDAS.xlsx'
                sheet_name = 'fichas_repetidas'
                duplicados.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                print(f"Dimensiones del DataFrame de duplicados: {duplicados.shape}")

                messagebox.showinfo("Éxito", f"Proceso completado. Se ha creado el archivo '{output_file}' con {duplicados.shape[0]} registros duplicados.")
                
                
            else:
                print("No se encontraron registros duplicados.")
                messagebox.showinfo("Información", "No se encontraron registros duplicados.")
            return duplicados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
        
    def rural_destino_invalido(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
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

            # Iterar sobre las filas del DataFrame, comenzando desde la fila 2
            for index, row in df.iloc[0:].iterrows():
                valor_b = str(row['NumCedulaCatastral'])
                valor_p = row['DestinoEcconomico']

                print(f"Fila {index}: Valor B = '{valor_b}',condicion:{valor_b[6]}, Valor P = '{valor_p}'")

                # Verificar las condiciones
                if valor_b[6] == '0' and (valor_p == '12|LOTE URBANIZADO NO CONSTRUIDO' or valor_p == '13|LOTE URBANIZABLE NO URBANIZADO' or valor_p == '14|LOTE NO URBANIZABLE'):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'DestinoEcconomico': row['DestinoEcconomico'],
                        'Observacion': 'En sector rural no es valido destinaciones 12,13 y 14',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")


            df_resultado = pd.DataFrame(resultados)
            '''
            output_file = 'rural_destino_invalido.xlsx'
            sheet_name = 'rural_destino_invalido'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            '''
            
            messagebox.showinfo("Éxito",
                                f"Proceso completado Rural destino invalido.' con {len(resultados)} registros.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def destino_economico_mayorcero(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        
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
                destino_economico = row['DestinoEcconomico']
                area_total_construida = row['AreaTotalConstruida']

                if destino_economico in ['12|Lote_Urbanizado_No_Construido', 
                                        '13|Lote_Urbanizable_No_Urbanizado', 
                                        '14|Lote_No_Urbanizable',
                                        '19|USO PUBLICO'] and area_total_construida > 0:
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'DestinoEconomico': destino_economico,
                        'AreaTotalConstruida': area_total_construida,
                        'Observacion': 'Destino económico no debe tener área construida mayor a cero',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index}: Agregado a resultados: {resultado}")

            print(f"Total de errores encontrados: {len(resultados)}")
            
            if resultados:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                '''
                output_file = 'ERRORES_DESTINO_ECONOMICO.xlsx'
                sheet_name = 'ErroresDestinoEconomico'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Proceso completado Destino Economico 12 13 14.con {len(resultados)} errores.")
            
                '''
                
            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros con errores.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def areaterrenocero(self):
        
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        
        try:
        
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                areatotalterreno = row['AreaTotalTerreno']
                area_total_construida = row['AreaTotalConstruida']

                if areatotalterreno == '' or areatotalterreno == 0 or pd.isna(areatotalterreno):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'AreaTotalTerreno':areatotalterreno,
                        'AreaTotalConstruida': area_total_construida,
                        'Observacion': 'El area total terreno es cero o null',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index}: Agregado a resultados: {resultado}")

            print(f"Total de errores encontrados: {len(resultados)}")
            
            if resultados:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                '''
                output_file = 'ERRORES_DESTINO_ECONOMICO.xlsx'
                sheet_name = 'ErroresDestinoEconomico'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Areasterreno es cero o null {len(resultados)} errores.")
            
                '''
                
            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros con errores.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
        
    def areaconstruccioncero(self):
        
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
        
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                areatotalterreno = row['AreaTotalTerreno']
                area_total_construida = row['AreaTotalConstruida']

                if area_total_construida <= 0 or pd.isna(areatotalterreno):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'AreaTotalTerreno':areatotalterreno,
                        'AreaTotalConstruida': area_total_construida,
                        'Observacion': 'Area total Construida es cero o null',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index}: Agregado a resultados: {resultado}")

            print(f"Total de errores encontrados: {len(resultados)}")
            
            if resultados:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                '''
                output_file = 'ERRORES_DESTINO_ECONOMICO.xlsx'
                sheet_name = 'ErroresDestinoEconomico'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"AreatotalConstruida es cero o null {len(resultados)} errores.")
            
                '''
                
            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros con errores.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
            
    def prediosindireccion(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        
        try:
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            resultados = []

            # Lista de palabras no permitidas
            palabras_no_permitidas = ['CALLE', 'calle', 'AVENIDA', 'avenida', 'KR','CRA','Cra', 'KL', 'CARRERA', 'Carrera', 'Diagonal']

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                DireccionReal = row['DireccionReal']

                # Verificar si la dirección está vacía o contiene palabras no permitidas
                if DireccionReal == '' or pd.isna(DireccionReal):
                    observacion = 'Predio sin dirección'
                elif any(palabra in DireccionReal for palabra in palabras_no_permitidas):
                    observacion = 'Contiene palabras no permitidas'
                else:
                    continue

                resultado = {
                    'NroFicha': row['NroFicha'],
                    'Direccion': row['DireccionReal'],
                    'Observacion': observacion,
                    'Nombre Hoja': nombre_hoja
                }
                resultados.append(resultado)
                print(f"Fila {index}: Agregado a resultados: {resultado}")

            print(f"Total de errores encontrados: {len(resultados)}")
            '''
            
            if resultados:
                
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                
                output_file = 'Direcciones.xlsx'
                sheet_name = 'Direcciones'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Predios sin dirección o con palabras no permitidas: {len(resultados)} errores.")
                
            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros con errores.")
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
        
    def validar_nrofichas_propietarios(self):
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return []
        
        try:
            # Leer las hojas Propietarios y Fichas
            df_propietarios = pd.read_excel(archivo_excel, sheet_name='Propietarios')
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            # Asegurarse de que las columnas 'NroFicha' sean de tipo string y sin espacios en blanco
            df_propietarios['NroFicha'] = df_propietarios['NroFicha'].astype(str).str.strip()
            df_fichas['NroFicha'] = df_fichas['NroFicha'].astype(str).str.strip()

            # Obtener los valores únicos y no nulos de 'NroFicha'
            nro_ficha_propietarios = set(df_propietarios['NroFicha'].dropna().unique())
            nro_ficha_fichas = set(df_fichas['NroFicha'].dropna().unique())

            # Verificar si hay NroFicha en Propietarios que no están en Fichas
            fichas_faltantes_en_fichas = nro_ficha_propietarios - nro_ficha_fichas

            resultados = []

            for nro_ficha in fichas_faltantes_en_fichas:
                resultado = {
                    'NroFicha': nro_ficha,
                    'Observacion': 'NroFicha en Propietarios no está en Fichas',
                    'Nombre Hoja': 'Propietarios'
                }
                resultados.append(resultado)
            '''
            
            if resultados:
                print("")
                #messagebox.showinfo("Resultados", f"Se encontraron {len(resultados)} fichas faltantes en Fichas.")
            else:
                messagebox.showinfo("Información", "No faltan fichas en Fichas desde Propietarios.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    def validar_nrofichas_faltantes(self):
        archivo_excel = self.archivo_entry.get()
        
        
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return []
        
        
        try:
            # Leer las hojas Propietarios y Fichas
            df_propietarios = pd.read_excel(archivo_excel, sheet_name='Propietarios')
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            # Asegurarse de que las columnas 'NroFicha' sean de tipo string y sin espacios en blanco
            df_propietarios['NroFicha'] = df_propietarios['NroFicha'].astype(str).str.strip()
            df_fichas['NroFicha'] = df_fichas['NroFicha'].astype(str).str.strip()

            # Obtener los valores únicos y no nulos de 'NroFicha'
            nro_ficha_propietarios = set(df_propietarios['NroFicha'].dropna().unique())
            nro_ficha_fichas = set(df_fichas['NroFicha'].dropna().unique())

            # Verificar si hay NroFicha en Fichas que no están en Propietarios
            fichas_faltantes_en_propietarios = nro_ficha_fichas - nro_ficha_propietarios

            resultados = []

            for nro_ficha in fichas_faltantes_en_propietarios:
                resultado = {
                    'NroFicha': nro_ficha,
                    'Observacion': 'NroFicha en Fichas no está en Propietarios',
                    'Nombre Hoja': 'Propietarios'
                }
                resultados.append(resultado)
            '''
            
            if resultados:
                
                
                df_resultado = pd.DataFrame(resultados)
                
                output_file = 'Fichas faltantes propietarios.xlsx'
                sheet_name = 'Fichas faltantes propietarios'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                messagebox.showinfo("Éxito", f"Fichas faltantes propietarios {len(resultados)} errores.")
                print(f"Archivo guardado: {output_file}")

            else:
                messagebox.showinfo("Información", "No faltan fichas en Propietarios.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")  # Python 3 - print con formato
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
    
    def porcentaje_litigiocero(self):
        
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        
        try:
        
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                PorcentajeLitigio = row['PorcentajeLitigio']
                

                if PorcentajeLitigio != 0 :
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Observacion': 'PorcentajeLitigio No puede ser diferente de cero',
                        'PorcentajeLitigio':row['PorcentajeLitigio'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index}: Agregado a resultados: {resultado}")

            print(f"Total de errores encontrados: {len(resultados)}")
            
            if resultados:
                # Crear un nuevo DataFrame con los resultados
                
                
                df_resultado = pd.DataFrame(resultados)
                '''
                output_file = 'PorcentajeLitigio.xlsx'
                sheet_name = 'PorcentajeLitigio'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                messagebox.showinfo("Éxito", f"PorcentajeLitigio diferente cero {len(resultados)} errores.")
                print(f"Archivo guardado: {output_file}")
                '''
                

            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros PorcentajeLitigiocero.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def validar_matriculas_duplicadas(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return []
        
        
        try:
            # Leer las hojas FICHAS y PROPIETARIOS
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')
            df_propietarios = pd.read_excel(archivo_excel, sheet_name='Propietarios')

            # Verificar que exista la columna NroFicha en FICHAS y Propietarios, y Documento en Propietarios
            if 'NroFicha' not in df_fichas.columns:
                messagebox.showerror("Error", "Falta la columna 'NroFicha' en la hoja FICHAS.")
                return []

            if 'NroFicha' not in df_propietarios.columns or 'Documento' not in df_propietarios.columns:
                messagebox.showerror("Error", "Faltan las columnas 'NroFicha' o 'Documento' en la hoja Propietarios.")
                return []

            # Eliminar filas donde 'MatriculaInmobiliaria' esté vacía en FICHAS
            df_fichas_limpio = df_fichas.dropna(subset=['MatriculaInmobiliaria'])

            # Identificar las filas duplicadas en 'MatriculaInmobiliaria'
            matriculas_duplicadas = df_fichas_limpio[df_fichas_limpio.duplicated('MatriculaInmobiliaria', keep=False)]

            resultados = []

            for matricula, group in matriculas_duplicadas.groupby('MatriculaInmobiliaria'):
                documentos = set() 
                for _, row in group.iterrows():
                    nro_ficha_fichas = row['NroFicha']

                    # Buscar el Documento en la hoja Propietarios usando NroFicha
                    propietario_fila = df_propietarios[df_propietarios['NroFicha'] == nro_ficha_fichas]

                    if not propietario_fila.empty:
                        documento_propietario = propietario_fila['Documento'].values[0]
                        documentos.add(documento_propietario) 
                        
                if len(documentos) > 1:
                    for _, row in group.iterrows():
                        nro_ficha_fichas = row['NroFicha']
                        resultado = {
                            'MatriculaInmobiliaria': matricula,
                            'NroFicha': nro_ficha_fichas,
                            'Observacion': 'Matrícula Inmobiliaria duplicada en FICHAS con Documentos diferentes en PROPIETARIOS',
                            'Nombre Hoja': 'Fichas'
                        }
                        resultados.append(resultado)
                '''
                df_resultado = pd.DataFrame(resultados)

                
                output_file = 'Matrícula Inmobiliaria.xlsx'
                sheet_name = 'Matrícula Inmobiliaria'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                '''
                
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    def validar_npn(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los errores
            errores = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                npn = row['Npn']  
                
                if isinstance(npn, str) and len(npn) >= 22:
                    if npn[21] == '0':
                        resto_caracteres = npn[22:]
                        if any(char != '0' for char in resto_caracteres):
                            # Si no son todos '0', agregar a los errores
                            error = {
                                'NroFicha': row['NroFicha'],
                                'Npn': npn,
                                'Observacion': 'El carácter 22 es 0 pero los caracteres restantes no son todos ceros',
                                'Nombre Hoja': nombre_hoja
                            }
                            errores.append(error)
                            print(f"Error encontrado en la fila {index}: {error}")
                else:
                    print(f"El valor de 'Npn' en la fila {index} no tiene suficientes caracteres.")

            print(f"Total de errores encontrados: {len(errores)}")

            # Si hay errores, guardarlos en un archivo Excel
            '''
            
            if errores:
                df_errores = pd.DataFrame(errores)
                output_file = 'ERRORES_NPN.xlsx'
                sheet_name = 'ErroresNpn'
                df_errores.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores", f"Se encontraron {len(errores)} errores y se guardaron en '{output_file}'.")
            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros con errores.")
            '''
            return errores

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def validar_npn14a17(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_npn")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                npn = str(row['Npn'])  # Convertir a cadena para asegurar el acceso por índice

                # Validar que el Npn tenga al menos 17 caracteres antes de acceder a las posiciones 14-17
                if len(npn) >= 17:
                    # Extraer los caracteres de las posiciones 14, 15, 16 y 17
                    subcadena_npn = npn[13:17]

                    # Verificar si las posiciones 14-17 son "0000"
                    if subcadena_npn == "0000":
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Npn': row['Npn'],
                            'Observacion': 'Npn contiene 0000 en las posiciones 14-17',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index}: Agregado a resultados: {resultado}")
                else:
                    print(f"El valor de 'Npn' en la fila {index} no tiene suficientes caracteres.")
            
            print(f"Total de errores encontrados: {len(resultados)}")
            '''
            
            if resultados:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                
                # Guardar el resultado en un archivo Excel
                output_file = 'ERRORES_NPN.xlsx'
                sheet_name = 'ErroresNPN'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Validación completada con {len(resultados)} errores.")
                
            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros con errores.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    
    def ultimo_digito(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_npn")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                npn = str(row['Npn'])  # Convertir a cadena para asegurar el acceso por índice

                # Validar que el Npn tenga al menos 17 caracteres antes de acceder a las posiciones 14-17
                if len(npn) >= 17:
                    # Extraer los caracteres de las posiciones 14-17
                
                    # Validar si el último dígito de 'Npn' es diferente de "0"
                    if npn[-1] != "0":
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'CaracteristicaPredio':row['CaracteristicaPredio'],
                            'Npn': row['Npn'],
                            'Observacion': 'El último dígito de Npn no es 0',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index}: Agregado a resultados: {resultado}")
                else:
                    print(f"El valor de 'Npn' en la fila {index} no tiene suficientes caracteres.")
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                    
                    # Guardar el resultado en un archivo Excel
                output_file = 'ERRORES_NPNUltimo.xlsx'
                sheet_name = 'ErroresNPN'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Validación completada con {len(resultados)} errores.")
                
            else:
                    print("No se encontraron errores.")
                    messagebox.showinfo("Información", "No se encontraron registros con errores.")
                    
            print(f"Total de errores encontrados: {len(resultados)}")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
                
    
    def validar_destino_economico_y_longitud_cedula(self):
        """
        Verifica que en la hoja 'FichasPrediales':
        1. Si el cuarto dígito de 'NumCedulaCatastral' es '2' y el 'DestinoEconomico' es uno de los valores especificados, se genera un error.
        2. Valida que todos los valores en 'NumCedulaCatastral' tengan exactamente 28 dígitos.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'FichasPrediales'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            # Lista de valores de DestinoEconomico para verificar
            destinos_invalidos = [
                "12|LOTE URBANIZADO NO CONSTRUIDO",
                "13|LOTE URBANIZABLE NO URBANIZADO",
                "14|LOTE NO URBANIZABLE"
            ]
            
            resultados = []

            # Validar longitud de 'NumCedulaCatastral'
            for index, row in df_fichas.iterrows():
                num_cedula_catastral = str(row['NumCedulaCatastral']).strip()  # Convertir a cadena

                # Validar longitud de 28 dígitos
                if len(num_cedula_catastral) != 28:
                    resultados.append({
                        'NroFicha': row['NroFicha'],
                        'NumCedulaCatastral': num_cedula_catastral,
                        'Observacion': 'NumCedulaCatastral no tiene 28 dígitos',
                        'Nombre Hoja': 'Fichas'
                    })

                # Validar DestinoEconomico si el cuarto dígito de NumCedulaCatastral es '2'
                destino_economico = row.get('DestinoEcconomico', '').strip()
                if num_cedula_catastral[3] == '2' and destino_economico in destinos_invalidos:
                    resultados.append({
                        'NroFicha': row['NroFicha'],
                        'NumCedulaCatastral': num_cedula_catastral,
                        'DestinoEconomico': destino_economico,
                        'Observacion': 'Destino Economico no valido para ficha Rural',
                        'Nombre Hoja': 'Fichas'
                    })
            '''
            
            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_NumCedulaCatastral_y_DestinoEconomico_FichasPrediales.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros que cumplen con las condiciones.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron registros que cumplan con las condiciones en 'FichasPrediales'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    
    def validar_direccion_referencia_y_nombre(self):
        """
        Verifica que en la hoja 'FichasPrediales', las columnas 'DireccionReferencia' y 'DireccionNombre' no contengan valores nulos.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'FichasPrediales'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            resultados = []

            # Validar si hay valores nulos en las columnas 'DireccionReferencia' y 'DireccionNombre'
            for index, row in df_fichas.iterrows():
                if pd.isnull(row['DireccionReferencia']):
                    resultados.append({
                        'NroFicha': row['NroFicha'],
                        'Columna': 'DireccionReferencia',
                        'Observacion': 'DireccionReferencia no está diligenciada',
                        'Nombre Hoja': 'FichasPrediales'
                    })

                if pd.isnull(row['DireccionNombre']):
                    resultados.append({
                        'NroFicha': row['NroFicha'],
                        'Columna': 'DireccionNombre',
                        'Observacion': 'DireccionNombre no está diligenciada',
                        'Nombre Hoja': 'FichasPrediales'
                    })
            '''
            
            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_DireccionReferencia_y_DireccionNombre_FichasPrediales.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros que cumplen con las condiciones.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron valores nulos en 'DireccionReferencia' o 'DireccionNombre' en 'FichasPrediales'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        

    def validar_tipo_documento(self):
        """
        Verifica que en la hoja 'Fichas', los valores en la columna 'TipoDocumento' no sean
        '10|CEDULA CIUDADANIA HOMBRE' o '10|CEDULA CIUDADANIA MUJER'.
        Si se encuentran estos valores, genera un error indicando que deben ser '10|CEDULA DE CIUDADANIA'.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Fichas'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Propietarios')

            # Valores no permitidos en 'TipoDocumento'
            valores_invalidos = [
                "10|CEDULA CIUDADANIA HOMBRE",
                "10|CEDULA CIUDADANIA MUJER"
            ]
            
            # Lista para almacenar los errores encontrados
            resultados = []

            # Validar cada fila en la columna 'TipoDocumento'
            for index, row in df_fichas.iterrows():
                tipo_documento = row.get('TipoDocumento', '')

                # Si 'TipoDocumento' contiene un valor no permitido
                if tipo_documento in valores_invalidos:
                    resultados.append({
                        'NroFicha': row['NroFicha'],
                        'TipoDocumento': tipo_documento,
                        'Observacion': "Debe ser '10|CEDULA DE CIUDADANIA'",
                        'Nombre Hoja': 'Propietarios'
                    })
            '''
            
            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_TipoDocumento_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(resultados)} registros con valores incorrectos en 'TipoDocumento'.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron valores incorrectos en la columna 'TipoDocumento' en la hoja 'Fichas'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
        
    def Validar_Longitud_NPN(self):
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'FichasPrediales'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

           
            
            resultados = []

            # Validar longitud de 'NumCedulaCatastral'
            for index, row in df_fichas.iterrows():
                Npn = str(row['Npn']).strip()  # Convertir a cadena

                # Validar longitud de 28 dígitos
                if len(Npn) != 30:
                    resultados.append({
                        'NroFicha': row['NroFicha'],
                        'Npn': Npn,
                        'Observacion': 'Npn no tiene 30 dígitos',
                        'Nombre Hoja': 'Fichas'
                    })

                # Validar DestinoEconomico si el cuarto dígito de NumCedulaCatastral es '2'
                
            '''
            
            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_NumCedulaCatastral_y_DestinoEconomico_FichasPrediales.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros que cumplen con las condiciones.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron registros que cumplan con las condiciones en 'FichasPrediales'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []   
    