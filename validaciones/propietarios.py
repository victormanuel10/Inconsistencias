import pandas as pd
from tkinter import messagebox
from datetime import datetime
from validaciones.ficha import Ficha
from validaciones.construcciones import Construcciones
from validaciones.califconstrucciones import CalificaionesConstrucciones
from validaciones.zonashomogeneas import ZonasHomogeneas
from validaciones.colindantes import Colindantes
from validaciones.cartografia import Cartografia
from ValidacionesRPH.erroresRPH import FichasRPHProcesador


class Propietarios:
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        self.resultados_generales = []
        
        
    def agregar_resultados(self, resultados):
        if isinstance(resultados, list):
            for resultado in resultados:
                self.resultados_generales.append(resultado)
        elif isinstance(resultados, pd.DataFrame):
            self.resultados_generales.extend(resultados.to_dict(orient='records'))
       
    def procesar_errores(self):
        construcciones = Construcciones(self.archivo_entry)
        #self.agregar_resultados(construcciones.validar_secuencia_convencional())
        #self.agregar_resultados(construcciones.validar_construcciones_No_convencionales())
        #self.agregar_resultados(construcciones.validar_secuencia_convencional_calificaciones())
        self.agregar_resultados(construcciones.validar_no_convencional_secuencia())
        '''
        
       
        self.agregar_resultados(construcciones.validar_construcciones_puntos())
        self.agregar_resultados(construcciones.validar_porcentaje_construido())
        self.agregar_resultados(construcciones.validar_edad_construccion())
        self.agregar_resultados(construcciones.validar_construcciones_No_convencionales())
        self.agregar_resultados(construcciones.areaconstruida_mayora1000())
        self.agregar_resultados(construcciones.tipo_construccion_noconvencionales())         
        self.agregar_resultados(construcciones.validar_secuencia_construcciones_vs_calificaciones())
        self.agregar_resultados(construcciones.validar_secuencia_calificaciones_vs_construcciones())
        '''
        ficha = Ficha(self.archivo_entry)
        #self.agregar_resultados(ficha.Validar_Longitud_NPN())
        #self.agregar_resultados(ficha.validar_tipo_documento())
        #self.agregar_resultados(ficha.validar_direccion_referencia_y_nombre())
        #self.agregar_resultados(ficha.validar_destino_economico_y_longitud_cedula())
        #self.agregar_resultados(ficha.ultimo_digito())
        #self.agregar_resultados(ficha.prediosindireccion())
        #self.agregar_resultados(ficha.validar_npn14a17())
        #self.agregar_resultados(ficha.validar_npn())
        #self.agregar_resultados(ficha.validar_nrofichas_faltantes())
        #self.agregar_resultados(ficha.validar_nrofichas_propietarios())
        #self.agregar_resultados(ficha.validar_matriculas_duplicadas())
        #self.agregar_resultados(ficha.porcentaje_litigiocero())
        #self.agregar_resultados(ficha.areaterrenocero())
        #self.agregar_resultados(ficha.areaconstruccioncero())
        ##self.agregar_resultados(ficha.destino_economico_mayorcero())
        #self.agregar_resultados(ficha.matricula_mejora())
        #self.agregar_resultados(ficha.terreno_cero())
        ##self.agregar_resultados(ficha.terreno_null())
        #self.agregar_resultados(ficha.informal_matricula())
        ##self.agregar_resultados(ficha.circulo_mejora())
        ##self.agregar_resultados(ficha.tomo_mejora())
        #self.agregar_resultados(ficha.modo_adquisicion_informal())
        #self.agregar_resultados(ficha.ficha_repetida())
        
        cartografia=Cartografia(self.archivo_entry)
        #self.agregar_resultados(cartografia.validar_fichas_faltantes())
        #self.agregar_resultados(cartografia.validar_cartografia_faltantes())
        #self.agregar_resultados(cartografia.validar_cartografia_columnas())
        colindantes=Colindantes(self.archivo_entry)
        #self.agregar_resultados(colindantes.validar_orientaciones_colindantes())
        zonashomogeneas= ZonasHomogeneas(self.archivo_entry)
        #self.agregar_resultados(zonashomogeneas.validar_tipo_zonas_homogeneas())
        
        
        #self.validar_tipo_documento()
        ##self.validar_documento_inicia_con_cero()
        ##self.validar_documento_sexo_masculino()
        ##self.validar_tipo_documento_sexo()
        ##self.validar_documento_sexo_femenino()
        #self.numerofallocero()
        #self.entidadvacio()
        #self.primer_apellido_blanco()
        #self.primer_nombre_blanco()
        #self.calidad_propietario_mun()
        #self.nit_diferente_mun()
        #self.derecho_diferente_cien()
        ##self.documento_blanco_cod_asig()
        #self.fecha_escritura_inferior()
        #self.fecha_escritura_mayor()
        
        calificonstrucciones= CalificaionesConstrucciones(self.archivo_entry)
        #self.agregar_resultados(calificonstrucciones.validar_sinCocina())
        #self.agregar_resultados(calificonstrucciones.conservacion_cubierta_bueno())
        #self.agregar_resultados(calificonstrucciones.validar_banios()) 
        #self.agregar_resultados(calificonstrucciones.Validar_armazon())
        #self.agregar_resultados(calificonstrucciones.Validar_fachada())
        
        errores_por_hoja = {}
        
        if self.resultados_generales:
            for resultado in self.resultados_generales:
                nombre_hoja = resultado.get('Nombre Hoja', 'Sin Nombre')  # Obtener el nombre de la hoja
                if nombre_hoja not in errores_por_hoja:
                    errores_por_hoja[nombre_hoja] = []  # Inicializa la lista para esa hoja
                errores_por_hoja[nombre_hoja].append(resultado)

            # Crear un archivo Excel con múltiples hojas
            with pd.ExcelWriter('ERRORES_CONSOLIDADOS.xlsx') as writer:
                for hoja, errores in errores_por_hoja.items():
                    df_resultado = pd.DataFrame(errores)
                    df_resultado.to_excel(writer, sheet_name=hoja, index=False)
                    print(f"Errores guardados en la hoja: {hoja}")

            messagebox.showinfo("Éxito", "Proceso completado. Se ha creado el archivo 'ERRORES_CONSOLIDADOS.xlsx'.")

        else:
            messagebox.showinfo("Sin errores", "No se encontraron errores en los archivos procesados.")
            
    def leer_archivo(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return None

        try:
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            return df
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error al leer el archivo: {str(e)}")
            return None
    '''
    
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
                    resultado = {
                    'NroFicha': row['NroFicha'],
                    'TipoDocumento': tipo_documento,
                    'Observacion': "Debe ser '10|CEDULA DE CIUDADANIA'",
                    'Nombre Hoja': 'Propietarios'
                    }
                    resultados.append(resultado)
            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_TipoDocumento_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(resultados)} registros con valores incorrectos en 'TipoDocumento'.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron valores incorrectos en la columna 'TipoDocumento' en la hoja 'Fichas'.")
                self.agregar_resultados(resultados)
            

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
    '''
    def validar_documento_inicia_con_cero(self):
        """
        Verifica que en la hoja 'Propietarios' no haya valores en la columna 'Documento' que inicien con '0'.
        Si los hay, genera un error por cada registro que cumple la condición.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Propietarios'
            df_propietarios = pd.read_excel(archivo_excel, sheet_name='Propietarios')
            
            # Filtrar los documentos que inician con '0'
            errores = df_propietarios[df_propietarios['Documento'].astype(str).str.startswith('0')]
            
            resultados = []

            # Generar una lista de errores
            for _, row in errores.iterrows():
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'Documento': row['Documento'],
                    'Observacion': 'El documento inicia con "0"',
                    'Nombre Hoja': 'Propietarios'
                }
                resultados.append(resultado)
            '''
            
            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Documento_Inicia_Con_Cero_Propietarios.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros con Documento que inicia con '0'.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron Documentos que inicien con '0' en la hoja 'Propietarios'.")
            '''
            self.agregar_resultados(resultados)

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    def validar_documento_sexo_femenino(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_documento_sexo")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                tipo_documento = row['TipoDocumento']
                documento = row['Documento']
                sexo = row['Sexo']

                # Intentar convertir el valor de 'Documento' a entero
                try:
                    documento = int(documento)
                except ValueError:
                    # Si no se puede convertir a entero, saltar la fila
                    print(f"Fila {index}: El valor del documento no es un número válido. Saltando fila.")
                    continue

                # Verificar si el Tipo de Documento es '10|CEDULA DE CIUDADANIA'
                if tipo_documento == '10|CEDULA DE CIUDADANIA':
                    # Validar que el Documento esté fuera del rango [20000000, 70000000] y que el Sexo sea 'F|FEMENINO'
                    if (documento <= 20000000 or documento >= 70000000) and sexo == 'F|FEMENINO':
                        resultado = {
                            'NroFicha': row['NroFicha'],  # Suponiendo que existe esta columna
                            'TipoDocumento': row['TipoDocumento'],
                            'Documento': row['Documento'],
                            'Sexo': row['Sexo'],
                            'Observacion': 'Documento fuera del rango para Sexo Femenino',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index}: Agregado a resultados: {resultado}")
            
            print(f"Total de errores encontrados: {len(resultados)}")
            '''
            
            if resultados:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                
                # Guardar el resultado en un archivo Excel
                output_file = 'ERRORES_DOCUMENTO_SEXO.xlsx'
                sheet_name = 'ErroresDocumentoSexo'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Validación completada con {len(resultados)} errores.")
                
            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros con errores.")
            '''
            self.agregar_resultados(resultados)
            

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")

        
   

    def validar_tipo_documento_sexo(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_tipo_documento_sexo")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                tipo_documento = row['TipoDocumento']
                sexo = row['Sexo']

                # Verificar si el Tipo de Documento es '3|NIT'
                if tipo_documento == '3|NIT':
                    # Validar que el Sexo sea 'N|NO BINARIO'
                    if sexo != 'N|NO BINARIO':
                        resultado = {
                            'NroFicha': row['NroFicha'],  # Suponiendo que existe esta columna
                            'TipoDocumento': row['TipoDocumento'],
                            'Sexo': row['Sexo'],
                            'Observacion': 'El tipo de documento es 3|NIT, pero el sexo no es N|NO BINARIO',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index}: Agregado a resultados: {resultado}")
            
            print(f"Total de errores encontrados: {len(resultados)}")
            '''
            if resultados:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                
                # Guardar el resultado en un archivo Excel
                output_file = 'ERRORES_TIPO_DOCUMENTO_SEXO.xlsx'
                sheet_name = 'ErroresTipoDocumentoSexo'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Validación completada con {len(resultados)} errores.")
               
            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros con errores.")
            '''
            self.agregar_resultados(resultados)
            

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")    
    
    def validar_documento_sexo_masculino(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_documento_sexo_masculino")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                tipo_documento = row['TipoDocumento']
                documento = row['Documento']
                sexo = row['Sexo']

                # Intentar convertir el valor de 'Documento' a entero
                try:
                    documento = int(documento)
                except ValueError:
                    # Si no se puede convertir a entero, saltar la fila
                    print(f"Fila {index}: El valor del documento no es un número válido. Saltando fila.")
                    continue

                # Verificar si el Tipo de Documento es '10|CEDULA DE CIUDADANIA'
                if tipo_documento == '10|CEDULA DE CIUDADANIA':
                    # Validar si el Documento está entre 20000000 y 69999999, y el Sexo es 'M|MASCULINO'
                    if 20000000 <= documento <= 69999999 and sexo == 'M|MASCULINO':
                        resultado = {
                            'NroFicha': row['NroFicha'],  # Suponiendo que existe esta columna
                            'TipoDocumento': row['TipoDocumento'],
                            'Documento': row['Documento'],
                            'Sexo': row['Sexo'],
                            'Observacion': 'Documento en rango para Cédula de Ciudadanía y Sexo Masculino',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index}: Agregado a resultados: {resultado}")
            
            print(f"Total de errores encontrados: {len(resultados)}")
            '''
            if resultados:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                
                # Guardar el resultado en un archivo Excel
                output_file = 'ERRORES_DOCUMENTO_SEXO_MASCULINO.xlsx'
                sheet_name = 'ErroresDocumentoSexoMasculino'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Validación completada con {len(resultados)} errores.")
            
            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros con errores.")
            '''
            self.agregar_resultados(resultados)
            

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    
    

        
        
    def primer_apellido_blanco(self):
        df = self.leer_archivo()
        if df is None:
            return
        
        
        resultados = []
        for index, row in df.iterrows():
            valor_a = row['PrimerApellido']
            valor_b = row['TipoDocumento']

            print(f"Fila {index}: Valor A = '{valor_a}'")

            if valor_b != '3|NIT' and (pd.isna(valor_a) or valor_a == ''):
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'Documento': row['Documento'],
                    'PrimerNombre': row['PrimerNombre'],
                    'SegundoNombre': row['SegundoNombre'],
                    'PrimerApellido': row['PrimerApellido'],
                    'SegundoApellido': row['SegundoApellido'],
                    'Observacion': 'Primer apellido en blanco'
                }
                resultados.append(resultado)
                self.agregar_resultados(resultados)
                print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

        df_resultado = pd.DataFrame(resultados)
        '''
        output_file = 'PRIMER_APELLIDO_BLANCO.xlsx'
        sheet_name = 'PRIMER_APELLIDO'
        df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
        print(f"Archivo guardado: {output_file}")
        messagebox.showinfo("Éxito",
                            f"Proceso completado Primer Apellido. con {len(resultados)} registros.")
        '''
        

        
        print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")
        
    def primer_nombre_blanco(self): 
        df = self.leer_archivo()
        if df is None:
            return
       
        resultados = []
        for index, row in df.iterrows():
            valor_a = row['PrimerNombre']
            valor_b = row['TipoDocumento']

            print(f"Fila {index}: Valor A = '{valor_a}'")

            if valor_b != '3|NIT' and (valor_a == '' or pd.isna(valor_a)):
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'Documento': row['Documento'],
                    'PrimerNombre': row['PrimerNombre'],
                    'SegundoNombre': row['SegundoNombre'],
                    'PrimerApellido': row['PrimerApellido'],
                    'SegundoApellido': row['SegundoApellido'],
                    'Observacion': 'Primer nombre en blanco'
                }
                resultados.append(resultado)
                self.agregar_resultados(resultados)
                print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

        df_resultado = pd.DataFrame(resultados)
        '''
        output_file = 'PRIMER_NOMBRE_BLANCO.xlsx'
        sheet_name = 'PRIMER_NOMBRE'
        df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
        print(f"Archivo guardado: {output_file}")
        
        
        

        messagebox.showinfo("Éxito",
                            f"Proceso completado PRIMER_NOMBRE. con {len(resultados)} registros.")
        '''
        print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")
        
        
    def calidad_propietario_mun(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'
        df = self.leer_archivo()
        if df is None:
            return
        

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: calidad_propietario_mun")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                valor_a = row['CalidadPropietario']
                valor_b = row['TipoDocumento']

                print(f"Fila {index}: Valor A = '{valor_a}'")

                # Verificar las condiciones
                if valor_a != '4|MUNICIPAL' and valor_b == '3|NIT':
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'TipoDocumento': row['TipoDocumento'],
                        'Documento': row['Documento'],
                        'CalidadPropietario': row['CalidadPropietario'],
                        'RazonSocial': row['RazonSocial'],
                        'Observacion': 'Calidad del propietario diferente para NIT del Municipio',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    # Solo se agrega el resultado actual, no toda la lista
                    self.agregar_resultados([resultado])
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)

            # Guardar el resultado en un nuevo archivo Excel
            '''
            output_file = 'CALIDAD_PROP_MUN.xlsx'
            sheet_name = 'CALIDAD_PROP_MUN'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            

            

            messagebox.showinfo("Éxito", f"Proceso completado Calidad prop mun. con {len(resultados)} registros.")
            '''
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def nit_diferente_mun(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'
        df = self.leer_archivo()
        if df is None:
            return
        

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: nit_diferente_mun")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                valor_a = row['CalidadPropietario']
                valor_b = row['TipoDocumento']

                print(f"Fila {index}: Valor A = '{valor_a}'")

                # Verificar las condiciones
                if valor_a == '4|MUNICIPAL' and valor_b != '3|NIT':
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'TipoDocumento': row['TipoDocumento'],
                        'Documento': row['Documento'],
                        'CalidadPropietario': row['CalidadPropietario'],
                        'RazonSocial': row['RazonSocial'],
                        'Observacion': 'tipo de documento diferente para nit del municipio',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    self.agregar_resultados(resultados)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)

            '''
            output_file = 'NIT_DIFERENTE_MUN.xlsx'
            sheet_name = 'NIT_DIFERENTE_MUN'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            
            

            messagebox.showinfo("Éxito",
                                f"Proceso completado Nit diferente num. con {len(resultados)} registros.")
            ''' 
            
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def derecho_diferente_cien(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: derecho_diferente_cien")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Agrupar el DataFrame por 'NroFicha'
            grouped = df.groupby('NroFicha')

            for name, group in grouped:
                valor_b_sum = group['Derecho'].sum()

                # Si la suma de 'Derecho' no es 100, guardar los valores
                if round(valor_b_sum, 3) != 100:
                    print(f"suma de derechos no es 100: {valor_b_sum}")
                    for _, row in group.iterrows():
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'TipoDocumento': row['TipoDocumento'],
                            'Documento': row['Documento'],
                            'Derecho': row['Derecho'],
                            'Observacion': 'Porcentaje de dominio incompleto diferente a cero, falta: ' + str(100 - valor_b_sum),
                            'Nombre Hoja': nombre_hoja
                        
                        }
                        resultados.append(resultado)
                        self.agregar_resultados(resultados)
                        print(f"Fila {_} agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")
            if resultados:
            
                fila_vacia = {key: '' for key in resultados[0].keys()}
                resultados.append(fila_vacia)
            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)

            # Guardar el resultado en un nuevo archivo Excel
            '''
                output_file = 'DERECHO_DIFERENTE_CIEN.xlsx'
                sheet_name = 'DERECHO_DIFERENTE_CIEN'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito",
                f"Proceso completado Derecho dirente cien. con {len(resultados)} registros.")
            
            '''            
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def documento_blanco_cod_asig(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: documento_blanco_cod_asig")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iloc[0:].iterrows():
                valor_a = row['TipoDocumento']
                valor_b = row['Documento']

                print(f"Fila {index}: Valor A = '{valor_a}', Valor B = '{valor_b}'")

                # Verificar las condiciones
                if valor_a == '8|CODIGO ASIGNADO' and pd.notna(valor_b):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'TipoDocumento': row['TipoDocumento'],
                        'Documento': row['Documento'],
                        'PrimerNombre': row['PrimerNombre'],
                        'SegundoNombre': row['SegundoNombre'],
                        'PrimerApellido': row['PrimerApellido'],
                        'SegundoApellido': row['SegundoApellido'],
                        'Escritura': row['Escritura'],
                        'FechaEscritura': row['FechaEscritura'],
                        'Entidad': row['Entidad'],
                        'Observacion': 'Documento diferente a blanco para código asignado',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    # Agregar solo el resultado actual
                    self.agregar_resultados([resultado])
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)

            # Guardar el resultado en un nuevo archivo Excel
            '''
            output_file = 'DOCUMENTO_CODIGO_ASIGNADO.xlsx'
            sheet_name = 'DOCUMENTO_CODIGO_ASIGNADO'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            

            
            messagebox.showinfo("Éxito", f"Proceso completado Codigo Asignado. con {len(resultados)} registros.")
            '''
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def fecha_escritura_inferior(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

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

            # Umbral de fecha (1 de enero de 1778)
            fecha_umbral = datetime(1778, 1, 1)

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                fecha_escritura = row['FechaEscritura']

                # Verificar si 'FechaEscritura' es un string y convertirlo si es necesario
                if isinstance(fecha_escritura, str):
                    try:
                        fecha_obj = datetime.strptime(fecha_escritura, "%d/%m/%Y")
                    except ValueError:
                        print(f"Error en el formato de fecha en la fila {index}: '{fecha_escritura}'")
                        continue  # Saltar a la siguiente fila si la conversión falla
                elif isinstance(fecha_escritura, datetime):
                    fecha_obj = fecha_escritura
                else:
                    print(f"Tipo no válido en la fila {index}: '{fecha_escritura}'")
                    continue  # Saltar a la siguiente fila si el tipo es inesperado

                # Verificar si la fecha es anterior a 1778
                if fecha_obj < fecha_umbral:
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'FechaEscritura': fecha_obj.strftime("%d/%m/%Y"),
                        'Observacion': 'Fecha anterior a 1778',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    self.agregar_resultados(resultados)
                    print(f"Fila {index}: Fecha '{fecha_obj}' es anterior a 1778. Agregado a resultados.")

            print(f"Total de fechas anteriores a 1778 encontradas: {len(resultados)}")

            if len(resultados) > 0:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)

                # Guardar el resultado en un nuevo archivo Excel
                '''
                output_file = 'FECHAS_ESCRITURA_INFERIORES_1778.xlsx'
                sheet_name = 'fechas_inferiores_1778'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")
                

                messagebox.showinfo("Éxito",
                                    f"Proceso completado. Fechas inferiores a 1778: {len(resultados)} registros.")
                '''
            else:
                print("No se encontraron fechas anteriores a 1778.")
                messagebox.showinfo("Información", "No se encontraron registros con fechas anteriores a 1778.")

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    
    
    def fecha_escritura_mayor(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

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

            # Obtener la fecha actual (sin tiempo para evitar diferencias por horas/minutos)
            fecha_actual = datetime.now().date()

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                fecha_escritura = row['FechaEscritura']

                # Verificar si 'FechaEscritura' es nula o NaT
                if pd.isnull(fecha_escritura):
                    print(f"Fila {index}: Fecha de escritura es nula. Ignorada.")
                    continue

                # Convertir a fecha si es necesario
                if isinstance(fecha_escritura, str):
                    try:
                        fecha_escritura = datetime.strptime(fecha_escritura, "%d/%m/%Y").date()
                    except ValueError:
                        print(f"Error en el formato de fecha en la fila {index}: '{fecha_escritura}'")
                        continue

                elif isinstance(fecha_escritura, datetime):
                    fecha_escritura = fecha_escritura.date()  # Tomar solo la parte de la fecha

                # Verificar si la fecha es superior a la fecha actual
                if fecha_escritura > fecha_actual:
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'FechaEscritura': fecha_escritura.strftime("%d/%m/%Y"),
                        'Observacion': 'Fecha de escritura es superior a la fecha actual',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    # Solo se agrega el resultado actual, no toda la lista
                    self.agregar_resultados([resultado])
                    print(f"Fila {index}: Fecha '{fecha_escritura}' es superior a la fecha actual. Agregado a resultados.")

            print(f"Total de fechas superiores a la fecha actual encontradas: {len(resultados)}")

            if len(resultados) > 0:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                '''
                # Guardar el resultado en un nuevo archivo Excel
                output_file = 'FECHAS_ESCRITURA_SUPERIORES.xlsx'
                sheet_name = 'fechas_superiores'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

                messagebox.showinfo("Éxito", f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
                '''
            else:
                print("No se encontraron fechas superiores a la fecha actual.")
                messagebox.showinfo("Información", "No se encontraron registros con fechas superiores a la fecha actual.")

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
               
               
    
    def entidadvacio(self):
        
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            resultados = []

            

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                
               Entidad=row['Entidad'] 
               EntidadDepartamento = row['EntidadDepartamento']
               EntidadMunicipio= row['EntidadMunicipio']
                        
               if pd.isna(EntidadDepartamento) or EntidadDepartamento=='' or EntidadMunicipio=='' or pd.isna(EntidadMunicipio) or Entidad=='':
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'EntidadDepartamento':row['EntidadDepartamento'],
                        'EntidadMunicipio':row['EntidadMunicipio'],
                        'Observacion': 'Entidad no puede ser vacia o null',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    self.agregar_resultados([resultado])
                    print(f"Fila {index}: Agregado a resultados: {resultado}")
                    
            print(f"Entidades vacias: {len(resultados)}")
            '''
            if len(resultados) > 0:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                
                # Guardar el resultado en un nuevo archivo Excel
                output_file = 'EntidadesVacias.xlsx'
                sheet_name = 'fechas_superiores'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

                messagebox.showinfo("Éxito", f"Proceso completado. Entidades vacias '{output_file}' con {len(resultados)} registros.")
               
            else:
                print("No se encontraron Entidades vacias.")
                messagebox.showinfo("Información", "No se encontraron registros con fechas superiores a la fecha actual.")
             '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def numerofallocero(self):
        
        archivo_excel= self.archivo_entry.get()
        nombre_hoja = 'Propietarios'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            resultados = []

            

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
               NumeroFallo = row['NumeroFallo']
               
                        
               if   NumeroFallo== '0' or NumeroFallo=='' or pd.isna(NumeroFallo):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'EntidadDepartamento':row['EntidadDepartamento'],
                        'EntidadMunicipio':row['EntidadMunicipio'],
                        'NumeroFallo':row['NumeroFallo'],
                        'Observacion': 'El numero fallo es cero o vacio',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    self.agregar_resultados([resultado])
                    print(f"Fila {index}: Agregado a resultados: {resultado}")
                    
            print(f"Entidades vacias: {len(resultados)}")
            '''
            if len(resultados) > 0:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                
                # Guardar el resultado en un nuevo archivo Excel
                output_file = 'NumeroFallo.xlsx'
                sheet_name = 'Numero fallo'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

                messagebox.showinfo("Éxito", f"Proceso completado. Numerofallo '{output_file}' con {len(resultados)} registros.")
                
            else:
                print("No se encontraron Numerofallo.")
                messagebox.showinfo("Información", "No se encontraron registros con fechas superiores a la fecha actual.")
            '''
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
        
        
    
    