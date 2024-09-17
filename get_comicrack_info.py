# coding: utf-8
#get_comicrack_info.py
#Author: Tomas Platero
#Date: 08/08/2023
#Description: Get the comicrack api info
#Version: 1.0
#
#
#@Name Get CR API Info
#@Hook Books
#@Key Get CR API Info
#@Enabled true
#@Get all the ComicRack API Info
#@Image get_comicrack_info.png
import clr
import sys
import os
import System
import csv
import re
from System import Reflection
import datetime

debug = True

# Mapping common .NET types to more readable names
tipo_mapeo = {
    "System.String": "str",
    "System.Int32": "int32",
    "System.Boolean": "bool",
    "System.Object": "object",
    "System.Void": "void",
    "System.Double": "double",
    "System.Float": "float",
    "System.Decimal": "decimal",
    "System.Char": "char",
    "System.Byte": "byte",
    "System.Int64": "int64",
    "System.UInt32": "uint32",
    "System.UInt64": "uint64",
    "System.Int16": "int16",
    "System.UInt16": "uint16",
    # Add more if necesary
}

# .NET methods to filter
metodos_a_filtrar = ["Equals", 
                     "GetHashCode", 
                     "GetType", 
                     "ToString", 
                     "Clone", 
                     "Set", 
                     "Match", 
                     "IsSame", 
                     "CompareTo"
                     # Add more if necesary
                     ]

# Classes to be filtered
clases_a_filtrar = ["mscorlib", 
                    "System",
                    "cYo.Projects.ComicRack.Engine.CacheManager"]


def get_comicrack_info(True):

    # Replace with the path of your DLL
    dll_path = "C:\\Program Files\\ComicRack Community Edition\\cYo.Common.dll"

    # Load DLL using the full path
    clr.AddReferenceToFileAndPath(dll_path)

    # Get the assembly loaded
    assembly = clr.LoadAssemblyFromFileWithPath(dll_path)

    # Create a list to store the data
    datos = []

    # Get all types within the assembly
    tipos = assembly.GetTypes()

    # Iterating over types to extract namespaces and their methods and properties
    for tipo in tipos:
        namespace = tipo.Namespace
        if debug: print("Namespace: " + str(namespace))

        # Filter out non-.NET namespaces
        if namespace and not (namespace.startswith("Microsoft") or namespace.startswith("System")):
            
            # Get only public
            if tipo.IsPublic:
                
                # Obtain the classname name
                tipo_class = tipo.FullName
                if debug: print("Class: " + str(tipo_class))

                # Get and list all public and invocable methods
                metodos = tipo.GetMethods(Reflection.BindingFlags.Public | Reflection.BindingFlags.Instance | Reflection.BindingFlags.Static)
                for metodo in metodos:
                    if metodo.IsPublic and not metodo.IsSpecialName and filter_method(metodo):
                        parametros = metodo.GetParameters()
                        parametros_info = ', '.join("{} ({})".format(param.Name, get_legible_type(param.ParameterType)) for param in parametros)
                        parametros_info_clean = clean_strings(parametros_info)
                        tipo_retorno = get_legible_type(metodo.ReturnType)
                        nombre_metodo = metodo.Name
                        

                        # Add method data to the array
                        datos.append([namespace, tipo_class, " Method ", nombre_metodo, parametros_info_clean , tipo_retorno])

                        if debug:
                            print("\tMethod Name: " + metodo.Name)
                            print("\t\tCall Parameters: " + parametros_info_clean)
                            print("\t\tReturn Type: " + str(tipo_retorno))
                            print("")
                
                # Get and list all public and invocable properties
                propiedades = tipo.GetProperties(Reflection.BindingFlags.Public | Reflection.BindingFlags.Instance | Reflection.BindingFlags.Static)
                for propiedad in propiedades:
                    if (propiedad.CanRead and propiedad.GetMethod.IsPublic) or (propiedad.CanWrite and propiedad.SetMethod.IsPublic):
                        tipo_propiedad = get_legible_type(propiedad.PropertyType)
                        tipo_propiedad_str = str(tipo_propiedad)
                        tipo_propiedad_str = clean_strings(tipo_propiedad_str)
                        if not tipo_propiedad_str.startswith("System"):
                            # Determine whether the property is get only or get/set
                            acceso = "get" if propiedad.CanRead else ""
                            if propiedad.CanWrite:
                                acceso += "/set" if acceso else "set"

                            # There are no parameters in the properties, but we unify the property with its type
                            firma = "{} ({})".format(tipo_propiedad,acceso)

                            if debug:
                                print("\tPropierty Name: " + propiedad.Name)
                                print("\t\tReturn Type: " + tipo_propiedad_str)
                                print("\t\tAccess: " + acceso)
                                print("")
                            
                            # Add property data to the array
                            datos.append([namespace, tipo_class, " Propierty ", propiedad.Name, "" , tipo_propiedad_str , acceso ])
                            
            # Add a blank line 
            datos.append(["","","","","","",""]) # Añadir 
            if debug: print("")
                            
    # Get the base name of the DLL file (without the path and without the extension)
    nombre_archivo_base = os.path.splitext(os.path.basename(dll_path))[0]

    # Get the directory where the script is located
    script_directory = os.path.dirname(os.path.abspath(__file__))

    # Create a date and time stamp
    fecha_hora = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    # Create the full name of the CSV file with the date and time stamp in the same directory as the script
    csv_filename = os.path.join(script_directory, nombre_archivo_base + "_" + fecha_hora + ".csv")

    # Write the data to a CSV file
    print ("Creating the file: " + csv_filename)
    with open(csv_filename, mode='wb') as file:
        writer = csv.writer(file, delimiter=';')  # Specify the ";" delimiter
        # write header
        writer.writerow(["Namespace", "Class", "Type", "Name", "Parameters", "Return" , "Access"])

        # Write the data
        writer.writerows(datos)
    print ("The CSV file '" + csv_filename + "' has been successfully created.")
    return

def filter_method(metodo):
    # Check if the method should be filtered
    if metodo.Name in metodos_a_filtrar:
       return False  

    # Check if the method belongs to a standard .NET assembly
    declaring_assembly = metodo.DeclaringType.Assembly.GetName().Name
    for clase in clases_a_filtrar:
        if declaring_assembly == clase or declaring_assembly.startswith(clase + "."):
            return False  
    
    return True  

# Function to get the mapped type name
def get_legible_type(tipo):
    tipo_str = str(tipo)
    tipo_mapeo_final = tipo_mapeo.get(tipo_str, tipo_str)
    #.split('.')[-1])
    return  tipo_mapeo_final  # Returns the mapped type or the last segment of the full name


# Function to make call parameters more readable
def clean_strings(parameter):
    v_parameter = re.sub(r'\(([^)]+)\)', lambda x: '(' + x.group(1).split('.')[-1] + ')', parameter)
    return v_parameter


# Delete unwanted characters
def strip_invalid_chars(text):
   def is_valid(c):
      return c == 0x9 or c == 0xA or c == 0xD or\
         (c >= 0x20 and c <= 0xD7FF) or\
         (c >= 0xE000 and c <= 0xFFFD) or\
         (c >= 0x10000 and c <= 0x10FFFF)
   if text:
      text = ''.join([c for c in text if is_valid(ord(c))])
   return text
