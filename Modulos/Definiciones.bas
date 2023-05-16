Attribute VB_Name = "Definiciones"
'****************************************************************************************
'* PROYECTO      : DESARROLLO GENERAL
'* CONTENIDO     : DEFINICIONES GLOBALES
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO / MIGUEL QUINTEIRO FERNANDEZ
'* INICIO        : 99 DE XXXXXXX DE 9999
'* ACTUALIZACION : 99 DE XXXXXXX DE 9999
'****************************************************************************************
Option Explicit
Option Base 1

' Declaración de indices genéricos
Public i As Double
Public j As Double
Public k As Double

' Declaración de variables de trabajo
Public auxiliar As Double                   ' Variable auxiliar
Public vector(16) As Double                 ' Vector unidimensional
Public matriz(1 To 16, 1 To 16) As Double   ' Matriz dos dimensiones

' Declaración de los objetos
Public oObjeto01 As clsClase01 ' Objeto de la clase 01

' Declaración de vector de objetos
Public vecObjeto(16) As clsClase01  ' Contiene vector de objetos de la clase 01


