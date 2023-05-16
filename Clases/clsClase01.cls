VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClase01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : DESARROLLO GENERAL
'* CONTENIDO     : CLASE PARA DEFINIR TIPO DE OBJETO
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO / MIGUEL QUINTEIRO FERNANDEZ
'* INICIO        : 99 DE XXXXXXX DE 9999
'* ACTUALIZACION : 99 DE XXXXXXX DE 9999
'****************************************************************************************
Option Explicit
Option Base 1

' AREA DE DATOS
'*******************************************************************
' Declaracion de los Atributos privados para encapsulamiento
Private aAtributo01 As Double
    
' Metodos Get y Let para manipulacion de datos de la clase
' Get Atributo01
Public Property Get Atributo01() As Double
    Atributo01 = aAtributo01
End Property
' Let Atributo01
Public Property Let Atributo01(ByVal valor As Double)
    aAtributo01 = valor
End Property

' AREA DE METODOS Y FUNCIONALIDAD
'*******************************************************************
' Metodo constructor y destructor de la clase
' Constructor
Private Sub Class_Initialize()
    ' Que hacer al instanciar el objeto
    aAtributo01 = 0
End Sub
' Destructor
Sub Class_Terminate()
    ' Que hacer al destruir el objeto
End Sub

' Metodos publicos para la funcionalidad de la clase
' Imprimir el Objeto
Public Sub Imprimir()
    ' Instrucciones
    frmPrincipal.Cls
    frmPrincipal.Print Atributo01
End Sub

