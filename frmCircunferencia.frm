VERSION 5.00
Begin VB.Form frmCircunferencia 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Primos sobre Circunferencia"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCuenta1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtCuenta2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   8160
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCircunferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim X1 As Double
Dim Y1 As Double
Dim miPi As Double
Dim miRadio As Long
Dim miFactorCircular As Double
Dim miN As Long
Dim miColor As Long
Dim miCuentaPrimos As Long
Dim miCuentaSuperior As Long
Dim miCuentaInferior As Long
'Dim miCuentaAntesPhi As Long
'Dim miCuentaDespuesPhi As Long
Dim miRelacionInfSupe As Double
'Dim miPhi As Double
'Dim miGrande As Long
'Dim miPequeño As Long


' AL HACER DOBLE CLICK
Private Sub Form_DblClick()
' Inicialización de variable
  miPi = 3.1415926535
  '   miPhi = 1.61803398874989
  miRadio = 4000
  miFactorCircular = 1.15
  miN = 5000
  miColor = 12
  miCuentaSuperior = 0
  miCuentaInferior = 0
  '    miCuentaAntesPhi = 0
  '    miCuentaDespuesPhi = 0

  ' Dibuja Circulo
  miN = InputBox("Ingrese el número N (Entre 1 y 30000)")
  If miN <= 10000000 Then
    Cls

    '        ' Consigue la sección Aurea
    '        miPequeño = miN / (1 + miPhi)
    '        miGrande = miN - miPequeño
    '        Print miGrande
    '        Print miPequeño

    ' Marco
    'Line (100, 100)-(9500, 9500), , B

    ' Ejes de Coordenadas
    Line (4750, 0)-(4750, 9500), vbWhite
    Line (0, 4750)-(9500, 4750), vbWhite
    Line (0, 0)-(9500, 9500), vbWhite
    Line (0, 9500)-(9500, 0), vbWhite

    ' Borra el área de la circunferencia
    Dim r As Long
    For r = 1 To miRadio * miFactorCircular
      Circle (4750, 4750), r, frmCircunferencia.BackColor
      'Circle (4750, 4750), r, vbRed
    Next r

    ' Recorre la circunferencia
    Dim i As Long
    For i = 1 To miN
      X1 = 4750 + (miRadio * Cos((360 / miN) * (miPi / 180) * i) * miFactorCircular)
      Y1 = 4750 + (miRadio * -Sin((360 / miN) * (miPi / 180) * i) * miFactorCircular)

      ' Dibuja Números
      If i / 2 = Int(i / 2) Then
        Circle (X1, Y1), 5, vbRed
        Line (X1, Y1)-(4750, 4750), vbRed
        miCuentaSuperior = miCuentaSuperior + 1
      Else
        Circle (X1, Y1), 5, vbWhite
        Line (X1, Y1)-(4750, 4750), vbWhite
        miCuentaInferior = miCuentaInferior + 1
      End If


      '            ' Relación con la seccion aurea
      '            If i < miPequeño Then
      '                miCuentaAntesPhi = miCuentaAntesPhi + 1
      '                Line (X1, Y1)-(4750, 4750), vbYellow
      '            Else
      '                miCuentaDespuesPhi = miCuentaDespuesPhi + 1
      '            End If


    Next i
    txtCuenta1.Text = miCuentaSuperior
    txtCuenta2.Text = miCuentaInferior
  End If
End Sub

' Funcion para saber si un numero es primo
Public Function Primo(ByVal pN As Long) As Boolean
  Dim i As Long
  Primo = True
  If pN = 1 Then
    Primo = False
  Else
    For i = 2 To Sqr(pN)
      If (pN / i) = Int(pN / i) Then
        Primo = False
      End If
    Next i
  End If
End Function

' Funcion para calcular claves
Public Function CalculaClave(ByVal pN As Long) As Boolean
  Dim i As Long
  Dim miCantidadFactores As Long
  Dim miElemento(4) As Long


  CalculaClave = False
  miCantidadFactores = 0
  For i = 1 To pN
    If pN / i = Int(pN / i) Then
      miCantidadFactores = miCantidadFactores + 1
      If miCantidadFactores <= 4 Then
        miElemento(miCantidadFactores) = i
      Else
        i = pN
      End If
    End If
  Next i

  If miCantidadFactores = 4 Then
    'Print Tab(20), pN, miElemento(2), miElemento(3)
    CalculaClave = True
  End If

End Function

