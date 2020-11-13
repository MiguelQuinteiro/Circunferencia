VERSION 5.00
Begin VB.Form frmCircunferencia 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Primos sobre Circunferencia"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   12885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Controles "
      Height          =   9135
      Left            =   9720
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.CheckBox chkRegular 
         Caption         =   "Ver"
         Height          =   495
         Left            =   1560
         TabIndex        =   6
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.ListBox lstResaltados 
         Height          =   6495
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   2535
      End
      Begin VB.CheckBox chkPrimos 
         Caption         =   "Ver"
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtResalta 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   1
         Text            =   "100"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Regular"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Mostrar Primos"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Resalta cada tantos Primos"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
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
Dim miRelacionInfSupe As Double
Dim miResalte As Long
Dim miResalteRegular As Long

Dim miCuentaAntesPhi As Long
Dim miCuentaDespuesPhi As Long
Dim miPhi As Double
Dim miGrande As Long
Dim miPequeño As Long


' AL HACER DOBLE CLICK
Private Sub Form_DblClick()
' Inicialización de variable
  miPi = 3.1415926535
  miRadio = 4000
  miFactorCircular = 1.15
  miN = 5000
  miColor = 12
  miCuentaSuperior = 0
  miCuentaInferior = 0
  miCuentaPrimos = 0
  miResalte = 0
  miResalteRegular = 0

  ' miPhi = 1.23
  miPhi = 1.61803398874989
  miCuentaAntesPhi = 0
  miCuentaDespuesPhi = 0

  ' Dibuja Circulo
  miN = InputBox("Ingrese el número N (Entre 1 y 30000)")
  If miN <= 3000000 Then
    Cls

    ' Consigue la sección Aurea
    miPequeño = miN / (1 + miPhi)
    miGrande = miN - miPequeño
    'Print miGrande
    'Print miPequeño

    ' Marco
    'Line (100, 100)-(9500, 9500), , B

    ' Ejes de Coordenadas
    Line (4750, 0)-(4750, 9500)
    Line (0, 4750)-(9500, 4750)
    Line (0, 0)-(9500, 9500)
    Line (0, 9500)-(9500, 0)

    ' Borra el área de la circunferencia
    Dim r As Long
    For r = 1 To miRadio * miFactorCircular
      Circle (4750, 4750), r, frmCircunferencia.BackColor
      'Circle (4750, 4750), r, vbRed
    Next r

    ' Pinta los primos
    lstResaltados.Clear
    Dim i As Long
    For i = 1 To miN

      ' Dibuja intervalos regulares
      miResalteRegular = miResalteRegular + 1
      If chkRegular.Value = 1 Then
        If miResalteRegular = 2 * Val(txtResalta.Text) Then
          miResalteRegular = 0
          Circle (X1, Y1), 5, vbYellow
          Line (X1, Y1)-(4750, 4750), vbYellow

        End If
      End If


      X1 = 4750 + (miRadio * Cos((360 / miN) * (miPi / 180) * i) * miFactorCircular)
      Y1 = 4750 + (miRadio * -Sin((360 / miN) * (miPi / 180) * i) * miFactorCircular)
      ' Calcula si es primo
      If Primo(i) = True Then
        miCuentaPrimos = miCuentaPrimos + 1
        miResalte = miResalte + 1
        ' Calcula cantidad Superior e inferior
        If Y1 <= 4750 Then
          miCuentaSuperior = miCuentaSuperior + 1
        Else
          miCuentaInferior = miCuentaInferior + 1
        End If
        ' Dibuja al primo
        If chkPrimos.Value = 1 Then
          Circle (X1, Y1), 5, QBColor(miColor)
          Line (X1, Y1)-(4750, 4750), QBColor(miColor)
        End If

        '                ' Dibuja el resaltado
        '                If miResalte = Val(txtResalta.Text) Then
        '                    miResalte = 0
        '                    Circle (X1, Y1), 5, vbBlack
        '                    Line (X1, Y1)-(4750, 4750), vbBlack
        '                    lstResaltados.AddItem i
        '                End If
        '
        '                ' Relación con la seccion aurea
        If i < miPequeño Then
          miCuentaAntesPhi = miCuentaAntesPhi + 1
          Line (X1, Y1)-(4750, 4750), vbBlack
        Else
          miCuentaDespuesPhi = miCuentaDespuesPhi + 1
        End If

      Else
        'Circle (X1, Y1), 10, QBColor(0)
      End If
    Next i
    Dim miTotalPrimos As Long

    '        ' Marca la mitad de los primos
    '        miTotalPrimos = miCuentaPrimos
    '        miCuentaPrimos = 0
    '        For i = 1 To miN
    '            If Primo(i) Then
    '                miCuentaPrimos = miCuentaPrimos + 1
    '                If miCuentaPrimos > (miTotalPrimos / 2) Then
    '                    X1 = 4750 + (miRadio * Cos((360 / miN) * (miPi / 180) * i) * miFactorCircular)
    '                    Y1 = 4750 + (miRadio * -Sin((360 / miN) * (miPi / 180) * i) * miFactorCircular)
    '                    Circle (X1, Y1), 5, vbWhite
    '                    Line (X1, Y1)-(4750, 4750), vbWhite
    '                End If
    '            End If
    '        Next i

    miRelacionInfSupe = miCuentaInferior / miCuentaSuperior
    MsgBox "C. Puntos =" + Str(miN) + vbCrLf + vbCrLf _
           + "C. Primos =" + Str(miCuentaPrimos) + vbCrLf + vbCrLf _
           + "Superior  =" + Str(miCuentaSuperior) + vbCrLf + vbCrLf _
           + "Inferior   =" + Str(miCuentaInferior) + vbCrLf + vbCrLf _
           + "Relación  =" + Str(miRelacionInfSupe) + vbCrLf + vbCrLf _
           + "Sobre Phi =" + Str(miCuentaAntesPhi) + vbCrLf + vbCrLf _
           + "Bajo Phi  =" + Str(miCuentaDespuesPhi), , "INFORMACIÓN"
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

