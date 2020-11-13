VERSION 5.00
Begin VB.Form frmCircunferencia 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   Caption         =   "Primos sobre Circunferencia"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   12885
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCeros 
      Height          =   450
      Left            =   8280
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   " Controles "
      Height          =   9135
      Left            =   9720
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.TextBox txtParticion 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   4440
         Width           =   2535
      End
      Begin VB.TextBox txtLn 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   10
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtRaiz 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   1560
         TabIndex        =   8
         Top             =   2400
         Width           =   1215
      End
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
         Height          =   2985
         Left            =   240
         TabIndex        =   5
         Top             =   5160
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
      Begin VB.Label Label6 
         Caption         =   "Partición Fómula"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Logaritmo Neperiano de N"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Raiz Cuadrada de N"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   2400
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
Dim miRaiz As Double
Dim miParticion As Double
Dim miLn As Double

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
  miParticion = 0
  miLn = 0

  ' Dibuja Circulo
  miN = InputBox("Ingrese el número N (Entre 1 y 30000)")
  If miN <= 300000 Then
    Cls

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


    ' Calcula ra raiz del número
    miRaiz = Sqr(miN)
    miLn = Log(miN)
    If miN < 8100 Then
      miParticion = (1 / (4 * miN * Sqr(3))) * Exp(miPi * Sqr((2 * miN / 3)))
    End If
    txtRaiz.Text = miRaiz
    txtLn.Text = miLn
    txtParticion.Text = miParticion

    ' Pinta los primos
    lstResaltados.Clear
    Dim i As Long
    For i = 1 To miN
      X1 = 4750 + (miRadio * Cos((360 / miN) * (miPi / 180) * i) * miFactorCircular)
      Y1 = 4750 + (miRadio * -Sin((360 / miN) * (miPi / 180) * i) * miFactorCircular)

      ' Dibuja intervalos regulares
      miResalteRegular = miResalteRegular + 1
      If chkRegular.Value = 1 Then
        If miResalteRegular = Val(txtResalta.Text) Then
          miResalteRegular = 0
          Circle (X1, Y1), 5, vbYellow
          'Circle (X1, Y1), i, vbRed
          Line (X1, Y1)-(4750, 4750), vbYellow

        End If
      End If


      'CalculaClave
      'If CalculaClave(i) = True Then

      ' Calcula si es primo
      'If Primo(i) = True Then
      If Riemann(i) = True Then
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

          ' PRIMOS GEMELOS
          'If Primo(i - 2) Then
          Circle (X1, Y1), 10, QBColor(miColor)
          'PSet (X1, Y1)
          'Print i
          Line (X1, Y1)-(4750, 4750), vbYellow
          'Else
          '    Circle (X1, Y1), 5, QBColor(miColor)
          'End If


        End If
        ' Dibuja el resaltado
        If miResalte = Val(txtResalta.Text) Then
          miResalte = 0
          Circle (X1, Y1), 5, vbBlack
          Line (X1, Y1)-(4750, 4750), vbBlack
          lstResaltados.AddItem i

        End If
      Else
        'Circle (X1, Y1), 10, QBColor(0)
      End If

      If Int(miRaiz) = i Then
        Circle (X1, Y1), 100, vbWhite
        Line (X1, Y1)-(4750, 4750), vbWhite
      End If
      If Int(miLn) = i Then
        Circle (X1, Y1), 100, vbWhite
        Line (X1, Y1)-(4750, 4750), vbWhite
      End If
      If Int(miParticion) = i Then
        Circle (X1, Y1), 100, vbWhite
        Line (X1, Y1)-(4750, 4750), vbWhite
      End If

    Next i

    miRelacionInfSupe = miCuentaInferior / miCuentaSuperior
    MsgBox "C. Puntos =" + Str(miN) + vbCrLf + vbCrLf _
           + "C. Primos =" + Str(miCuentaPrimos) + vbCrLf + vbCrLf _
           + "Superior  =" + Str(miCuentaSuperior) + vbCrLf + vbCrLf _
           + "Inferior   =" + Str(miCuentaInferior) + vbCrLf + vbCrLf _
           + "Relación  =" + Str(miRelacionInfSupe), , "INFORMACIÓN"
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


Public Function Riemann(ByVal pN As Long) As Boolean
  Dim i As Integer
  Riemann = False
  For i = 0 To lstCeros.ListCount - 1
    lstCeros.ListIndex = i
    If pN = Val(lstCeros.Text) Then
      Riemann = True
      i = lstCeros.ListCount
    End If
  Next i
End Function



Private Sub Form_Load()
  lstCeros.AddItem 14
  lstCeros.AddItem 21
  lstCeros.AddItem 25
  lstCeros.AddItem 30
  lstCeros.AddItem 33
  lstCeros.AddItem 38
  lstCeros.AddItem 41
  lstCeros.AddItem 43
  lstCeros.AddItem 48
  lstCeros.AddItem 50
  lstCeros.AddItem 53
  lstCeros.AddItem 56
  lstCeros.AddItem 59
  lstCeros.AddItem 61
  lstCeros.AddItem 65
  lstCeros.AddItem 67
  lstCeros.AddItem 70
  lstCeros.AddItem 72
  lstCeros.AddItem 76
  lstCeros.AddItem 77
  lstCeros.AddItem 79
  lstCeros.AddItem 83
  lstCeros.AddItem 85
  lstCeros.AddItem 87
  lstCeros.AddItem 89
  lstCeros.AddItem 92
  lstCeros.AddItem 95
  lstCeros.AddItem 96
  lstCeros.AddItem 99
  lstCeros.AddItem 101
  lstCeros.AddItem 104
  lstCeros.AddItem 105
  lstCeros.AddItem 107
  lstCeros.AddItem 111
  lstCeros.AddItem 112
  lstCeros.AddItem 114
  lstCeros.AddItem 116
  lstCeros.AddItem 119
  lstCeros.AddItem 121
  lstCeros.AddItem 123
  lstCeros.AddItem 124
  lstCeros.AddItem 128
  lstCeros.AddItem 130
  lstCeros.AddItem 131
  lstCeros.AddItem 133
  lstCeros.AddItem 135
  lstCeros.AddItem 138
  lstCeros.AddItem 140
  lstCeros.AddItem 141
  lstCeros.AddItem 143
  lstCeros.AddItem 146
  lstCeros.AddItem 147
  lstCeros.AddItem 150
  lstCeros.AddItem 151
  lstCeros.AddItem 153
  lstCeros.AddItem 156
  lstCeros.AddItem 158
  lstCeros.AddItem 159
  lstCeros.AddItem 161
  lstCeros.AddItem 163
  lstCeros.AddItem 166
  lstCeros.AddItem 167
  lstCeros.AddItem 169
  lstCeros.AddItem 170
  lstCeros.AddItem 173
  lstCeros.AddItem 175
  lstCeros.AddItem 176
  lstCeros.AddItem 178
  lstCeros.AddItem 180
  lstCeros.AddItem 182
  lstCeros.AddItem 185
  lstCeros.AddItem 186
  lstCeros.AddItem 187
  lstCeros.AddItem 189
  lstCeros.AddItem 192
  lstCeros.AddItem 193
  lstCeros.AddItem 195
  lstCeros.AddItem 197
  lstCeros.AddItem 198
  lstCeros.AddItem 201
  lstCeros.AddItem 202
  lstCeros.AddItem 204
  lstCeros.AddItem 205
  lstCeros.AddItem 208
  lstCeros.AddItem 210
  lstCeros.AddItem 212
  lstCeros.AddItem 213
  lstCeros.AddItem 215
  lstCeros.AddItem 216
  lstCeros.AddItem 219
  lstCeros.AddItem 221
  lstCeros.AddItem 221
  lstCeros.AddItem 224
  lstCeros.AddItem 225
  lstCeros.AddItem 227
  lstCeros.AddItem 229
  lstCeros.AddItem 231
  lstCeros.AddItem 232
  lstCeros.AddItem 234
  lstCeros.AddItem 237
  lstCeros.AddItem 238
  lstCeros.AddItem 240
  lstCeros.AddItem 241
  lstCeros.AddItem 243
  lstCeros.AddItem 244
  lstCeros.AddItem 247
  lstCeros.AddItem 248
  lstCeros.AddItem 250
  lstCeros.AddItem 251
  lstCeros.AddItem 253
  lstCeros.AddItem 255
  lstCeros.AddItem 256
  lstCeros.AddItem 259
  lstCeros.AddItem 260
  lstCeros.AddItem 261
  lstCeros.AddItem 264
  lstCeros.AddItem 266
  lstCeros.AddItem 267
  lstCeros.AddItem 268
  lstCeros.AddItem 270
  lstCeros.AddItem 271
  lstCeros.AddItem 273
  lstCeros.AddItem 276
  lstCeros.AddItem 276
  lstCeros.AddItem 278
  lstCeros.AddItem 279
  lstCeros.AddItem 282
  lstCeros.AddItem 283
  lstCeros.AddItem 285
  lstCeros.AddItem 287
  lstCeros.AddItem 288
  lstCeros.AddItem 290
  lstCeros.AddItem 292
  lstCeros.AddItem 294
  lstCeros.AddItem 295
  lstCeros.AddItem 296
  lstCeros.AddItem 298
  lstCeros.AddItem 300
  lstCeros.AddItem 302
  lstCeros.AddItem 303
  lstCeros.AddItem 305
  lstCeros.AddItem 306
  lstCeros.AddItem 307
  lstCeros.AddItem 310
  lstCeros.AddItem 311
  lstCeros.AddItem 312
  lstCeros.AddItem 314
  lstCeros.AddItem 315
  lstCeros.AddItem 318
  lstCeros.AddItem 319
  lstCeros.AddItem 321
  lstCeros.AddItem 322
  lstCeros.AddItem 323
  lstCeros.AddItem 325
  lstCeros.AddItem 327
  lstCeros.AddItem 329
  lstCeros.AddItem 330
  lstCeros.AddItem 331
  lstCeros.AddItem 334
  lstCeros.AddItem 334
  lstCeros.AddItem 337
  lstCeros.AddItem 338
  lstCeros.AddItem 340
  lstCeros.AddItem 341
  lstCeros.AddItem 342
  lstCeros.AddItem 345
  lstCeros.AddItem 346
  lstCeros.AddItem 347
  lstCeros.AddItem 349
  lstCeros.AddItem 350
  lstCeros.AddItem 352
  lstCeros.AddItem 353
  lstCeros.AddItem 356
  lstCeros.AddItem 357
  lstCeros.AddItem 358
  lstCeros.AddItem 360
  lstCeros.AddItem 361
  lstCeros.AddItem 363
  lstCeros.AddItem 365
  lstCeros.AddItem 366
  lstCeros.AddItem 368
  lstCeros.AddItem 369
  lstCeros.AddItem 370
  lstCeros.AddItem 373
  lstCeros.AddItem 374
  lstCeros.AddItem 376
  lstCeros.AddItem 376
  lstCeros.AddItem 378
  lstCeros.AddItem 380
  lstCeros.AddItem 381
  lstCeros.AddItem 383
  lstCeros.AddItem 385
  lstCeros.AddItem 386
  lstCeros.AddItem 387
  lstCeros.AddItem 389
  lstCeros.AddItem 391
  lstCeros.AddItem 392
  lstCeros.AddItem 393
  lstCeros.AddItem 396
  lstCeros.AddItem 396
  lstCeros.AddItem 398
  lstCeros.AddItem 400
  lstCeros.AddItem 402
  lstCeros.AddItem 403
  lstCeros.AddItem 404
  lstCeros.AddItem 405
  lstCeros.AddItem 408
  lstCeros.AddItem 409
  lstCeros.AddItem 411
  lstCeros.AddItem 412
  lstCeros.AddItem 413
  lstCeros.AddItem 415
  lstCeros.AddItem 415
  lstCeros.AddItem 418
  lstCeros.AddItem 420
  lstCeros.AddItem 421
  lstCeros.AddItem 422
  lstCeros.AddItem 424
  lstCeros.AddItem 425
  lstCeros.AddItem 427
  lstCeros.AddItem 428
  lstCeros.AddItem 430
  lstCeros.AddItem 431
  lstCeros.AddItem 432
  lstCeros.AddItem 434
  lstCeros.AddItem 436
  lstCeros.AddItem 438
  lstCeros.AddItem 439
  lstCeros.AddItem 440
  lstCeros.AddItem 442
  lstCeros.AddItem 443
  lstCeros.AddItem 444
  lstCeros.AddItem 447
  lstCeros.AddItem 447
  lstCeros.AddItem 449
  lstCeros.AddItem 450
  lstCeros.AddItem 451
  lstCeros.AddItem 454
  lstCeros.AddItem 455
  lstCeros.AddItem 456
  lstCeros.AddItem 458
  lstCeros.AddItem 460
  lstCeros.AddItem 460
  lstCeros.AddItem 462
  lstCeros.AddItem 464
  lstCeros.AddItem 466
  lstCeros.AddItem 467
  lstCeros.AddItem 467
  lstCeros.AddItem 470
  lstCeros.AddItem 471
  lstCeros.AddItem 473
  lstCeros.AddItem 474
  lstCeros.AddItem 476
  lstCeros.AddItem 477
  lstCeros.AddItem 478
  lstCeros.AddItem 479
  lstCeros.AddItem 482
  lstCeros.AddItem 483
  lstCeros.AddItem 484
  lstCeros.AddItem 486
  lstCeros.AddItem 487
  lstCeros.AddItem 488
  lstCeros.AddItem 490
  lstCeros.AddItem 491
  lstCeros.AddItem 493
  lstCeros.AddItem 494
  lstCeros.AddItem 495
  lstCeros.AddItem 496
  lstCeros.AddItem 499
  lstCeros.AddItem 500
  lstCeros.AddItem 502
  lstCeros.AddItem 502
  lstCeros.AddItem 504
  lstCeros.AddItem 505
  lstCeros.AddItem 506
  lstCeros.AddItem 509
  lstCeros.AddItem 510
  lstCeros.AddItem 512
  lstCeros.AddItem 513
  lstCeros.AddItem 514
  lstCeros.AddItem 515
  lstCeros.AddItem 518
  lstCeros.AddItem 518
  lstCeros.AddItem 520
  lstCeros.AddItem 522
  lstCeros.AddItem 522
  lstCeros.AddItem 524
  lstCeros.AddItem 525
  lstCeros.AddItem 528
  lstCeros.AddItem 528
  lstCeros.AddItem 530
  lstCeros.AddItem 531
  lstCeros.AddItem 533
  lstCeros.AddItem 534
  lstCeros.AddItem 536
  lstCeros.AddItem 537
  lstCeros.AddItem 538
  lstCeros.AddItem 540
  lstCeros.AddItem 541
  lstCeros.AddItem 542
  lstCeros.AddItem 544
  lstCeros.AddItem 546
  lstCeros.AddItem 547
  lstCeros.AddItem 548
  lstCeros.AddItem 549
  lstCeros.AddItem 551
  lstCeros.AddItem 552
  lstCeros.AddItem 554
  lstCeros.AddItem 556
  lstCeros.AddItem 557
  lstCeros.AddItem 558
  lstCeros.AddItem 559
  lstCeros.AddItem 560
  lstCeros.AddItem 563
  lstCeros.AddItem 564
  lstCeros.AddItem 565
  lstCeros.AddItem 567
  lstCeros.AddItem 568
  lstCeros.AddItem 569
  lstCeros.AddItem 570
  lstCeros.AddItem 572
  lstCeros.AddItem 574
  lstCeros.AddItem 575
  lstCeros.AddItem 576
  lstCeros.AddItem 577
  lstCeros.AddItem 579
  lstCeros.AddItem 580
  lstCeros.AddItem 582
  lstCeros.AddItem 583
  lstCeros.AddItem 585
  lstCeros.AddItem 586
  lstCeros.AddItem 587
  lstCeros.AddItem 588
  lstCeros.AddItem 591
  lstCeros.AddItem 592
  lstCeros.AddItem 593
  lstCeros.AddItem 594
  lstCeros.AddItem 596
  lstCeros.AddItem 596
  lstCeros.AddItem 598
  lstCeros.AddItem 600
  lstCeros.AddItem 602
  lstCeros.AddItem 603
  lstCeros.AddItem 604
  lstCeros.AddItem 605
  lstCeros.AddItem 606
  lstCeros.AddItem 608
  lstCeros.AddItem 609
  lstCeros.AddItem 611
  lstCeros.AddItem 612
  lstCeros.AddItem 614
  lstCeros.AddItem 615
  lstCeros.AddItem 616
  lstCeros.AddItem 618
  lstCeros.AddItem 619
  lstCeros.AddItem 620
  lstCeros.AddItem 622
  lstCeros.AddItem 622
  lstCeros.AddItem 624
  lstCeros.AddItem 626
  lstCeros.AddItem 627
  lstCeros.AddItem 628
  lstCeros.AddItem 630
  lstCeros.AddItem 631
  lstCeros.AddItem 632
  lstCeros.AddItem 634
  lstCeros.AddItem 636
  lstCeros.AddItem 637
  lstCeros.AddItem 638
  lstCeros.AddItem 639
  lstCeros.AddItem 641
  lstCeros.AddItem 642
  lstCeros.AddItem 643
  lstCeros.AddItem 645
  lstCeros.AddItem 646
  lstCeros.AddItem 648
  lstCeros.AddItem 649
  lstCeros.AddItem 650
  lstCeros.AddItem 651
  lstCeros.AddItem 654
  lstCeros.AddItem 654
  lstCeros.AddItem 656
  lstCeros.AddItem 657
  lstCeros.AddItem 658
  lstCeros.AddItem 660
  lstCeros.AddItem 661
  lstCeros.AddItem 662
  lstCeros.AddItem 664
  lstCeros.AddItem 665
  lstCeros.AddItem 667
  lstCeros.AddItem 667
  lstCeros.AddItem 669
  lstCeros.AddItem 670
  lstCeros.AddItem 672
  lstCeros.AddItem 673
  lstCeros.AddItem 674
  lstCeros.AddItem 676
  lstCeros.AddItem 677
  lstCeros.AddItem 678
  lstCeros.AddItem 680
  lstCeros.AddItem 682
  lstCeros.AddItem 683
  lstCeros.AddItem 684
  lstCeros.AddItem 685
  lstCeros.AddItem 686
  lstCeros.AddItem 688
  lstCeros.AddItem 689
  lstCeros.AddItem 690
  lstCeros.AddItem 692
  lstCeros.AddItem 693
  lstCeros.AddItem 695
  lstCeros.AddItem 696
  lstCeros.AddItem 697
  lstCeros.AddItem 699
  lstCeros.AddItem 700
  lstCeros.AddItem 701
  lstCeros.AddItem 702
  lstCeros.AddItem 704
  lstCeros.AddItem 705
  lstCeros.AddItem 706
  lstCeros.AddItem 708
  lstCeros.AddItem 709
  lstCeros.AddItem 711
  lstCeros.AddItem 712
  lstCeros.AddItem 713
  lstCeros.AddItem 714
  lstCeros.AddItem 716
  lstCeros.AddItem 717
  lstCeros.AddItem 719
  lstCeros.AddItem 720
  lstCeros.AddItem 721
  lstCeros.AddItem 722
  lstCeros.AddItem 724
  lstCeros.AddItem 725
  lstCeros.AddItem 727
  lstCeros.AddItem 728
  lstCeros.AddItem 729
  lstCeros.AddItem 730
  lstCeros.AddItem 731
  lstCeros.AddItem 733
  lstCeros.AddItem 735
  lstCeros.AddItem 736
  lstCeros.AddItem 737
  lstCeros.AddItem 739
  lstCeros.AddItem 740
  lstCeros.AddItem 741
  lstCeros.AddItem 742
  lstCeros.AddItem 744
  lstCeros.AddItem 745
  lstCeros.AddItem 746
  lstCeros.AddItem 748
  lstCeros.AddItem 748
  lstCeros.AddItem 751
  lstCeros.AddItem 751
  lstCeros.AddItem 753
  lstCeros.AddItem 754
  lstCeros.AddItem 756
  lstCeros.AddItem 757
  lstCeros.AddItem 758
  lstCeros.AddItem 759
  lstCeros.AddItem 760
  lstCeros.AddItem 763
  lstCeros.AddItem 764
  lstCeros.AddItem 764
  lstCeros.AddItem 766
  lstCeros.AddItem 767
  lstCeros.AddItem 768
  lstCeros.AddItem 770
  lstCeros.AddItem 771
  lstCeros.AddItem 773
  lstCeros.AddItem 774
  lstCeros.AddItem 775
  lstCeros.AddItem 776
  lstCeros.AddItem 777
  lstCeros.AddItem 779
  lstCeros.AddItem 780
  lstCeros.AddItem 782
  lstCeros.AddItem 783
  lstCeros.AddItem 784
  lstCeros.AddItem 786
  lstCeros.AddItem 786
  lstCeros.AddItem 787
  lstCeros.AddItem 790
  lstCeros.AddItem 791
  lstCeros.AddItem 792
  lstCeros.AddItem 793
  lstCeros.AddItem 794
  lstCeros.AddItem 796
  lstCeros.AddItem 797
  lstCeros.AddItem 799
  lstCeros.AddItem 800
  lstCeros.AddItem 802
  lstCeros.AddItem 803
  lstCeros.AddItem 803
  lstCeros.AddItem 805
  lstCeros.AddItem 806
  lstCeros.AddItem 808
  lstCeros.AddItem 809
  lstCeros.AddItem 810
  lstCeros.AddItem 811
  lstCeros.AddItem 813
  lstCeros.AddItem 814
  lstCeros.AddItem 815
  lstCeros.AddItem 817
  lstCeros.AddItem 818
  lstCeros.AddItem 819
  lstCeros.AddItem 821
  lstCeros.AddItem 822
  lstCeros.AddItem 822
  lstCeros.AddItem 825
  lstCeros.AddItem 826
  lstCeros.AddItem 827
  lstCeros.AddItem 828
  lstCeros.AddItem 829
  lstCeros.AddItem 831
  lstCeros.AddItem 832
  lstCeros.AddItem 833
  lstCeros.AddItem 835
  lstCeros.AddItem 837
  lstCeros.AddItem 837
  lstCeros.AddItem 838
  lstCeros.AddItem 839
  lstCeros.AddItem 841
  lstCeros.AddItem 842
  lstCeros.AddItem 844
  lstCeros.AddItem 845
  lstCeros.AddItem 846
  lstCeros.AddItem 848
  lstCeros.AddItem 848
  lstCeros.AddItem 850
  lstCeros.AddItem 851
  lstCeros.AddItem 853
  lstCeros.AddItem 854
  lstCeros.AddItem 855
  lstCeros.AddItem 856
  lstCeros.AddItem 857
  lstCeros.AddItem 859
  lstCeros.AddItem 860
  lstCeros.AddItem 861
  lstCeros.AddItem 863
  lstCeros.AddItem 864
  lstCeros.AddItem 866
  lstCeros.AddItem 866
  lstCeros.AddItem 868
  lstCeros.AddItem 869
  lstCeros.AddItem 871
  lstCeros.AddItem 872
  lstCeros.AddItem 873
  lstCeros.AddItem 874
  lstCeros.AddItem 876
  lstCeros.AddItem 877
  lstCeros.AddItem 878
  lstCeros.AddItem 879
  lstCeros.AddItem 881
  lstCeros.AddItem 882
  lstCeros.AddItem 883
  lstCeros.AddItem 884
  lstCeros.AddItem 885
  lstCeros.AddItem 887
  lstCeros.AddItem 888
  lstCeros.AddItem 890
  lstCeros.AddItem 891
  lstCeros.AddItem 892
  lstCeros.AddItem 893
  lstCeros.AddItem 895
  lstCeros.AddItem 895
  lstCeros.AddItem 897
  lstCeros.AddItem 899
  lstCeros.AddItem 900
  lstCeros.AddItem 901
  lstCeros.AddItem 902
  lstCeros.AddItem 903
  lstCeros.AddItem 905
  lstCeros.AddItem 906
  lstCeros.AddItem 908
  lstCeros.AddItem 908
  lstCeros.AddItem 910
  lstCeros.AddItem 911
  lstCeros.AddItem 912
  lstCeros.AddItem 913
  lstCeros.AddItem 915
  lstCeros.AddItem 916
  lstCeros.AddItem 918
  lstCeros.AddItem 919
  lstCeros.AddItem 919
  lstCeros.AddItem 921
  lstCeros.AddItem 923
  lstCeros.AddItem 923
  lstCeros.AddItem 925
  lstCeros.AddItem 927
  lstCeros.AddItem 928
  lstCeros.AddItem 929
  lstCeros.AddItem 930
  lstCeros.AddItem 931
  lstCeros.AddItem 932
  lstCeros.AddItem 934
  lstCeros.AddItem 935
  lstCeros.AddItem 936
  lstCeros.AddItem 938
  lstCeros.AddItem 939
  lstCeros.AddItem 940
  lstCeros.AddItem 941
  lstCeros.AddItem 942
  lstCeros.AddItem 944
  lstCeros.AddItem 945
  lstCeros.AddItem 947
  lstCeros.AddItem 947
  lstCeros.AddItem 948
  lstCeros.AddItem 950
  lstCeros.AddItem 951
  lstCeros.AddItem 953
  lstCeros.AddItem 954
  lstCeros.AddItem 955
  lstCeros.AddItem 957
  lstCeros.AddItem 958
  lstCeros.AddItem 958
  lstCeros.AddItem 959
  lstCeros.AddItem 962
  lstCeros.AddItem 963
  lstCeros.AddItem 964
  lstCeros.AddItem 965
  lstCeros.AddItem 966
  lstCeros.AddItem 967
  lstCeros.AddItem 969
  lstCeros.AddItem 970
  lstCeros.AddItem 971
  lstCeros.AddItem 973
  lstCeros.AddItem 974
  lstCeros.AddItem 975
  lstCeros.AddItem 976
  lstCeros.AddItem 977
  lstCeros.AddItem 979
  lstCeros.AddItem 981
  lstCeros.AddItem 981
  lstCeros.AddItem 982
  lstCeros.AddItem 984
  lstCeros.AddItem 985
  lstCeros.AddItem 986
  lstCeros.AddItem 987
  lstCeros.AddItem 989
  lstCeros.AddItem 990
  lstCeros.AddItem 991
  lstCeros.AddItem 993
  lstCeros.AddItem 993
  lstCeros.AddItem 994
  lstCeros.AddItem 996
  lstCeros.AddItem 998
  lstCeros.AddItem 999
  lstCeros.AddItem 1000
  lstCeros.AddItem 1001
  lstCeros.AddItem 1002
  lstCeros.AddItem 1003
  lstCeros.AddItem 1005
  lstCeros.AddItem 1006
  lstCeros.AddItem 1008
  lstCeros.AddItem 1009
  lstCeros.AddItem 1010
  lstCeros.AddItem 1011
  lstCeros.AddItem 1012
  lstCeros.AddItem 1013
  lstCeros.AddItem 1015
  lstCeros.AddItem 1016
  lstCeros.AddItem 1017
  lstCeros.AddItem 1019
  lstCeros.AddItem 1020
  lstCeros.AddItem 1021
  lstCeros.AddItem 1022
  lstCeros.AddItem 1023
  lstCeros.AddItem 1025
  lstCeros.AddItem 1026
  lstCeros.AddItem 1027
  lstCeros.AddItem 1028
  lstCeros.AddItem 1029
  lstCeros.AddItem 1031
  lstCeros.AddItem 1032
  lstCeros.AddItem 1033
  lstCeros.AddItem 1035
  lstCeros.AddItem 1036
  lstCeros.AddItem 1037
  lstCeros.AddItem 1038
  lstCeros.AddItem 1039
  lstCeros.AddItem 1040
  lstCeros.AddItem 1042
  lstCeros.AddItem 1044
  lstCeros.AddItem 1045
  lstCeros.AddItem 1045
  lstCeros.AddItem 1047
  lstCeros.AddItem 1048
  lstCeros.AddItem 1049
  lstCeros.AddItem 1050
  lstCeros.AddItem 1052
  lstCeros.AddItem 1053
  lstCeros.AddItem 1055
  lstCeros.AddItem 1055
  lstCeros.AddItem 1057
  lstCeros.AddItem 1057
  lstCeros.AddItem 1059
  lstCeros.AddItem 1060
  lstCeros.AddItem 1062
  lstCeros.AddItem 1063
  lstCeros.AddItem 1064
  lstCeros.AddItem 1065
  lstCeros.AddItem 1066
  lstCeros.AddItem 1067
  lstCeros.AddItem 1068
  lstCeros.AddItem 1071
  lstCeros.AddItem 1072
  lstCeros.AddItem 1073
  lstCeros.AddItem 1074
  lstCeros.AddItem 1075
  lstCeros.AddItem 1076
  lstCeros.AddItem 1077
  lstCeros.AddItem 1079
  lstCeros.AddItem 1080
  lstCeros.AddItem 1081
  lstCeros.AddItem 1083
  lstCeros.AddItem 1083
  lstCeros.AddItem 1084
  lstCeros.AddItem 1086
  lstCeros.AddItem 1087
  lstCeros.AddItem 1089
  lstCeros.AddItem 1090
  lstCeros.AddItem 1091
  lstCeros.AddItem 1092
  lstCeros.AddItem 1093
  lstCeros.AddItem 1094
  lstCeros.AddItem 1095
  lstCeros.AddItem 1096
  lstCeros.AddItem 1099
  lstCeros.AddItem 1099
  lstCeros.AddItem 1101
  lstCeros.AddItem 1102
  lstCeros.AddItem 1103
  lstCeros.AddItem 1104
  lstCeros.AddItem 1106
  lstCeros.AddItem 1107
  lstCeros.AddItem 1108
  lstCeros.AddItem 1109
  lstCeros.AddItem 1110
  lstCeros.AddItem 1111
  lstCeros.AddItem 1112
  lstCeros.AddItem 1113
  lstCeros.AddItem 1115
  lstCeros.AddItem 1117
  lstCeros.AddItem 1118
  lstCeros.AddItem 1119
  lstCeros.AddItem 1119
  lstCeros.AddItem 1121
  lstCeros.AddItem 1122
  lstCeros.AddItem 1123
  lstCeros.AddItem 1125
  lstCeros.AddItem 1126
  lstCeros.AddItem 1128
  lstCeros.AddItem 1128
  lstCeros.AddItem 1130
  lstCeros.AddItem 1130
  lstCeros.AddItem 1131
  lstCeros.AddItem 1134
  lstCeros.AddItem 1135
  lstCeros.AddItem 1136
  lstCeros.AddItem 1137
  lstCeros.AddItem 1138
  lstCeros.AddItem 1139
  lstCeros.AddItem 1141
  lstCeros.AddItem 1141
  lstCeros.AddItem 1143
  lstCeros.AddItem 1145
  lstCeros.AddItem 1145
  lstCeros.AddItem 1147
  lstCeros.AddItem 1148
  lstCeros.AddItem 1149
  lstCeros.AddItem 1150
  lstCeros.AddItem 1152
  lstCeros.AddItem 1153
  lstCeros.AddItem 1154
  lstCeros.AddItem 1155
  lstCeros.AddItem 1157
  lstCeros.AddItem 1157
  lstCeros.AddItem 1158
  lstCeros.AddItem 1159
  lstCeros.AddItem 1161
  lstCeros.AddItem 1162
  lstCeros.AddItem 1164
  lstCeros.AddItem 1165
  lstCeros.AddItem 1165
  lstCeros.AddItem 1167
  lstCeros.AddItem 1168
  lstCeros.AddItem 1170
  lstCeros.AddItem 1170
  lstCeros.AddItem 1172
  lstCeros.AddItem 1173
  lstCeros.AddItem 1174
  lstCeros.AddItem 1175
  lstCeros.AddItem 1177
  lstCeros.AddItem 1177
  lstCeros.AddItem 1180
  lstCeros.AddItem 1181
  lstCeros.AddItem 1181
  lstCeros.AddItem 1183
  lstCeros.AddItem 1184
  lstCeros.AddItem 1185
  lstCeros.AddItem 1186
  lstCeros.AddItem 1187
  lstCeros.AddItem 1189
  lstCeros.AddItem 1190
  lstCeros.AddItem 1191
  lstCeros.AddItem 1192
  lstCeros.AddItem 1193
  lstCeros.AddItem 1194
  lstCeros.AddItem 1196
  lstCeros.AddItem 1197
  lstCeros.AddItem 1199
  lstCeros.AddItem 1199
  lstCeros.AddItem 1201
  lstCeros.AddItem 1202
  lstCeros.AddItem 1203
  lstCeros.AddItem 1204
  lstCeros.AddItem 1205
  lstCeros.AddItem 1207
  lstCeros.AddItem 1208
  lstCeros.AddItem 1209
  lstCeros.AddItem 1210
  lstCeros.AddItem 1211
  lstCeros.AddItem 1212
  lstCeros.AddItem 1214
  lstCeros.AddItem 1215
  lstCeros.AddItem 1216
  lstCeros.AddItem 1217
  lstCeros.AddItem 1219
  lstCeros.AddItem 1220
  lstCeros.AddItem 1221
  lstCeros.AddItem 1222
  lstCeros.AddItem 1223
  lstCeros.AddItem 1225
  lstCeros.AddItem 1226
  lstCeros.AddItem 1227
  lstCeros.AddItem 1228
  lstCeros.AddItem 1229
  lstCeros.AddItem 1231
  lstCeros.AddItem 1232
  lstCeros.AddItem 1233
  lstCeros.AddItem 1234
  lstCeros.AddItem 1236
  lstCeros.AddItem 1236
  lstCeros.AddItem 1238
  lstCeros.AddItem 1238
  lstCeros.AddItem 1239
  lstCeros.AddItem 1241
  lstCeros.AddItem 1243
  lstCeros.AddItem 1244
  lstCeros.AddItem 1245
  lstCeros.AddItem 1246
  lstCeros.AddItem 1247
  lstCeros.AddItem 1248
  lstCeros.AddItem 1249
  lstCeros.AddItem 1251
  lstCeros.AddItem 1252
  lstCeros.AddItem 1254
  lstCeros.AddItem 1254
  lstCeros.AddItem 1255
  lstCeros.AddItem 1256
  lstCeros.AddItem 1258
  lstCeros.AddItem 1259
  lstCeros.AddItem 1260
  lstCeros.AddItem 1262
  lstCeros.AddItem 1263
  lstCeros.AddItem 1264
  lstCeros.AddItem 1265
  lstCeros.AddItem 1266
  lstCeros.AddItem 1267
  lstCeros.AddItem 1268
  lstCeros.AddItem 1270
  lstCeros.AddItem 1271
  lstCeros.AddItem 1272
  lstCeros.AddItem 1273
  lstCeros.AddItem 1274
  lstCeros.AddItem 1275
  lstCeros.AddItem 1277
  lstCeros.AddItem 1278
  lstCeros.AddItem 1279
  lstCeros.AddItem 1280
  lstCeros.AddItem 1282
  lstCeros.AddItem 1283
  lstCeros.AddItem 1283
  lstCeros.AddItem 1285
  lstCeros.AddItem 1286
  lstCeros.AddItem 1287
  lstCeros.AddItem 1289
  lstCeros.AddItem 1290
  lstCeros.AddItem 1290
  lstCeros.AddItem 1292
  lstCeros.AddItem 1293
  lstCeros.AddItem 1294
  lstCeros.AddItem 1295
  lstCeros.AddItem 1297
  lstCeros.AddItem 1298
  lstCeros.AddItem 1299
  lstCeros.AddItem 1300
  lstCeros.AddItem 1301
  lstCeros.AddItem 1302
  lstCeros.AddItem 1303
  lstCeros.AddItem 1305
  lstCeros.AddItem 1307
  lstCeros.AddItem 1307
  lstCeros.AddItem 1309
  lstCeros.AddItem 1309
  lstCeros.AddItem 1311
  lstCeros.AddItem 1312
  lstCeros.AddItem 1313
  lstCeros.AddItem 1314
  lstCeros.AddItem 1316
  lstCeros.AddItem 1317
  lstCeros.AddItem 1318
  lstCeros.AddItem 1319
  lstCeros.AddItem 1320
  lstCeros.AddItem 1322
  lstCeros.AddItem 1322
  lstCeros.AddItem 1324
  lstCeros.AddItem 1325
  lstCeros.AddItem 1326
  lstCeros.AddItem 1328
  lstCeros.AddItem 1329
  lstCeros.AddItem 1329
  lstCeros.AddItem 1330
  lstCeros.AddItem 1332
  lstCeros.AddItem 1334
  lstCeros.AddItem 1335
  lstCeros.AddItem 1336
  lstCeros.AddItem 1337
  lstCeros.AddItem 1338
  lstCeros.AddItem 1339
  lstCeros.AddItem 1340
  lstCeros.AddItem 1341
  lstCeros.AddItem 1343
  lstCeros.AddItem 1344
  lstCeros.AddItem 1345
  lstCeros.AddItem 1346
  lstCeros.AddItem 1348
  lstCeros.AddItem 1348
  lstCeros.AddItem 1349
  lstCeros.AddItem 1351
  lstCeros.AddItem 1352
  lstCeros.AddItem 1353
  lstCeros.AddItem 1354
  lstCeros.AddItem 1356
  lstCeros.AddItem 1357
  lstCeros.AddItem 1358
  lstCeros.AddItem 1358
  lstCeros.AddItem 1360
  lstCeros.AddItem 1361
  lstCeros.AddItem 1363
  lstCeros.AddItem 1364
  lstCeros.AddItem 1365
  lstCeros.AddItem 1365
  lstCeros.AddItem 1367
  lstCeros.AddItem 1368
  lstCeros.AddItem 1370
  lstCeros.AddItem 1371
  lstCeros.AddItem 1372
  lstCeros.AddItem 1373
  lstCeros.AddItem 1374
  lstCeros.AddItem 1375
  lstCeros.AddItem 1376
  lstCeros.AddItem 1377
  lstCeros.AddItem 1380
  lstCeros.AddItem 1380
  lstCeros.AddItem 1381
  lstCeros.AddItem 1382
  lstCeros.AddItem 1383
  lstCeros.AddItem 1384
  lstCeros.AddItem 1386
  lstCeros.AddItem 1387
  lstCeros.AddItem 1388
  lstCeros.AddItem 1390
  lstCeros.AddItem 1391
  lstCeros.AddItem 1392
  lstCeros.AddItem 1393
  lstCeros.AddItem 1393
  lstCeros.AddItem 1395
  lstCeros.AddItem 1397
  lstCeros.AddItem 1398
  lstCeros.AddItem 1399
  lstCeros.AddItem 1400
  lstCeros.AddItem 1400
  lstCeros.AddItem 1403
  lstCeros.AddItem 1403
  lstCeros.AddItem 1404
  lstCeros.AddItem 1406
  lstCeros.AddItem 1407
  lstCeros.AddItem 1408
  lstCeros.AddItem 1409
  lstCeros.AddItem 1410
  lstCeros.AddItem 1411
  lstCeros.AddItem 1412
  lstCeros.AddItem 1414
  lstCeros.AddItem 1416
  lstCeros.AddItem 1416
  lstCeros.AddItem 1417
  lstCeros.AddItem 1419
  lstCeros.AddItem 1419
  lstCeros.AddItem 1420
  lstCeros.AddItem 1422
  lstCeros.AddItem 1422
  lstCeros.AddItem 1424
  lstCeros.AddItem 1426
  lstCeros.AddItem 1427
  lstCeros.AddItem 1427
  lstCeros.AddItem 1429
  lstCeros.AddItem 1430
  lstCeros.AddItem 1431
  lstCeros.AddItem 1432
  lstCeros.AddItem 1434
  lstCeros.AddItem 1435
  lstCeros.AddItem 1436
  lstCeros.AddItem 1437
  lstCeros.AddItem 1438
  lstCeros.AddItem 1439
  lstCeros.AddItem 1440
  lstCeros.AddItem 1442
  lstCeros.AddItem 1443
  lstCeros.AddItem 1444
  lstCeros.AddItem 1445
  lstCeros.AddItem 1446
  lstCeros.AddItem 1447
  lstCeros.AddItem 1448
  lstCeros.AddItem 1449
  lstCeros.AddItem 1451
  lstCeros.AddItem 1452
  lstCeros.AddItem 1454
  lstCeros.AddItem 1454
  lstCeros.AddItem 1455
  lstCeros.AddItem 1456
  lstCeros.AddItem 1458
  lstCeros.AddItem 1458
  lstCeros.AddItem 1461
  lstCeros.AddItem 1461
  lstCeros.AddItem 1462
  lstCeros.AddItem 1463
  lstCeros.AddItem 1465
  lstCeros.AddItem 1466
  lstCeros.AddItem 1467
  lstCeros.AddItem 1468
  lstCeros.AddItem 1469
  lstCeros.AddItem 1471
  lstCeros.AddItem 1471
  lstCeros.AddItem 1473
  lstCeros.AddItem 1474
  lstCeros.AddItem 1474
  lstCeros.AddItem 1476
  lstCeros.AddItem 1477
  lstCeros.AddItem 1478
  lstCeros.AddItem 1480
  lstCeros.AddItem 1481
  lstCeros.AddItem 1482
  lstCeros.AddItem 1483
  lstCeros.AddItem 1484
  lstCeros.AddItem 1485
  lstCeros.AddItem 1486
  lstCeros.AddItem 1488
  lstCeros.AddItem 1489
  lstCeros.AddItem 1490
  lstCeros.AddItem 1491
  lstCeros.AddItem 1492
  lstCeros.AddItem 1493
  lstCeros.AddItem 1494
  lstCeros.AddItem 1496
  lstCeros.AddItem 1497
  lstCeros.AddItem 1498
End Sub
