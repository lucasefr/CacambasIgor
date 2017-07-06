VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPesquisa 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9885
   Icon            =   "frmPesquisa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   735
      Left            =   8745
      TabIndex        =   2
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdPesquisa 
      Caption         =   "Seleciona"
      Height          =   735
      Left            =   7560
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid flxPesquisa 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   10186
      _Version        =   393216
      BackColorBkg    =   12648447
      GridColor       =   16777215
   End
End
Attribute VB_Name = "frmPesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vPosicaoColunaID                 As Integer
Public vColunaID                        As Integer

Public Sub carregaGridPesquisa(pNumeroColunas As Integer, pstrSQL As String, pNomeColunas As String, pPosicaoColunaID)
                
        Dim rsPesquisa                  As ADODB.Recordset
        Dim vField                      As ADODB.Field
        Dim PosicaoColuna               As Integer
        
        Set rsPesquisa = New ADODB.Recordset
        rsPesquisa.Open pstrSQL, dbCacamba, adOpenStatic, adLockReadOnly
        
        If rsPesquisa.EOF Then
                Unload Me
        End If
        
        With flxPesquisa
                .FormatString = vbTab & Replace(pNomeColunas, ",", "|", , , vbTextCompare)
                .Cols = rsPesquisa.Fields.Count + 1
                .Rows = 1
                .Redraw = False
                
                Do While Not rsPesquisa.EOF
                        
                        .Rows = .Rows + 1
                        
                        PosicaoColuna = 1
                        
                        For Each vField In rsPesquisa.Fields
                                Select Case vField.Type
                                        Case DataTypeEnum.adDate, DataTypeEnum.adDBTimeStamp
                                                .TextMatrix(.Rows - 1, PosicaoColuna) = Format(vField.Value, "dd/mm/YYYY")
                                        Case DataTypeEnum.adCurrency
                                                .TextMatrix(.Rows - 1, PosicaoColuna) = Format(vField.Value, "###,###,###,##0.00")
                                        Case Else
                                                .TextMatrix(.Rows - 1, PosicaoColuna) = NVL(vField.Value, "")
                                End Select
                                If .ColWidth(PosicaoColuna) < Len("" & vField.Value) * 125 Then
                                .ColWidth(PosicaoColuna) = Len("" & vField.Value) * 125
                                End If
                                If .ColWidth(PosicaoColuna) < 240 Then .ColWidth(PosicaoColuna) = 240
                                If .ColWidth(PosicaoColuna) > 6000 And PosicaoColuna <= 2 Then .ColWidth(PosicaoColuna) = 6000
                
                                PosicaoColuna = PosicaoColuna + 1
                        Next vField

                        rsPesquisa.MoveNext
                Loop
                .Redraw = True
                .SelectionMode = flexSelectionByRow
        End With
        
        vPosicaoColunaID = pPosicaoColunaID
End Sub

Private Sub cmdPesquisa_Click()
        
        With flxPesquisa
                If .TextMatrix(.Row, vPosicaoColunaID) = "" Then
                        MsgBox "Selecione um item da pesquisa!", vbInformation, "Caçamba"
                        Exit Sub
                End If
                
                vgRetornoConsulta = .TextMatrix(.Row, vPosicaoColunaID)
        End With
        
        Unload Me
        
End Sub

Private Sub cmdSair_Click()
        vgRetornoConsulta = 0
        Unload Me
End Sub

