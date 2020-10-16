VERSION 5.00
Object = "{61159A24-3E03-4E76-9CA9-2396C6822B8F}#1.0#0"; "ChamaleonBtn.ocx"
Begin VB.Form AtualizaII 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualizar balança"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2490
   Icon            =   "AtualizaII.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   2490
   StartUpPosition =   2  'CenterScreen
   Begin ChamaleonBtn.chameleonButton Botao 
      Height          =   450
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   794
      BTYPE           =   5
      TX              =   "Atualizar dados"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "AtualizaII.frx":0CCA
      PICN            =   "AtualizaII.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "AtualizaII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mconn As ADODB.Connection
Dim tb1 As ADODB.Recordset

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long


Private Sub Botao_Click(Index As Integer)
'On Error GoTo err

Dim valorProd As Double
Dim SQL As String, prodSemPonto As String
   
   If Dir("C:\Filizola", vbDirectory) = Empty Then
      MkDir "C:\Filizola"
   End If
   
   Set mconn = New ADODB.Connection
   Set tb1 = New Recordset
    
   With mconn
     .Provider = "SIBPROvider.2"
     .ConnectionString = "Data Source=" & ReadINI("PATH", "LOCALDB") & "\Banco.FDB; User ID=" & "SYSDBA" & "; Password=" & "masterkey"""
     .Open
   End With
 
   If Dir("c:\Filizola\CADTXT.TXT") <> "" Then Kill "c:\Filizola\CADTXT.TXT"
   If Dir("c:\Filizola\SETORTXT.TXT") <> "" Then Kill "c:\Filizola\SETORTXT.TXT"
  
   Me.Enabled = False
   Screen.MousePointer = 11
 
   Open "c:\Filizola\CADTXT.TXT" For Output As #2
   Open "c:\Filizola\SETORTXT.TXT" For Output As #3
   
   If tb1.State = 1 Then
      tb1.Close
   End If
     
   tb1.Open "Select campo0, campo1, campo2, coalesce(campo6,0) as campo6, COALESCE(CAMPO47,5) as campo47, coalesce(CAMPO61,0) as campo61, coalesce(CAMPO62, dateadd(-1 DAY to CURRENT_DATE)) as campo62, coalesce(CAMPO55,'') as campo55 from produto where campo46=1", mconn, adOpenStatic, adLockReadOnly
  
   Do While Not tb1.EOF
      DoEvents
        
      valorProd = IIf(tb1!campo62 >= Date, tb1!CAMPO61, tb1!campo6)
        
      SQL = IIf(tb1!CAMPO55 = "", Left(tb1!CAMPO1, 17) & " " & tb1!CAMPO2, tb1!CAMPO55)
      SQL = SQL & Space(22 - Len(SQL))
      
      prodSemPonto = Format(valorProd, "#,###,###,###0.00")
      prodSemPonto = Replace(prodSemPonto, ".", "")
      prodSemPonto = Replace(prodSemPonto, ",", "")
      prodSemPonto = Format(prodSemPonto, "0000000")
        
      Print #2, Format(tb1!CAMPO0, "000000") & IIf(tb1!CAMPO2 = "KG", "P", "U") & SQL & prodSemPonto & Format(tb1!CAMPO47, "000")
      Print #3, "A           " & Format(tb1!CAMPO0, "000000") & "0000001"
    
      tb1.MoveNext
    Loop
    Close #2
    Close #3
    
    tb1.Close
    mconn.Close
    
    MsgBox "Exportação concluida!"

 Me.Enabled = True
 Screen.MousePointer = 0

Exit Sub
err:
  Close #2
  Close #3
  Me.Enabled = True
  Screen.MousePointer = 0
  MsgBox err.Number & vbCrLf & err.Description, vbCritical, "Atenção"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Public Function ReadINI(Secao As String, Entrada As String)
 Dim retlen As String
 Dim Ret As String
 Ret = String$(255, 0)
 retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), App.Path & "\config.ini")
 Ret = Left$(Ret, retlen)
 ReadINI = UCase(Ret)
End Function

