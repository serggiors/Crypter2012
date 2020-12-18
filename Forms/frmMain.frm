VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   Caption         =   "Crypter_by_zRG3R"
   ClientHeight    =   2250
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5100
   FillColor       =   &H000000FF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtkey 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Text            =   "Contraseña"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdProteger 
      Caption         =   "Proteger"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "..."
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtarchivo 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4320
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "RED PILL"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   """Crypter zRG3R"""
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBuscar_Click()


With CD
        .DialogTitle = "Seleccione el archivo a encriptar!"
        .Filter = "Aplicaciones EXE|*.exe"
        .ShowOpen
        End With
        
        If Not CD.FileName = vbNullString Then
        txtarchivo.Text = CD.FileName
        MsgBox "Archivo cargado correctamente!", vbInformation, Me.Caption
        End If
End Sub

Private Sub cmdProteger_Click()
Dim Stub As String, Archivo As String


If txtarchivo.Text = vbNullString Then
MsgBox "Primero debe cargar un archivo para encriptar!", vbExclamation, Me.Caption 'Mostramos un mensaje de exclamacion
Exit Sub
Else

Open App.Path & "\Stub.exe" For Binary As #1
Stub = Space(LOF(1))
Get #1, , Stub
Close #1

Open txtarchivo.Text For Binary As #1
Archivo = Space(LOF(1))
Get #1, , Archivo
Close #1


With CD
        .DialogTitle = "Selecione la ruta donde guardar el archivo encriptado!"
        .Filter = "Aplicaciones EXE|*.exe"
        .ShowSave
        End With
        
        If Not CD.FileName = vbNullString Then
        
        Archivo = RC4(Archivo, txtkey.Text)
        
        Open CD.FileName For Binary As #1
        Put #1, , Stub & "##$$##" & Archivo & "##$$##" & txtkey.Text & "##$$##" 'Y Ahora tambien ponemos la contraseña en el Archivo 'Aqui lo que hacemos es meter los datos del Stub al Archivo y tambien metemos el Archivo encriptado separado por Unos Splits k en este caso son "##$$##" para poder separar mas adelante los datos en el Stub !
        Close #1
        
        MsgBox "Archivo Encriptado Correctamente!", vbInformation, Me.Caption
        End If
End If
End Sub

Public Function RC4(ByVal Data As String, ByVal Password As String) As String
On Error Resume Next
Dim F(0 To 255) As Integer, X, Y As Long, Key() As Byte
Key() = StrConv(Password, vbFromUnicode)
For X = 0 To 255
    Y = (Y + F(X) + Key(X Mod Len(Password))) Mod 256
    F(X) = X
Next X
Key() = StrConv(Data, vbFromUnicode)
For X = 0 To Len(Data)
    Y = (Y + F(Y) + 1) Mod 256
    Key(X) = Key(X) Xor F(Temp + F((Y + F(Y)) Mod 254))
Next X
RC4 = StrConv(Key, vbUnicode)
End Function

