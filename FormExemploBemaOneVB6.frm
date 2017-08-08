VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Exemplo BemaOne VB6"
   ClientHeight    =   8505
   ClientLeft      =   3285
   ClientTop       =   3240
   ClientWidth     =   15600
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   15600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnEfetuarConfiguracoes 
      Caption         =   "Configurar"
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   2400
      Width           =   2775
   End
   Begin VB.CommandButton btnCancelaCupom 
      Caption         =   "Cancelar Cupom"
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton CommandNotaErro 
      BackColor       =   &H000000FF&
      Caption         =   "Emitir nota com erro "
      DownPicture     =   "FormExemploBemaOneVB6.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MaskColor       =   &H0000FFFF&
      Picture         =   "FormExemploBemaOneVB6.frx":E625
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consulta de Notas"
      Height          =   3135
      Left            =   120
      TabIndex        =   17
      Top             =   5040
      Width           =   3255
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Text            =   "Text6"
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "PDF"
         Height          =   555
         Left            =   1320
         TabIndex        =   22
         Top             =   1920
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "JSON"
         Height          =   555
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Text            =   "Text4"
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton ConsultaNota 
         Caption         =   "Consultar Nota"
         Height          =   555
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Label Label8 
         Caption         =   "Chave de Acesso"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Série"
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Número "
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   8280
      TabIndex        =   16
      Text            =   "002"
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   5520
      TabIndex        =   13
      Text            =   "189"
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton CommandExecutar 
      Caption         =   "Executar Comando de Venda"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   2775
   End
   Begin VB.ComboBox ComboComando 
      Height          =   315
      ItemData        =   "FormExemploBemaOneVB6.frx":2764F
      Left            =   240
      List            =   "FormExemploBemaOneVB6.frx":27651
      TabIndex        =   10
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox TextJson 
      BackColor       =   &H00C0FFC0&
      Height          =   3015
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1440
      Width           =   7095
   End
   Begin VB.Frame TextoGenerico 
      Caption         =   "Imprimir Texto Genérico"
      Height          =   7095
      Left            =   11040
      TabIndex        =   6
      Top             =   360
      Width           =   4335
      Begin VB.CheckBox CheckCut 
         Caption         =   "Guilhotina"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   6000
         Width           =   1095
      End
      Begin VB.TextBox TextSaida 
         BackColor       =   &H00C0FFFF&
         Height          =   5415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "FormExemploBemaOneVB6.frx":27653
         Top             =   360
         Width           =   3975
      End
      Begin VB.CommandButton TextoLivre 
         Caption         =   "Texto Livre"
         Height          =   495
         Left            =   1680
         TabIndex        =   7
         Top             =   6360
         Width           =   2295
      End
   End
   Begin VB.CommandButton Sair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   13440
      TabIndex        =   5
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton ObterStatusImpressora 
      Caption         =   "Obter Status da Impressora"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton ObterInfoSistema 
      Caption         =   "Obter Informações do Sistema"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox TextRetorno 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FFFF&
      Height          =   2895
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   5280
      Width           =   7095
   End
   Begin VB.CommandButton ListarConfiguracoes 
      Caption         =   "Listar Configurações"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Número da nota"
      Height          =   255
      Left            =   6960
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Série da Nota"
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Envio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6600
      TabIndex        =   12
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Retorno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   4800
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Bematech_Fiscal_AbrirNota Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_EstornarNota Lib "BemaOne32.dll" () As Long
Private Declare Function Bematech_Fiscal_FecharNota Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_VenderItem Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_EstornarVendaItem Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_EfetuarPagamento Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_EstornarPagamento Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_ListarNotas Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_InutilizarNumeracao Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_ConsultarNota Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_EnviarNotaEmail Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_CancelarNota Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_ObterStatusImpressora Lib "BemaOne32.dll" () As Long
Private Declare Function Bematech_Fiscal_ImprimirTextoLivre Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_ImprimirDocumentoFiscal Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_AcionarGaveta Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_ObterInformacoesSistema Lib "BemaOne32.dll" () As Long
Private Declare Function Bematech_Fiscal_ListarConfiguracoes Lib "BemaOne32.dll" () As Long
Private Declare Function Bematech_Fiscal_EfetuarConfiguracoes Lib "BemaOne32.dll" (ByVal dados As String) As Long
Private Declare Function Bematech_Fiscal_ObterInformacoesContingencia Lib "BemaOne32.dll" () As Long
Private Declare Function Bematech_Fiscal_TrocarEstadoContingencia Lib "BemaOne32.dll" () As Long
Private Declare Function ConvCStringToVBString Lib "kernel32" Alias "lstrcpyA" (ByVal lpsz As String, ByVal pt As Long) As Long
Public dataHora As String
Dim sReturn, sReturn2 As Long
Dim sFunctionReturn, sFunctionReturn2, json As String

'Segue abaixo o método para conversão de Base64 caso necessite... neste programa, utilizei para passar o comando
'de acionamento de guilhotina

'**********************************************************************************************************
' A Base64 Encoder/Decoder.
'
' This module is used to encode and decode data in Base64 format as described in RFC 1521.
'
' Home page: www.source-code.biz.
' License: GNU/LGPL (www.gnu.org/licenses/lgpl.html).
' Copyright 2007: Christian d'Heureuse, Inventec Informatik AG, Switzerland.
' This module is provided "as is" without warranty of any kind.


Option Explicit

Private InitDone  As Boolean
Private Map1(0 To 63)  As Byte
Private Map2(0 To 127) As Byte

' Encodes a string into Base64 format.
' No blanks or line breaks are inserted.
' Parameters:
'   S         a String to be encoded.
' Returns:    a String with the Base64 encoded data.
Public Function Base64EncodeString(ByVal s As String) As String
   Base64EncodeString = Base64Encode(ConvertStringToBytes(s))
   End Function

' Encodes a byte array into Base64 format.
' No blanks or line breaks are inserted.
' Parameters:
'   InData    an array containing the data bytes to be encoded.
' Returns:    a string with the Base64 encoded data.
Public Function Base64Encode(InData() As Byte)
   Base64Encode = Base64Encode2(InData, UBound(InData) - LBound(InData) + 1)
   End Function

' Encodes a byte array into Base64 format.
' No blanks or line breaks are inserted.
' Parameters:
'   InData    an array containing the data bytes to be encoded.
'   InLen     number of bytes to process in InData.
' Returns:    a string with the Base64 encoded data.
Public Function Base64Encode2(InData() As Byte, ByVal InLen As Long) As String
   If Not InitDone Then Init
   If InLen = 0 Then Base64Encode2 = "": Exit Function
   Dim ODataLen As Long: ODataLen = (InLen * 4 + 2) \ 3     ' output length without padding
   Dim OLen As Long: OLen = ((InLen + 2) \ 3) * 4           ' output length including padding
   Dim Out() As Byte
   ReDim Out(0 To OLen - 1) As Byte
   Dim ip0 As Long: ip0 = LBound(InData)
   Dim ip As Long
   Dim op As Long
   Do While ip < InLen
      Dim i0 As Byte: i0 = InData(ip0 + ip): ip = ip + 1
      Dim i1 As Byte: If ip < InLen Then i1 = InData(ip0 + ip): ip = ip + 1 Else i1 = 0
      Dim i2 As Byte: If ip < InLen Then i2 = InData(ip0 + ip): ip = ip + 1 Else i2 = 0
      Dim o0 As Byte: o0 = i0 \ 4
      Dim o1 As Byte: o1 = ((i0 And 3) * &H10) Or (i1 \ &H10)
      Dim o2 As Byte: o2 = ((i1 And &HF) * 4) Or (i2 \ &H40)
      Dim o3 As Byte: o3 = i2 And &H3F
      Out(op) = Map1(o0): op = op + 1
      Out(op) = Map1(o1): op = op + 1
      Out(op) = IIf(op < ODataLen, Map1(o2), Asc("=")): op = op + 1
      Out(op) = IIf(op < ODataLen, Map1(o3), Asc("=")): op = op + 1
      Loop
   Base64Encode2 = ConvertBytesToString(Out)
   End Function

' Decodes a string from Base64 format.
' Parameters:
'    s        a Base64 String to be decoded.
' Returns     a String containing the decoded data.
Public Function Base64DecodeString(ByVal s As String) As String
   If s = "" Then Base64DecodeString = "": Exit Function
   Base64DecodeString = ConvertBytesToString(Base64Decode(s))
   End Function

' Decodes a byte array from Base64 format.
' Parameters
'   s         a Base64 String to be decoded.
' Returns:    an array containing the decoded data bytes.
Public Function Base64Decode(ByVal s As String) As Byte()
   If Not InitDone Then Init
   Dim IBuf() As Byte: IBuf = ConvertStringToBytes(s)
   Dim ILen As Long: ILen = UBound(IBuf) + 1
   If ILen Mod 4 <> 0 Then Err.Raise vbObjectError, , "Length of Base64 encoded input string is not a multiple of 4."
   Do While ILen > 0
      If IBuf(ILen - 1) <> Asc("=") Then Exit Do
      ILen = ILen - 1
      Loop
   Dim OLen As Long: OLen = (ILen * 3) \ 4
   Dim Out() As Byte
   ReDim Out(0 To OLen - 1) As Byte
   Dim ip As Long
   Dim op As Long
   Do While ip < ILen
      Dim i0 As Byte: i0 = IBuf(ip): ip = ip + 1
      Dim i1 As Byte: i1 = IBuf(ip): ip = ip + 1
      Dim i2 As Byte: If ip < ILen Then i2 = IBuf(ip): ip = ip + 1 Else i2 = Asc("A")
      Dim i3 As Byte: If ip < ILen Then i3 = IBuf(ip): ip = ip + 1 Else i3 = Asc("A")
      If i0 > 127 Or i1 > 127 Or i2 > 127 Or i3 > 127 Then _
         Err.Raise vbObjectError, , "Illegal character in Base64 encoded data."
      Dim b0 As Byte: b0 = Map2(i0)
      Dim b1 As Byte: b1 = Map2(i1)
      Dim b2 As Byte: b2 = Map2(i2)
      Dim b3 As Byte: b3 = Map2(i3)
      If b0 > 63 Or b1 > 63 Or b2 > 63 Or b3 > 63 Then _
         Err.Raise vbObjectError, , "Illegal character in Base64 encoded data."
      Dim o0 As Byte: o0 = (b0 * 4) Or (b1 \ &H10)
      Dim o1 As Byte: o1 = ((b1 And &HF) * &H10) Or (b2 \ 4)
      Dim o2 As Byte: o2 = ((b2 And 3) * &H40) Or b3
      Out(op) = o0: op = op + 1
      If op < OLen Then Out(op) = o1: op = op + 1
      If op < OLen Then Out(op) = o2: op = op + 1
      Loop
   Base64Decode = Out
   End Function

Private Sub Init()
   Dim c As Integer, i As Integer
   ' set Map1
   i = 0
   For c = Asc("A") To Asc("Z"): Map1(i) = c: i = i + 1: Next
   For c = Asc("a") To Asc("z"): Map1(i) = c: i = i + 1: Next
   For c = Asc("0") To Asc("9"): Map1(i) = c: i = i + 1: Next
   Map1(i) = Asc("+"): i = i + 1
   Map1(i) = Asc("/"): i = i + 1
   ' set Map2
   For i = 0 To 127: Map2(i) = 255: Next
   For i = 0 To 63: Map2(Map1(i)) = i: Next
   InitDone = True
   End Sub

Private Function ConvertStringToBytes(ByVal s As String) As Byte()
   Dim b1() As Byte: b1 = s
   Dim l As Long: l = (UBound(b1) + 1) \ 2
   If l = 0 Then ConvertStringToBytes = b1: Exit Function
   Dim b2() As Byte
   ReDim b2(0 To l - 1) As Byte
   Dim p As Long
   For p = 0 To l - 1
      Dim c As Long: c = b1(2 * p) + 256 * CLng(b1(2 * p + 1))
      If c >= 256 Then c = Asc("?")
      b2(p) = c
      Next
   ConvertStringToBytes = b2
   End Function

Private Function ConvertBytesToString(b() As Byte) As String
   Dim l As Long: l = UBound(b) - LBound(b) + 1
   Dim b2() As Byte
   ReDim b2(0 To (2 * l) - 1) As Byte
   Dim p0 As Long: p0 = LBound(b)
   Dim p As Long
   For p = 0 To l - 1: b2(2 * p) = b(p0 + p): Next
   Dim s As String: s = b2
   ConvertBytesToString = s
   End Function
'*****************************************************************************************************************************


Public Function GetStringFromPointer(ByVal lpString As Long) As String
    Dim NullCharPos As Long
    Dim szBuffer As String
    szBuffer = String(10000, 0)
    ConvCStringToVBString szBuffer, lpString
    NullCharPos = InStr(szBuffer, vbNullChar)
    
    If NullCharPos = 0 Then
        GetStringFromPointer = szBuffer
    Else
        GetStringFromPointer = Left(szBuffer, NullCharPos - 1)
    End If
    
End Function


Private Sub btnCancelaCupom_Click()
                Dim IdUltimoCupomSatEmitido As String
                IdUltimoCupomSatEmitido = Text6.Text
                json = "{""id"": """ & IdUltimoCupomSatEmitido & """, "
                json = json & """identificacao"": {"
                json = json & """cnpj"": """ & "16716114000172" & ""","
                json = json & """numeroCaixa"": ""001"","
                json = json & """signAC"": ""SGR-SAT SISTEMA DE GESTAO E RETAGUARDA DO SAT"""
                json = json & "},"
                json = json & """destinatario"": {"
                json = json & """cpf"": ""06123800922""}"
                json = json & "}"
                TextJson.Text = json
                
                sReturn = Bematech_Fiscal_CancelarNota(json)
                sFunctionReturn2 = GetStringFromPointer(sReturn)
                TextRetorno.Text = sFunctionReturn + "*************************************************" + vbCrLf
                TextRetorno.Text = TextRetorno.Text & "***************Status da Nota***************" + vbCrLf + sFunctionReturn2
End Sub

Private Sub btnEfetuarConfiguracoes_Click()
Dim json As String

json = json & "{"
json = json & """sistema"":"
json = json & "{"
json = json & """path"": ""C:"","
json = json & """nivelLog"": ""1"""
json = json & "},"
json = json & """nfe"":"
json = json & "{"
json = json & """timeoutWebservice"": """""
json = json & "}"
json = json & "}"

sReturn = Bematech_Fiscal_EfetuarConfiguracoes(json)
sFunctionReturn2 = GetStringFromPointer(sReturn)
TextRetorno.Text = sFunctionReturn + "*************************************************" + vbCrLf
TextRetorno.Text = TextRetorno.Text & "***************Configuracoes***************" + vbCrLf + sFunctionReturn2


End Sub

Private Sub ComboComando_Click()

'Se o json for do AbrirNota
If ComboComando.ListIndex = 0 Then
TextJson.Text = "{"
TextJson.Text = TextJson.Text & """identificacao"": {"
TextJson.Text = TextJson.Text & """cnpj"": ""16716114000172"","
TextJson.Text = TextJson.Text & """numeroCaixa"": ""001"","
TextJson.Text = TextJson.Text & """signAC"": ""SGR-SAT SISTEMA DE GESTAO E RETAGUARDA DO SAT"""
TextJson.Text = TextJson.Text & "},"

TextJson.Text = TextJson.Text & """emitente"": {"
TextJson.Text = TextJson.Text & """cnpj"": ""82373077000171"","
TextJson.Text = TextJson.Text & """ie"": ""111111111111"","
TextJson.Text = TextJson.Text & """indRatISSQN"": ""S"""
TextJson.Text = TextJson.Text & "},"

TextJson.Text = TextJson.Text & """destinatario"": {"
TextJson.Text = TextJson.Text & """cpf"": ""06123800922"","
TextJson.Text = TextJson.Text & """xNome"": ""Manoel Gustavo Aguiar Dutra"""
TextJson.Text = TextJson.Text & "},"

TextJson.Text = TextJson.Text & """entrega"": {"
TextJson.Text = TextJson.Text & """cpf"": ""06123800922"","
TextJson.Text = TextJson.Text & """endereco"": {"
TextJson.Text = TextJson.Text & """xLgr"": ""Rua das Alfafas"","
TextJson.Text = TextJson.Text & """nro"": ""450"","
TextJson.Text = TextJson.Text & """xCpl"": ""Bloco 14 ap 11"","
TextJson.Text = TextJson.Text & """xBairro"": ""Campo Comprido"","
TextJson.Text = TextJson.Text & """xMun"": ""Curitiba"","
TextJson.Text = TextJson.Text & """uf"": ""PR"""
TextJson.Text = TextJson.Text & "}"
TextJson.Text = TextJson.Text & "}"
TextJson.Text = TextJson.Text & "}"

End If

'Se o json for do VenderItem
If ComboComando.ListIndex = 1 Then
TextJson.Text = "{"
TextJson.Text = TextJson.Text & """produto"": {"
TextJson.Text = TextJson.Text & """cean"": ""7897238304177"","
TextJson.Text = TextJson.Text & """ncm"": ""12345678"","
TextJson.Text = TextJson.Text & """cfop"": ""5101"","
TextJson.Text = TextJson.Text & """indTot"": 1,"
TextJson.Text = TextJson.Text & """vUnCom"": 1.000,"
TextJson.Text = TextJson.Text & """uTrib"": ""UN"","
TextJson.Text = TextJson.Text & """vUnTrib"": ""1.000"","
TextJson.Text = TextJson.Text & """cProd"": ""85258029901234"","
TextJson.Text = TextJson.Text & """xProd"": ""Agua Mineral"","
TextJson.Text = TextJson.Text & """uCom"": ""UN"","
TextJson.Text = TextJson.Text & """qTrib"": 1.000,"
TextJson.Text = TextJson.Text & """qCom"": ""1.0000"","
TextJson.Text = TextJson.Text & """vProd"": 1.00,"
TextJson.Text = TextJson.Text & """indRegra"": ""A"","
TextJson.Text = TextJson.Text & """vDesc"": 0.00,"
TextJson.Text = TextJson.Text & """vOutro"": 0.00,"
TextJson.Text = TextJson.Text & """qCom"": ""1.0000"""
TextJson.Text = TextJson.Text & "},"
TextJson.Text = TextJson.Text & """imposto"": {"
TextJson.Text = TextJson.Text & """icms"": {"
TextJson.Text = TextJson.Text & """icms00"": {"
TextJson.Text = TextJson.Text & """orig"": 0,"
TextJson.Text = TextJson.Text & """cst"": ""00"","
TextJson.Text = TextJson.Text & """picms"": 0.00"
TextJson.Text = TextJson.Text & "}"
TextJson.Text = TextJson.Text & "},"
TextJson.Text = TextJson.Text & """cofins"": {"
TextJson.Text = TextJson.Text & """cofinsnt"": {"
TextJson.Text = TextJson.Text & """cst"": ""08"""
TextJson.Text = TextJson.Text & "}"
TextJson.Text = TextJson.Text & "},"
TextJson.Text = TextJson.Text & """pis"": {"
TextJson.Text = TextJson.Text & """pisnt"": {"
TextJson.Text = TextJson.Text & """cst"": ""08"""
TextJson.Text = TextJson.Text & "}"
TextJson.Text = TextJson.Text & "}"
TextJson.Text = TextJson.Text & "},"
TextJson.Text = TextJson.Text & """vItem12741"": ""1"","
TextJson.Text = TextJson.Text & """nItem"": 1"
TextJson.Text = TextJson.Text & "}"


End If

'Se o json for do EfetuarPagamento
If ComboComando.ListIndex = 2 Then
TextJson.Text = "{   ""tPag"": 1,   ""vPag"": 1.00 }"
End If

'Se o json for do FecharNota
If ComboComando.ListIndex = 3 Then
TextJson.Text = "{"
TextJson.Text = TextJson.Text & """total"": {"
TextJson.Text = TextJson.Text & """vCFeLei12741"": ""0.00"""
TextJson.Text = TextJson.Text & "},"
TextJson.Text = TextJson.Text & """informacaoAdicional"": {"
TextJson.Text = TextJson.Text & """infCpl"": ""Sequencia 003 Nota sem Cliente"""
TextJson.Text = TextJson.Text & "}"
TextJson.Text = TextJson.Text & "}"

End If

End Sub

Private Sub CommandExecutar_Click()
Dim sReturn, sReturn2 As Long
Dim sFunctionReturn, sFunctionReturn2 As String
Dim json As String

'Se escolhido AbrirNota no ComboBox, carregar json no TextJson
If ComboComando.ListIndex = 0 Then
    sReturn = Bematech_Fiscal_AbrirNota(TextJson.Text)
    sFunctionReturn = GetStringFromPointer(sReturn)
    TextRetorno.Text = sFunctionReturn
End If

'Se escolhido VenderItem no ComboBox, carregar json no TextJson
If ComboComando.ListIndex = 1 Then
    sReturn = Bematech_Fiscal_VenderItem(TextJson.Text)
    sFunctionReturn = GetStringFromPointer(sReturn)
    TextRetorno.Text = sFunctionReturn
End If

'Se escolhido EfetuarPagamento no ComboBox, carregar json no TextJson
If ComboComando.ListIndex = 2 Then
    sReturn = Bematech_Fiscal_EfetuarPagamento(TextJson.Text)
    sFunctionReturn = GetStringFromPointer(sReturn)
    TextRetorno.Text = sFunctionReturn
End If

'Se escolhido FecharNota no ComboBox, carregar json no TextJson
If ComboComando.ListIndex = 3 Then
    sReturn = Bematech_Fiscal_FecharNota(TextJson.Text)
    sFunctionReturn = GetStringFromPointer(sReturn)
    
    json = "{ "
    json = json & """id"": """ + Text2.Text + ""","
    json = json & """formato"": """ + Text3.Text + """ "
    json = json & "}"
  
    sReturn2 = Bematech_Fiscal_ConsultarNota(json)
    sFunctionReturn2 = GetStringFromPointer(sReturn2)
    TextRetorno.Text = sFunctionReturn + "*************************************************" + vbCrLf
    TextRetorno.Text = TextRetorno.Text & "***************Status da Nota***************" + vbCrLf + sFunctionReturn2
    
End If

End Sub

Private Sub CommandNotaErro_Click()
Dim sReturn, sReturn2 As Long
Dim sFunctionReturn, sFunctionReturn2, json As String

json = "{"
json = json & """versao"": ""3.10"","
json = json & """configuracao"": {"
json = json & """imprimir"": true,"
json = json & """email"": false"
json = json & "},"
json = json & """identificacao"": {"
json = json & """cuf"": ""41"","
json = json & """cnf"": ""0000" + Text3.Text + ""","
json = json & """natOp"": ""VENDA"","
json = json & """indPag"": 0,"
json = json & """mod"": ""65"","
json = json & """serie"": """ + Text2.Text + ""","
json = json & """nnf"": """ + Text3.Text + ""","
json = json & """dhEmi"": """ + dataHora + ""","
json = json & """tpNF"": ""1"","
json = json & """idDest"": 1,"
json = json & """tpImp"": 4,"
json = json & """tpEmis"": 1,"
json = json & """cdv"": 8,"
json = json & """tpAmb"": 2,"
json = json & """finNFe"": 1,"
json = json & """indFinal"": 1,"
json = json & """indPres"": 1,"
json = json & """procEmi"": 0,"
json = json & """verProc"": ""1.0.0.0"","
json = json & """cMunFG"": ""4106902"""
json = json & "},"
json = json & """emitente"": {"
json = json & """cnpj"": ""82373077000171"","
json = json & """endereco"": {"
json = json & """nro"": ""1341"","
json = json & """uf"": ""PR"","
json = json & """cep"": ""81320400"","
json = json & """fone"": ""4184848484"","
json = json & """xBairro"": ""Jardim Botânico"","
json = json & """xLgr"": ""AV Comendador Franco"","
json = json & """cMun"": ""4106902"","
json = json & """cPais"": ""1058"","
json = json & """xPais"": ""BRASIL"","
json = json & """xMun"": ""Curitiba"""
json = json & "},"
json = json & """ie"": ""1018146530"","
json = json & """crt"": 3,"
json = json & """xNome"": ""BEMATECH SA"","
json = json & """xFant"": ""BEMATECH"""
json = json & "},"
json = json & """destinatario"": {"
json = json & """cpf"": ""76643539129"","
json = json & """endereco"": {"
json = json & """nro"": ""842"","
json = json & """uf"": ""PR"","
json = json & """cep"": ""80020320"","
json = json & """fone"": ""41927598874"","
json = json & """xBairro"": ""Centro"","
json = json & """xLgr"": ""Marechal Deodoro"","
json = json & """cMun"": ""4106902"","
json = json & """cPais"": ""1058"","
json = json & """xPais"": ""Brasil"","
json = json & """xMun"": ""Curitiba"""
json = json & "},"
json = json & """indIEDest"": 9,"
json = json & """email"": ""teste@teste.com"","
json = json & """xNome"": ""NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL """
json = json & "},"
'json = json & "}"
'json = json & "{"
json = json & """produto"": {"
json = json & """cean"": ""7897238304177"","
json = json & """ncm"": ""85258029"","
json = json & """cfop"": ""5101"","
json = json & """indTot"": 1,"
json = json & """vUnCom"": 1.000,"
json = json & """uTrib"": ""UN"","
json = json & """vUnTrib"": ""1.000"","
json = json & """cProd"": ""85258029901234"","
json = json & """xProd"": ""Produto Teste"","
json = json & """uCom"": ""UN"","
json = json & """qTrib"": 1.000,"
json = json & """qCom"": ""1.000"","
json = json & """vProd"": 1.00"
json = json & "},"
json = json & """imposto"": {"
json = json & """icms"": {"
json = json & """icms00"": {"
json = json & """orig"": 1,"
json = json & """cst"": ""00"","
json = json & """modBC"": 3,"
json = json & """vbc"": 1.00,"
json = json & """picms"": 1.01,"
json = json & """vicms"": 0.01"
json = json & "}"
json = json & "},"
json = json & """vTotTrib"": 0.00"
json = json & "},"
'json = json & "}"
json = json & """tPag"": 1,"
json = json & """vPag"": 1.00,"
'json = json & "{"
json = json & """total"": {"
json = json & """icmsTotal"": {"
json = json & """vbc"": 1.00,"
json = json & """vicms"": 0.01,"
json = json & """vicmsDeson"": 0.00,"
json = json & """vbcst"": 0.00,"
json = json & """vst"": 0.00,"
json = json & """vii"": 0.00,"
json = json & """vipi"": 0.00,"
json = json & """vpis"": 0.00,"
json = json & """vcofins"": 0.00,"
json = json & """vnf"": 1.00,"
json = json & """vTotTrib"": 0.00,"
json = json & """vDesc"": 0.00,"
json = json & """vProd"": 1.00,"
json = json & """vOutro"": 0.00,"
json = json & """vSeg"": 0.00,"
json = json & """vFrete"": 0.00"
json = json & "}"
json = json & "},"
json = json & """informacaoAdicional"": {"
json = json & """infCpl"": "
json = json & """******************************************************************"
json = json & "                     Obrigado, volte sempre!                      "
json = json & "                Exemplo de uso com a BemaOne.dll                  "
json = json & "         Desenvolvido em VB6 - Bematech Software Partners         "
json = json & "******************************************************************"","
json = json & """observacoesContribuintes"": ["
json = json & "{"
json = json & """xTexto"": ""0.00"","
json = json & """xCampo"": ""Troco"""
json = json & "}"
json = json & "]"
json = json & "}"
json = json & "}"
TextJson.Text = json

sReturn = Bematech_Fiscal_FecharNota(json)
sFunctionReturn = GetStringFromPointer(sReturn)
TextRetorno.Text = sFunctionReturn
json = "{ "
    json = json & """id"": """ + Text2.Text + ""","
    json = json & """formato"": """ + Text3.Text + """ "
    json = json & "}"
  
    sReturn2 = Bematech_Fiscal_ConsultarNota(json)
    sFunctionReturn2 = GetStringFromPointer(sReturn2)
    TextRetorno.Text = sFunctionReturn + "*************************************************" + vbCrLf + "***************Status da Nota***************" + vbCrLf + sFunctionReturn2
End Sub

Private Sub ConsultaNota_Click()
Dim formato, json As String

If Option1.Value = True Then
    formato = "json"
End If
If Option2.Value = True Then
    formato = "pdf"
End If

json = "{ "
json = json & """id"": """ + Text4.Text + ""","
json = json & """formato"": """ + formato + """ "
json = json & "}"
  
    sReturn = Bematech_Fiscal_ConsultarNota(json)
    sFunctionReturn = GetStringFromPointer(sReturn)
    TextRetorno.Text = sFunctionReturn
End Sub

Private Sub Form_Load()

ComboComando.AddItem "AbrirNota"
ComboComando.AddItem "VendeItem"
ComboComando.AddItem "EfetuarPagamento"
ComboComando.AddItem "FecharNota"

dataHora = Format(Now, "yyyy-mm-dd" + "T" + "hh:mm:ss" + "-02:00")

End Sub

Private Sub ListarConfiguracoes_Click()
Dim sReturn As Long
Dim sFunctionReturn As String
sReturn = Bematech_Fiscal_ListarConfiguracoes()
sFunctionReturn = GetStringFromPointer(sReturn)
TextRetorno.Text = sFunctionReturn

End Sub

Private Sub ObterInfoSistema_Click()
Dim sReturn As Long
Dim sFunctionReturn As String
sReturn = Bematech_Fiscal_ObterInformacoesSistema()
sFunctionReturn = GetStringFromPointer(sReturn)
TextRetorno.Text = sFunctionReturn
End Sub

Private Sub ObterStatusImpressora_Click()
Dim sReturn As Long
Dim sFunctionReturn As String
sReturn = Bematech_Fiscal_ObterStatusImpressora()
sFunctionReturn = GetStringFromPointer(sReturn)
TextRetorno.Text = sFunctionReturn
End Sub

Private Sub Sair_Click()
Unload Me
End Sub

Private Sub TextoLivre_Click()
Dim sReturn As Long
Dim sFunctionReturn As String
Dim json, cut, jsonCut As String
cut = Base64EncodeString("TextSaida.Text+Chr(27) + m")
'cut = "VGV4dG8gYSBzZXIgaW1wcmVzc28rQ2hyKDI3KStDaHIoMTE5KQ=="

If CheckCut.Value = True Then
    json = "{""dados"": """ + cut + """, ""base64"": true}"
    sReturn = Bematech_Fiscal_ImprimirTextoLivre(json)
    
    sFunctionReturn = GetStringFromPointer(sReturn)
    TextRetorno.Text = sFunctionReturn
   
End If

If CheckCut.Value = False Then
    json = "{""dados"": """ + TextSaida.Text + """, ""base64"": false}"
    sReturn = Bematech_Fiscal_ImprimirTextoLivre(json)
    sFunctionReturn = GetStringFromPointer(sReturn)
    TextRetorno.Text = sFunctionReturn
End If

End Sub
