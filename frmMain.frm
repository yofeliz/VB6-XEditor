VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "X-Editor"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LockControls    =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgDialogo 
      Left            =   720
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "X-Editor"
      Filter          =   "Archivos de texto plano (*.txt)|*.txt| Todos los archivos (*.*)|*.*"
      InitDir         =   "./"
      MaxFileSize     =   32000
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   120
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "Salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":059C
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":06AE
            Key             =   "Rehacer"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":07C0
            Key             =   "Deshacer"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08D2
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09E4
            Key             =   "Ajustar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AF6
            Key             =   "Copiar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C08
            Key             =   "Fuente"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D1A
            Key             =   "Pegar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E2C
            Key             =   "Cortar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F3E
            Key             =   "Eliminar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1050
            Key             =   "Comprobar"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1162
            Key             =   "Abrir"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1274
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1386
            Key             =   "Ayuda"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1498
            Key             =   "Guardar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbBarra 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "1"
            TextSave        =   "1"
            Key             =   "Fila"
            Object.ToolTipText     =   "Fila"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "1"
            TextSave        =   "1"
            Key             =   "Columna"
            Object.ToolTipText     =   "Columna"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "1"
            TextSave        =   "1"
            Key             =   "Caracteres"
            Object.ToolTipText     =   "Caracteres"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Key             =   "Modificado"
            Object.ToolTipText     =   "Modificado"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "14:04"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbEditor 
      Height          =   6615
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11668
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   1e7
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":15AA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbBarra 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgLista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageKey        =   "Nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Abrir"
            Object.ToolTipText     =   "Abrir"
            ImageKey        =   "Abrir"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Guardar"
            Object.ToolTipText     =   "Guardar"
            ImageKey        =   "Guardar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageKey        =   "Imprimir"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Deshacer"
            Object.ToolTipText     =   "Deshacer"
            ImageKey        =   "Deshacer"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copiar"
            Object.ToolTipText     =   "Copiar"
            ImageKey        =   "Copiar"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cortar"
            Object.ToolTipText     =   "Cortar"
            ImageKey        =   "Cortar"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pegar"
            Object.ToolTipText     =   "Pegar"
            ImageKey        =   "Pegar"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageKey        =   "Eliminar"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ajustar"
            Object.ToolTipText     =   "Ajuste de Línea"
            Object.Tag             =   "0"
            ImageKey        =   "Ajustar"
            Style           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Fuente"
            Object.ToolTipText     =   "Fuente"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageKey        =   "Salir"
         EndProperty
      EndProperty
   End
   Begin VB.Menu menArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu menNuevo 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu menAbrir 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu menGuardar 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu menImprimir 
         Caption         =   "Im&pimir"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu menSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu menEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu menCopiar 
         Caption         =   "&Copiar"
      End
      Begin VB.Menu menCortar 
         Caption         =   "Co&rtar"
      End
      Begin VB.Menu menPegar 
         Caption         =   "&Pegar"
      End
      Begin VB.Menu menEliminar 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu menDeshacer 
         Caption         =   "&Deshacer"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu menSeleccionar 
         Caption         =   "&Seleccionar todo"
      End
   End
   Begin VB.Menu menFormato 
      Caption         =   "&Formato"
      Begin VB.Menu menAjuste 
         Caption         =   "&Ajuste de línea"
      End
      Begin VB.Menu menFuente 
         Caption         =   "&Fuente"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu menAcerca 
         Caption         =   "&Acerca de..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intervaloCursor As Long

Private Sub salir()
    Form_QueryUnload 0, 0
End Sub

Private Sub nuevo()
    On Error Resume Next
    If stbBarra.Panels("Modificado").Text = "*" Then
        If MsgBox("El documento ha sido modificado." + vbCrLf + _
                  "¿Desea guardar los cambios?", vbExclamation + vbYesNo, "Guardar cambios") = vbYes Then
               guardar
        End If
    End If
    
    rtbEditor.Text = ""
    stbBarra.Panels("Modificado").Text = ""
End Sub

Private Sub abrir()
    On Error GoTo errorSeleccion
    
    Dim nombreFichero As String
    
    If stbBarra.Panels("Modificado").Text = "*" Then
        If MsgBox("El documento ha sido modificado." + vbCrLf + _
                  "¿Desea guardar los cambios?", vbExclamation + vbYesNo, "Guardar cambios") = vbYes Then
               guardar
        End If
    End If
    
    dlgDialogo.DialogTitle = "Abrir..."
    dlgDialogo.ShowOpen
    nombreFichero = dlgDialogo.FileName
   
    On Error GoTo errorCarga
    rtbEditor.LoadFile nombreFichero, rtfText
    Me.Caption = App.EXEName + "   [ " + nombreFichero + " ]"
    stbBarra.Panels("Modificado").Text = ""

    Exit Sub

errorSeleccion:
    Exit Sub
errorCarga:
    MsgBox "El fichero seleccionado no se ha podido cargar.", vbCritical, "Error"
    Exit Sub
End Sub

Private Sub guardar()
    On Error GoTo errorSeleccion
    
    Dim nombreFichero As String
    
    dlgDialogo.DialogTitle = "Guardar..."
    dlgDialogo.ShowSave
    nombreFichero = dlgDialogo.FileName
    
    On Error GoTo errorGuarda
    If LCase(Right(nombreFichero, 4)) <> ".txt" Then
        nombreFichero = nombreFichero + ".txt"
    End If

    rtbEditor.SaveFile nombreFichero, rtfText
    Me.Caption = App.EXEName + "   [ " + nombreFichero + " ]"
    stbBarra.Panels("Modificado").Text = ""
    Exit Sub

errorSeleccion:
    Exit Sub
errorGuarda:
    MsgBox "El documento no se ha podido guardar.", vbCritical, "Error"
    Exit Sub
End Sub

Private Sub copiar()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText rtbEditor.SelText
End Sub

Private Sub pegar()
    On Error Resume Next
    rtbEditor.SelText = Clipboard.GetText
End Sub

Private Sub cortar()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText rtbEditor.SelText
    rtbEditor.SelText = ""
End Sub

Private Sub eliminar()
    rtbEditor.SelText = ""
End Sub

Private Sub ajustar()
    On Error Resume Next
    If tlbBarra.Buttons("Ajustar").Tag = 0 Then
        rtbEditor.RightMargin = 0
        tlbBarra.Buttons("Ajustar").Tag = 1
        tlbBarra.Buttons("Ajustar").Value = tbrPressed
    Else
        rtbEditor.RightMargin = 9999999
        tlbBarra.Buttons("Ajustar").Tag = 0
        tlbBarra.Buttons("Ajustar").Value = tbrUnpressed
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    intervaloCursor = GetCaretBlinkTime
    Call SetCaretBlinkTime(2147483647)
End Sub

Private Sub menAbrir_Click()
    abrir
End Sub

Private Sub menAcerca_Click()
    frmAcerca.Show vbModal
End Sub

Private Sub menAjuste_Click()
    ajustar
End Sub

Private Sub menCopiar_Click()
    copiar
End Sub

Private Sub menCortar_Click()
    cortar
End Sub

Private Sub menDeshacer_Click()
    deshacer
End Sub

Private Sub menEliminar_Click()
    eliminar
End Sub

Private Sub menFuente_Click()
    fuente
End Sub

Private Sub menGuardar_Click()
    guardar
End Sub

Private Sub menNuevo_Click()
    nuevo
End Sub

Private Sub menPegar_Click()
    pegar
End Sub

Private Sub menSalir_Click()
    Form_QueryUnload 0, 0
End Sub

Private Sub menSeleccionar_Click()
    On Error Resume Next
    rtbEditor.SelStart = 0
    rtbEditor.SelLength = Len(rtbEditor.Text)
End Sub

Private Sub fuente()
    On Error Resume Next
    
    dlgDialogo.Flags = cdlCFScreenFonts
    With dlgDialogo
        .FontName = rtbEditor.Font.Name
        .FontSize = rtbEditor.Font.Size
        .FontBold = rtbEditor.Font.Bold
        .FontItalic = rtbEditor.Font.Italic
    End With
    dlgDialogo.ShowFont
    With rtbEditor
        .Font = dlgDialogo.FontName
        .Font.Size = dlgDialogo.FontSize
        .Font.Bold = dlgDialogo.FontBold
        .Font.Italic = dlgDialogo.FontItalic
    End With
End Sub
   
Private Sub deshacer()
    On Error Resume Next
    Call SendMessage(rtbEditor.hwnd, EM_UNDO, 0&, 0&)
End Sub

Private Sub rtbEditor_Change()
    stbBarra.Panels("Modificado").Text = "*"
    refrescarDatos
End Sub

Private Sub refrescarDatos()
    On Error Resume Next
    
    Dim fila As Variant
    Dim columna As Variant
    Dim caracteres As Variant
    Dim temp2 As Variant

    fila = SendMessage(rtbEditor.hwnd, EM_LINEFROMCHAR, rtbEditor.SelStart, 0&)
    caracteres = rtbEditor.SelStart
    temp2 = CLng(SendMessage(rtbEditor.hwnd, EM_LINEINDEX, fila, 0&))
    columna = (CLng(rtbEditor.SelStart) - temp2) + 1
    stbBarra.Panels("Fila").Text = Format(CStr(fila + 1), "#")
    stbBarra.Panels("Columna").Text = Format(Str(columna), "#")
    stbBarra.Panels("Caracteres").Text = Format(CStr(caracteres + 1), "#")
    Call CreateCaret(rtbEditor.hwnd, 0, rtbEditor.Font.Size - 1, rtbEditor.Font.Size + 5)
    Call ShowCaret(rtbEditor.hwnd)
End Sub

Private Sub rtbEditor_Click()
    refrescarDatos
End Sub

Private Sub rtbEditor_DblClick()
    refrescarDatos
End Sub

Private Sub rtbEditor_GotFocus()
    refrescarDatos
End Sub

Private Sub rtbEditor_KeyDown(KeyCode As Integer, Shift As Integer)
    refrescarDatos
End Sub

Private Sub rtbEditor_KeyPress(KeyAscii As Integer)
    refrescarDatos
End Sub

Private Sub rtbEditor_KeyUp(KeyCode As Integer, Shift As Integer)
    refrescarDatos
End Sub

Private Sub rtbEditor_LostFocus()
    refrescarDatos
End Sub

Private Sub rtbEditor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    refrescarDatos
End Sub

Private Sub rtbEditor_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    refrescarDatos
End Sub

Private Sub rtbEditor_SelChange()
    refrescarDatos
End Sub

Private Sub rtbEditor_Validate(Cancel As Boolean)
    refrescarDatos
End Sub

Private Sub tlbBarra_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case Is = "Nuevo"
            nuevo
        Case Is = "Abrir"
            abrir
        Case Is = "Guardar"
            guardar
        Case Is = "Imprimir"
            
        Case Is = "Copiar"
            copiar
        Case Is = "Pegar"
            pegar
        Case Is = "Cortar"
            cortar
        Case Is = "Eliminar"
            eliminar
        Case Is = "Deshacer"
            deshacer
        Case Is = "Ajustar"
            ajustar
        Case Is = "Fuente"
            fuente
        Case Is = "Salir"
            salir
   End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If stbBarra.Panels("Modificado").Text = "*" Then
        If MsgBox("El documento ha sido modificado." + vbCrLf + _
                  "¿Desea guardar los cambios?", vbExclamation + vbYesNo, "Guardar cambios") = vbYes Then
            guardar
        End If
    End If
    Call SetCaretBlinkTime(intervaloCursor)
    End
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rtbEditor.Width = frmMain.Width - 120
    rtbEditor.Height = frmMain.Height - 1375
    rtbEditor.Top = tlbBarra.Height
End Sub
