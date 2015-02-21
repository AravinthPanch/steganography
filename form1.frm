VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   8184
   ClientLeft      =   60
   ClientTop       =   756
   ClientWidth     =   10032
   FillColor       =   &H00E0E0E0&
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   8.4
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8184
   ScaleWidth      =   10032
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   7200
      Width           =   855
   End
   Begin VB.CheckBox chkShowPixels 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show Pixels"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.4
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   855
   End
   Begin MSComDlg.CommonDialog dlgImage 
      Left            =   240
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Grpahic Files|*.bmp;*.gif;*.jpg|All Files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdDecode 
      Caption         =   "Decode"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "Encode"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   7335
   End
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4128
      Left            =   1200
      ScaleHeight     =   4080
      ScaleWidth      =   6660
      TabIndex        =   0
      Top             =   2640
      Width           =   6708
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Password  "
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Message"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
      Begin VB.Menu mnuCloseapplication 
         Caption         =   "&CloseApplication"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ArrangeControls()
Dim wid As Single

    Width = picImage.Left + picImage.Width + Width - ScaleWidth + 120
    Height = picImage.Top + picImage.Height + Height - ScaleHeight + 120
    wid = ScaleWidth - txtMessage.Left - 120
    If wid < 120 Then wid = 120
    txtMessage.Width = wid
    txtPassword.Width = wid
End Sub

'For exit command in form1
Private Sub cmdExit_Click()
Unload Form1
Unload Form2
End
End Sub
'routine to be loaded when the form is being loaded.
Private Sub Form_Load()
    picImage.ScaleMode = vbPixels
    picImage.AutoRedraw = True
    dlgImage.InitDir = App.Path
    ArrangeControls
End Sub
'for closeapplication in 'exit' file menu
Private Sub mnuCloseapplication_Click()
Unload Form1
Form1.Visible = False
Unload Form2
Form2.Visible = False
End
End Sub

'for 'Open' in file menu
Private Sub mnuFileOpen_Click()
    On Error Resume Next
    dlgImage.CancelError = True
    dlgImage.Flags = _
        cdlOFNFileMustExist Or _
        cdlOFNHideReadOnly Or _
        cdlOFNLongNames
    dlgImage.ShowOpen
    If Err.Number <> 0 Then Exit Sub

    picImage.Picture = LoadPicture(dlgImage.FileName)
    ArrangeControls
    If Err.Number <> 0 Then Exit Sub

    dlgImage.InitDir = dlgImage.FileName
    dlgImage.FileName = dlgImage.FileTitle
End Sub
'for SaveAs in file menu
Private Sub mnuFileSaveAs_Click()
    On Error Resume Next
    dlgImage.CancelError = True
    dlgImage.Flags = _
        cdlOFNOverwritePrompt Or _
        cdlOFNHideReadOnly Or _
        cdlOFNLongNames
    dlgImage.ShowSave
    If Err.Number <> 0 Then Exit Sub

    SavePicture picImage.Picture, dlgImage.FileName
    If Err.Number <> 0 Then Exit Sub

    dlgImage.InitDir = dlgImage.FileName
    dlgImage.FileName = dlgImage.FileTitle
End Sub

' Pick an unused (r, c, pixel) combination.
Private Sub PickPosition(ByVal used_positions As Collection, ByVal wid As Integer, ByVal hgt As Integer, ByRef r As Integer, ByRef c As Integer, ByRef pixel As Integer)
Dim position_code As String

    On Error Resume Next
    Do
        ' Pick a position.
        r = Int(Rnd * wid)
        c = Int(Rnd * hgt)
        pixel = Int(Rnd * 3)

        ' See if the position is unused.
        position_code = "(" & r & "," & c & "," & pixel & ")"
        used_positions.Add position_code, position_code
        If Err.Number = 0 Then Exit Do
        Err.Clear
    Loop
End Sub
' Return the color's components.
Private Sub UnRGB(ByVal color As OLE_COLOR, ByRef r As Byte, ByRef g As Byte, ByRef b As Byte)
    'this is done in order to explore the pixel's values.
    r = color And &HFF&
    'ff in hexadecimal is equal to 255 in decimal
    'ff is ANDed with the color component to get its value
    g = (color And &HFF00&) \ &H100&
    'the second byte of the pixel  has the value of green
    'so and it with ffoo;you will get a 4 digit number;
    'but we need only the right most 2 digits
    'so divide it by 100;
    'all numbers correpond to hexadecimal value;
    b = (color And &HFF0000) \ &H10000
    'the same explanation as given for the green color.
End Sub

' Translate a password into an offset value.
Private Function NumericPassword(ByVal password As String) As Long
'in this function we translate the string password,
'into a LONG numeric value.In VB the numbers are cosidered
'as string and are processed, if not specified explicitly.
Dim Value As Long
'value is the variable where we store the converted string value
Dim ch As Long
Dim shift1 As Long
Dim shift2 As Long
Dim i As Integer
'str_len variable tells the string's length.
Dim str_len As Integer

    ' Initialize the shift values to different
    ' non-zero values.
    shift1 = 3
    shift2 = 17

    'I. Process the message.
    'I.a. first find the length using Len function
    Debug.Print Value
    str_len = Len(password)
    For i = 1 To str_len
        ' II.Add the next letter.
        'Mid$ function is used to extract the letters from
        'the position i to 1. ie.,i and its next position
        'and then convert it to ascii value
        ch = Asc(Mid$(password, i, 1))
        Value = Value Xor (ch * 2 ^ shift1)
        Value = Value Xor (ch * 2 ^ shift2)

        ' Change the shift offsets.
        'changing the shifts offset inorder that it does'nt
        'collide with already existing values.
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
        'the numbers used in value ,shift1 and shift2
        'calculations are randomly choosen numbers with any
        'order.
    Next i
    NumericPassword = Value
    Debug.Print password
    Debug.Print NumericPassword
End Function
'For encode command in form1
Private Sub cmdEncode_Click()
Dim msg As String
Dim i As Integer
' a collection is nothing but  similar to an array where the data
' related to an entity can be accessed easily by specifying their names
Dim used_positions As Collection
'width abbrevated as wid
Dim wid As Integer
'height abbreated for height
Dim hgt As Integer
Dim show_pixels As Boolean

    'activeX Methods(study it latter!just use it now.)
    Screen.MousePointer = vbHourglass
    DoEvents

    ' Initialize the random number generator.
    Rnd -1
    'convert the typed password to a number if the passwd is a text
    ' so pass the password to a function which gives a numeric value
    Randomize NumericPassword(txtPassword.Text)

    wid = picImage.ScaleWidth
    hgt = picImage.ScaleHeight
    'restrict the message length upto 255 characters from the left
    msg = Left$(txtMessage.Text, 255)
    'checking chkshowpixels box to find whether to show the encoded pixels
    show_pixels = chkShowPixels.Value
    'define a new collection to maintain a record of used pixels
    Set used_positions = New Collection

    ' Encode the message length.
    EncodeByte CByte(Len(msg)), used_positions, wid, hgt, show_pixels

    ' Encode the message.
    For i = 1 To Len(msg)
        EncodeByte Asc(Mid$(msg, i, 1)), used_positions, wid, hgt, show_pixels
    Next i
    picImage.Picture = picImage.Image

    Screen.MousePointer = vbDefault
End Sub


' Encode this byte's data.
Private Sub EncodeByte(ByVal Value As Byte, ByVal used_positions As Collection, ByVal wid As Integer, ByVal hgt As Integer, ByVal show_pixels As Boolean)
Dim i As Integer
Dim byte_mask As Integer
Dim r As Integer
Dim c As Integer
Dim pixel As Integer
Dim clrr As Byte
Dim clrg As Byte
Dim clrb As Byte
Dim color_mask As Integer

    byte_mask = 1
    For i = 1 To 8
        ' Pick a random pixel and RGB component.
        PickPosition used_positions, wid, hgt, r, c, pixel

        ' Get the pixel's color components.
        'the point method extracts the values of the
        '3 bytes of a pixel in row r and column c
        'to be clear, the below function is
        'UnRGB (picImage.point(r,c)),clrr,clrg,clrb
        'the component picImage.point(r,c) gives the pixels' values
        UnRGB picImage.Point(r, c), clrr, clrg, clrb
        If show_pixels Then
            clrr = 255
            'we represent the encoded pixels in red color
            'so hide the green and blue color component
            'that is make their values zero
            'this is done by ANDing with the hex H1
            clrg = clrg And &H1
            'green color component made 0 by the above line
            clrb = clrb And &H1
            'blue color component made 0 by the above line
        End If

        ' Get the value we must store.
        If Value And byte_mask Then
            color_mask = 1
        Else
            color_mask = 0
        End If

        ' Update the color.
        'FE in hexadecimal is equal to 254in decimal
        'thereby we  are seperating the least significant bit
        Select Case pixel
            Case 0
                clrr = (clrr And &HFE) Or color_mask
            Case 1
                clrg = (clrg And &HFE) Or color_mask
            Case 2
                clrb = (clrb And &HFE) Or color_mask
        End Select

        ' Set the pixel's color.
        picImage.PSet (r, c), RGB(clrr, clrg, clrb)

        byte_mask = byte_mask * 2
    Next i
End Sub
'For decode command in form1
Private Sub cmdDecode_Click()
Dim msg_length As Byte
Dim msg As String
Dim ch As Byte
Dim i As Integer
Dim used_positions As Collection
Dim wid As Integer
Dim hgt As Integer
Dim show_pixels As Boolean

    Screen.MousePointer = vbHourglass
    DoEvents

    ' Initialize the random number generator.
    Rnd -1
    Randomize NumericPassword(txtPassword.Text)

    wid = picImage.ScaleWidth
    hgt = picImage.ScaleHeight
    show_pixels = chkShowPixels.Value
    Set used_positions = New Collection

    ' Decode the message length.
    msg_length = DecodeByte(used_positions, wid, hgt, show_pixels)

    ' Decode the message.
    For i = 1 To msg_length
        ch = DecodeByte(used_positions, wid, hgt, show_pixels)
        msg = msg & Chr$(ch)
    Next i
    picImage.Picture = picImage.Image

    txtMessage.Text = msg

    Screen.MousePointer = vbDefault
End Sub

' Decode this byte's data.
Private Function DecodeByte(ByVal used_positions As Collection, ByVal wid As Integer, ByVal hgt As Integer, ByVal show_pixels As Boolean) As Byte
Dim Value As Integer
Dim i As Integer
Dim byte_mask As Integer
Dim r As Integer
Dim c As Integer
Dim pixel As Integer
Dim clrr As Byte
Dim clrg As Byte
Dim clrb As Byte
Dim color_mask As Integer

    byte_mask = 1
    For i = 1 To 8
        ' Pick a random pixel and RGB component.
        PickPosition used_positions, wid, hgt, r, c, pixel

        ' Get the pixel's color components.
        'point method is simpler and used in image application
        'point method reads the pixel value directly from the
        'picture box's picture.point's two arguments are the
        'row and column position of the image respectively.
        'UnRGB is my function. i have written it below.
        UnRGB picImage.Point(r, c), clrr, clrg, clrb

        ' Get the stored value.
        Select Case pixel
            Case 0
                color_mask = (clrr And &H1)
            Case 1
                color_mask = (clrg And &H1)
            Case 2
                color_mask = (clrb And &H1)
        End Select

        If color_mask Then
            Value = Value Or byte_mask
        End If

        If show_pixels Then
            picImage.PSet (r, c), RGB(clrr And &H1, clrg And &H1, clrb And &H1)
        End If

        byte_mask = byte_mask * 2
    Next i

    DecodeByte = CByte(Value)
End Function


