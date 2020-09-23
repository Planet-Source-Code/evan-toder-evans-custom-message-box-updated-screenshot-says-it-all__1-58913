VERSION 5.00
Begin VB.UserControl GlassBox 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   315
   InvisibleAtRuntime=   -1  'True
   Picture         =   "GlassBox.ctx":0000
   ScaleHeight     =   330
   ScaleWidth      =   315
   ToolboxBitmap   =   "GlassBox.ctx":0342
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1800
      Top             =   1170
   End
   Begin VB.PictureBox picBoxPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   1980
      ScaleHeight     =   870
      ScaleWidth      =   1005
      TabIndex        =   1
      Top             =   135
      Width           =   1005
   End
   Begin VB.PictureBox picGlass 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   810
      Picture         =   "GlassBox.ctx":0654
      ScaleHeight     =   1725
      ScaleWidth      =   870
      TabIndex        =   0
      Top             =   -45
      Width           =   870
   End
End
Attribute VB_Name = "GlassBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ
'  I HAVE TO GIVE MUCH THANKS TO IRBme FOR HIS RECENT ARTICLE ON
'  EFFECTIVE CODE POSTED ON FEB 12 2005.
'  EVEN THOUGH I KNEW ABOUT ALL THE THINGS HE SUGGESTED IN THE ARTICLE
'  READING HIS ARTICLE HAD THE "LIGHTBULB GOING OFF IN THE HEAD" EFFECT
'  AND THIS SUBMISSION IS CODED WITH THE STANDARDS HIS ARTICLE SET FORTH
'  THANKS THERE IRBme (IF YOUR READING THIS :-)
'ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ

'API DECLARATIONS
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Enum enButtons
    None = 0
    OK = 1
    YesNo = 2
End Enum

Enum enGlassBoxColor
    GlassRed = 0
    GlassCandy = 1
    GlassOrange = 2
    GlassBlue = 3
    GlassWhite = 4
    GlassYellow = 5
End Enum

'DIFFERENT CONTROLS ON THE MESSAGE BOX
Private WithEvents lblMsg      As Label
Attribute lblMsg.VB_VarHelpID = -1
Private WithEvents lblCaption  As Label
Attribute lblCaption.VB_VarHelpID = -1
Private WithEvents lblOk       As Label
Attribute lblOk.VB_VarHelpID = -1
Private WithEvents lblNo       As Label
Attribute lblNo.VB_VarHelpID = -1
Private WithEvents lblYes      As Label
Attribute lblYes.VB_VarHelpID = -1
Private WithEvents F           As Form
Attribute F.VB_VarHelpID = -1


Dim m_messageReturnVal         As Long
Dim m_lineColor                As Long
Dim m_messageColor             As Long

'Default Property Values:
Const m_def_PictureMaskColor = &HFF00FF
Const m_def_GlassBoxColor = 0

'Property Variables:
Dim m_PictureMaskColor         As OLE_COLOR
Dim m_GlassBoxColor            As enGlassBoxColor



Function Message(strMessage As String, _
                 Optional strCaption As String, _
                 Optional Buttons As enButtons, _
                 Optional AutoHideSeconds As Long) As Long
         
        m_messageReturnVal = 0
        '  set the properties of the labels that
        '  make up the caption and the message
        Call SetBoxFonts(strMessage, strCaption)
        '  position all of frmMessage controls
        Call PositionFormsControls
        '  center the form horizontally and vertically
        '  in relation to the form holding this control
        F.Move PositionMeLeft(UserControl.Parent), _
               PositionMeTop(UserControl.Parent)
        '  is this going to be a self closing messagebox?
        Call SetSelfCloseOptions(Buttons, AutoHideSeconds)
        Call ShowCorrectButtons(Buttons)
        Call Paint
        '  show our message box
        F.Show DeterminedModality(Buttons), UserControl.Parent
        '  the return val if its a yes no message box
        If Buttons = YesNo Then Message = m_messageReturnVal

End Function
 
Private Sub SetBoxFonts(strMsg As String, strTitle As String)

  With frmMessage
        '  the font specifics for the message
       .lblMsg.Font.Size = 10
       .lblMsg.AutoSize = True
       .lblMsg.Alignment = 0 'left justify
       .lblMsg = MinMsgSizeString(strMsg)
        '  the font specifics for the caption
       .lblCaption.Font.Size = 8
       .lblCaption.Font.Bold = True
       .lblCaption.Alignment = 2 'centered
       .lblCaption = strTitle
   End With
   
End Sub

Private Function MinMsgSizeString(strMessage As String) As String
'
'the reason for this sub is that the label that holds
'the message has its [autosize]=True. The form then
'resizes to "wrap" around the label this allowing
'the form size to readjust to the message, just like
'regular message box.
'If the user supplies a [strMessage] of "", then things
'wont exactly look quite right
'
 Dim msgLen   As Long, lenDif   As Long
 
 msgLen = Len(strMessage)
 If msgLen < 40 Then
    lenDif = (40 - msgLen)
    MinMsgSizeString = strMessage & String(lenDif, " ")
 Else
    MinMsgSizeString = strMessage
 End If

End Function
Private Sub PositionFormsControls()
  '
  'we need to position the controls so that
  'this mbox form matches the size of lblMessage
  'and that the buttons are in the right place
  With lblMsg
     .Move 100, 300
      '  readjusts the forms dimensions to fit the message
      F.Width = .Width + 150
      F.Height = .Height + 800
  End With
  
  '  position caption label
  lblCaption.Move 0, 20, F.Width, 230
  '  position OK button
  With lblOk
     .Move (F.Width * 0.5) - (.Width * 0.5), (F.Height - 300)
  End With
  '  position Yes button
  With lblYes
     .Move (F.Width * 0.5) - (.Width + 50), (F.Height - 300)
  End With
  '  position No button
  With lblNo
     .Move (F.Width * 0.5) + 50, (F.Height - 300)
  End With
 
End Sub

Private Function PositionMeLeft(callingForm As Object) As Long

  Dim halfOfMe As Long, halfOfYou As Long, leftPoint As Long

  With callingForm
     '  store coods for positioning this left
     halfOfMe = (F.Width * 0.5)
     halfOfYou = UserControl.Parent.Left + (UserControl.Parent.Width * 0.5)
     PositionMeLeft = (halfOfYou - halfOfMe)
  End With

End Function

Private Function PositionMeTop(callingForm As Object) As Long

  Dim halfOfMe As Long, halfOfYou As Long, toppoint As Long

  With callingForm
     '  store coods for positioning this top
     halfOfMe = (F.Height * 0.5)
     halfOfYou = UserControl.Parent.Top + (UserControl.Parent.Height * 0.5)
     PositionMeTop = (halfOfYou - halfOfMe)
  End With

End Function

Private Sub SetSelfCloseOptions(Buttons As enButtons, AutoHideSeconds As Long)
 
   '  are we autoclosing this ?
   '  autoclose is NOT a valid option
   '  if the messagebox is a yesNo messagebox
   If Buttons <> YesNo Then
      If Buttons = None Then
         '  however, if there are not buttons,
         '  this HAS to be self closing
         If AutoHideSeconds < 1 Then
           Timer1.Interval = 3000
         Else
           Timer1.Interval = (AutoHideSeconds * 1000)
         End If
      Else
        '  user chose OK button
        If AutoHideSeconds > 0 Then
           Timer1.Interval = (AutoHideSeconds * 1000)
        End If
      End If

      Timer1.Enabled = True
   End If

End Sub

Private Function DeterminedModality(Buttons As enButtons) As Long
  '
  'this function will determine when and when not
  'to make this message box modal or not, obviously
  'with no buttons it cant be modal (and better be
  'self closing)
  DeterminedModality = 1

  If Buttons = None Then
      DeterminedModality = 0
  End If

End Function

Private Sub ShowCorrectButtons(Buttons As enButtons)
  '
  '  show and position specified buttons
  '
  'set all hidden as default
  lblOk.Visible = False
  lblYes.Visible = False
  lblNo.Visible = False
  '
  If Buttons = None Then

  ElseIf Buttons = OK Then
      lblOk.Visible = True
  ElseIf Buttons = YesNo Then
      lblYes.Visible = True
      lblNo.Visible = True
  End If

End Sub

Private Sub DetermineSubtleLineColor()
 '
 'THIS SUB SELECTS THE RIGHT COLOR FOR THE SUBTLE LINE
 'THAT RUN HORIZONATALY ACROSS THE BOX AND THE CAPTION COLO
 '
 '
 '  set an assumed forcolor for the titlebar because
 '  in all but two of these colors, the titlebar caption
 '  will be white
 lblCaption.ForeColor = vbWhite
 
 If m_GlassBoxColor = GlassBlue Then
     m_lineColor = RGB(240, 240, 255)
     m_messageColor = RGB(50, 180, 230)
     
 ElseIf m_GlassBoxColor = GlassCandy Then
     m_lineColor = RGB(255, 235, 235)
     m_messageColor = RGB(225, 140, 160)
     
 ElseIf m_GlassBoxColor = GlassOrange Then
     m_lineColor = RGB(255, 240, 230)
     m_messageColor = RGB(255, 150, 50)
     
 ElseIf m_GlassBoxColor = GlassRed Then
     m_lineColor = RGB(255, 240, 240)
     m_messageColor = vbRed
     
 ElseIf m_GlassBoxColor = GlassWhite Then
     m_lineColor = RGB(240, 240, 245)
     m_messageColor = vbBlack
     lblCaption.ForeColor = RGB(80, 80, 100)
     
 ElseIf m_GlassBoxColor = GlassYellow Then
     m_lineColor = RGB(255, 255, 220)
     m_messageColor = RGB(175, 175, 0)
     lblCaption.ForeColor = RGB(145, 145, 0)
     
 End If
  
 '  forecolor for all the labels (buttons/captions)
 lblYes.ForeColor = m_messageColor
 lblNo.ForeColor = m_messageColor
 lblOk.ForeColor = m_messageColor
 lblMsg.ForeColor = m_messageColor
 
End Sub
 
 
Private Sub Paint()
   
   '  paint the titlebar
   Dim boxPixWid  As Long
   boxPixWid = (F.Width / Screen.TwipsPerPixelX)
   Dim srcPicY    As Long
   srcPicY = (m_GlassBoxColor * 19)
   StretchBlt F.hdc, 0, 0, boxPixWid, 20, _
              picGlass.hdc, 0, srcPicY, 30, 20, SRCCOPY
   
   'paint the lines
   Dim lcnt As Long
   For lcnt = 330 To F.Height Step 30
      F.Line (50, lcnt)-(F.Width - 50, lcnt), m_lineColor
   Next lcnt
   
   '  dark line surrounding message part of box
   F.Line (0, 300)-(F.Width - 20, F.Height - 20), RGB(160, 160, 175), B
   
   
   With picBoxPicture
      If .Picture = 0 Then Exit Sub
       
      Dim pixwid  As Long, pixhei As Long
      Dim toppoint As Long, titlebarHeight As Long

      '  get api friendly measurements of the source pic
      pixwid = (.Width / Screen.TwipsPerPixelX)
      pixhei = (.Height / Screen.TwipsPerPixelY)
      titlebarHeight = (290 / Screen.TwipsPerPixelX)
      toppoint = _
          ((F.Height * 0.5) / Screen.TwipsPerPixelX) - _
          ((.Height * 0.5) / Screen.TwipsPerPixelY)
      ' we dont want the top position of the picture to
      '  be any higher up than the bottom of our titlebar
      If toppoint <= titlebarHeight Then
         toppoint = titlebarHeight
      End If
      '  paint the pic
      TransparentBlt F.hdc, 5, toppoint, pixwid, pixhei, _
         .hdc, 0, 0, pixwid, pixhei, m_PictureMaskColor
  End With
  
End Sub

'MOVE THE MESSAGE BOX AROUND BY DRAGGING ON LBLMSG OR LBLCAPTION
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call MoveItem(F.hwnd)
End Sub
Private Sub lblMsg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call MoveItem(F.hwnd)
End Sub

'RETURN VAL FROM THE YES OR NO BUTTON BEING CLICKED
Private Sub lblNo_Click()
   m_messageReturnVal = vbNo
   F.Visible = False
End Sub
Private Sub lblYes_Click()
   m_messageReturnVal = vbYes
   F.Visible = False
End Sub
Private Sub lblOk_Click()
   F.Visible = False
End Sub


Private Sub Timer1_Timer()
 '
 'the purpose of this timer is to autoclose the messagebox
 F.Visible = False
 Timer1.Interval = 0
 
End Sub

Private Sub UserControl_Resize()
    Size (16 * Screen.TwipsPerPixelX), (16 * Screen.TwipsPerPixelY)
End Sub

Private Sub UserControl_Terminate()
   On Error Resume Next
   Unload F
End Sub
'GlassBoxColor
Public Property Get GlassBoxColor() As enGlassBoxColor
Attribute GlassBoxColor.VB_Description = "The color for the message boxes titlebar, message text, and its subtle horizontal lines"
    GlassBoxColor = m_GlassBoxColor
End Property
Public Property Let GlassBoxColor(ByVal New_GlassBoxColor As enGlassBoxColor)
    m_GlassBoxColor = New_GlassBoxColor
    PropertyChanged "GlassBoxColor"
    
    Call DetermineSubtleLineColor
    Call Paint
End Property
'messagebox picture
Public Property Get Picture() As Picture
    Set Picture = picBoxPicture.Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
    Set picBoxPicture.Picture = New_Picture
    PropertyChanged "Picture"
End Property
'PictureMaskColor
Public Property Get PictureMaskColor() As OLE_COLOR
    PictureMaskColor = m_PictureMaskColor
End Property
Public Property Let PictureMaskColor(ByVal New_PictureMaskColor As OLE_COLOR)
    m_PictureMaskColor = New_PictureMaskColor
    PropertyChanged "PictureMaskColor"
End Property


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_GlassBoxColor = m_def_GlassBoxColor
    m_PictureMaskColor = m_def_PictureMaskColor
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("GlassBoxColor", m_GlassBoxColor, m_def_GlassBoxColor)
    Call PropBag.WriteProperty("Picture", picBoxPicture.Picture, Nothing)
    Call PropBag.WriteProperty("PictureMaskColor", m_PictureMaskColor, m_def_PictureMaskColor)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  '
  'SET REFERENCE TO
  '    THE MESSAGEBOX FORM
  '    ITS YES BUTTON
  '    ITS NO BUTTON
  'BECAUSE WE NEED TO TRAP WHEN ITS CLICKED ON SO WE
  'CAN RETURN EITHER VBYES OR VBNO
  '(when buttons yes/no are used)
  '
  m_GlassBoxColor = PropBag.ReadProperty("GlassBoxColor", m_def_GlassBoxColor)
  Set picBoxPicture.Picture = PropBag.ReadProperty("Picture", Nothing)
  m_PictureMaskColor = PropBag.ReadProperty("PictureMaskColor", m_def_PictureMaskColor)

  If Ambient.UserMode Then
     With frmMessage
        Set F = frmMessage
        Set lblMsg = .lblMsg
        Set lblCaption = .lblCaption
        Set lblYes = .lblYes
        Set lblNo = .lblNo
        Set lblOk = .lblOk
        F.AutoRedraw = True
        Call DetermineSubtleLineColor
     End With
  Else
     Set lblYes = Nothing
     Set lblNo = Nothing
     Set lblOk = Nothing
     Set F = Nothing
  End If
  
  
End Sub


Private Sub MoveItem(item_hwnd As Long)
    '
    'for moving the messagebox around
    '
    On Error Resume Next
    Const WM_NCLBUTTONDOWN As Long = &HA1
    Const HTCAPTION As Long = 2
    ReleaseCapture
    SendMessage item_hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub
 


