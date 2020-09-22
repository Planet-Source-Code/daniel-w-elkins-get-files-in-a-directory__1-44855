VERSION 5.00
Begin VB.UserControl Button 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   375
   ScaleWidth      =   1455
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TDButton"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      MouseIcon       =   "Button.ctx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Image ButtonUp 
      Height          =   375
      Left            =   0
      Picture         =   "Button.ctx":0152
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1455
   End
   Begin VB.Image ButtonDown 
      Height          =   375
      Left            =   0
      Picture         =   "Button.ctx":2456
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click()

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ButtonUp.Visible = False
ButtonDown.Visible = True
End If
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ButtonDown.Visible = False
ButtonUp.Visible = True
RaiseEvent Click
End If

End Sub

Private Sub UserControl_InitProperties()
Caption = Ambient.DisplayName
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
End Sub

Private Sub UserControl_Resize()
ButtonUp.Height = UserControl.Height
ButtonDown.Height = ButtonUp.Height
ButtonUp.Width = UserControl.Width
ButtonDown.Width = ButtonUp.Width
lblCaption.Width = ButtonUp.Width
lblCaption.Top = (UserControl.Height * 0.5) - 100
End Sub

Public Property Get Caption() As String
Caption = lblCaption.Caption

End Property

Public Property Let Caption(ByVal NewValue As String)
lblCaption.Caption = NewValue
UserControl.PropertyChanged "Caption"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Caption", Caption, Ambient.DisplayName
End Sub
