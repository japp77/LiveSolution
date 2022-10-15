VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_SearchValue As Variant
Private m_Name As String
Private m_Label As String
Private m_Control As Control
Private m_MaxLength As Integer
Private m_Visible As Boolean

Public Property Get SearchValue() As Variant
    SearchValue = m_SearchValue
End Property

Public Property Let SearchValue(ByVal NewValue As Variant)
    m_SearchValue = NewValue
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal NewValue As String)
    m_Name = NewValue
End Property

Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property

Public Property Let Visible(ByVal NewValue As Boolean)
    m_Visible = NewValue
End Property

Public Property Get Label() As String
    Label = m_Label
End Property

Public Property Let Label(ByVal NewValue As String)
    m_Label = NewValue
End Property

Public Property Get Control() As Control
    Set Control = m_Control
End Property

Public Property Set Control(ByVal NewValue As Control)
    Set m_Control = NewValue
End Property

Public Property Get MaxLength() As Integer
    MaxLength = m_MaxLength
End Property

Public Property Let MaxLength(ByVal NewValue As Integer)
    m_MaxLength = NewValue
End Property
