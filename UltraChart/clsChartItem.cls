VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChartItem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public ItemID         As String
Public Description    As String

Public Category       As String
Public Series         As String

Private m_Left        As Long
Private m_Right       As Long
Private m_Loc         As Long
Private m_Top         As Long
Private m_Height      As Long
Private m_Width       As Long
Private m_ColorOne    As Long
Private m_ColorTwo    As Long

Private m_Host        As UltraChart
Private m_Value       As Double
Private m_OldValue    As Double
Private m_ValueHolder As Double

Friend Property Get Left() As Long
   Left = m_Left
End Property

Friend Property Let Left(ByVal vNewValue As Long)
   m_Left = vNewValue
End Property

Friend Property Get Right() As Long
   Right = m_Right
End Property

Friend Property Let Right(ByVal vNewValue As Long)
   m_Right = vNewValue
End Property

Friend Property Get Loc() As Long
   Loc = m_Loc
End Property

Friend Property Let Loc(ByVal vNewValue As Long)
   m_Loc = vNewValue
End Property

Friend Property Get Top() As Long
   Top = m_Top
End Property

Friend Property Let Top(ByVal vNewValue As Long)
   m_Top = vNewValue
End Property

Friend Property Get Height() As Long
   Height = m_Height
End Property

Friend Property Let Height(ByVal vNewValue As Long)
   m_Height = vNewValue
End Property

Friend Property Get Width() As Long
   Width = m_Width
End Property

Friend Property Let Width(ByVal vNewValue As Long)
   m_Width = vNewValue
End Property

Friend Property Get ColorOne() As Long
   ColorOne = m_ColorOne
End Property

Friend Property Let ColorOne(ByVal vNewValue As Long)
   m_ColorOne = vNewValue
End Property

Friend Property Get ColorTwo() As Long
   ColorTwo = m_ColorTwo
End Property

Friend Property Let ColorTwo(ByVal vNewValue As Long)
   m_ColorTwo = vNewValue
End Property

Friend Property Get Host() As UltraChart
   Host = m_Host
End Property

Friend Property Set Host(ByVal oNew As UltraChart)
   Set m_Host = oNew
End Property

Public Property Get Value() As Double
   Value = m_Value
End Property

Public Property Let Value(ByVal vNewValue As Double)
   m_OldValue = m_Value
   m_Value = vNewValue
   m_ValueHolder = vNewValue
   
   If Not m_Host Is Nothing Then m_Host.fChartItemChanged Me, m_OldValue
   
End Property

Friend Property Get OldValue() As Double
   OldValue = m_OldValue
End Property

Friend Property Let OldValue(ByVal vNewValue As Double)
   m_OldValue = vNewValue
End Property

Friend Sub fSetValue(ByVal NewValue As Double)
   m_Value = NewValue
End Sub

Friend Sub fSetValues()
   'Debug.Print m_Value, m_OldValue, m_ValueHolder,
   m_Value = m_ValueHolder
   m_OldValue = m_Value
   
   'Debug.Print m_Value, m_OldValue, m_ValueHolder
End Sub
