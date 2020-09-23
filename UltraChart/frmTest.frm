VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "UltraChart Test"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Random"
      Height          =   345
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   795
   End
   Begin ChartTest.UltraChart UltraChart 
      Height          =   5205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   9181
      uTopMargin      =   750
      uBottomMargin   =   825
      uLeftMargin     =   825
      uRightMargin    =   1700
      uContentBorder  =   -1  'True
      uSelectable     =   -1  'True
      uHotTracking    =   -1  'True
      uSelectedColumn =   -1
      uChartTitle     =   "UltraChart"
      uChartSubTitle  =   "The New Version"
      uDisplayCategory=   -1  'True
      uDisplayYAxis   =   -1  'True
      uColorBars      =   -1  'True
      uIntersectMajor =   10
      uIntersectMinor =   2
      uMaxYValue      =   100
      uDisplayDescript=   -1  'True
      uXAxisLabel     =   ""
      uYAxislabel     =   ""
      BackColor       =   -2147483643
      ForeColor       =   16777215
      ActiveTheme     =   0
      RefreshOnChangeValue=   0   'False
      TextDegree      =   90
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

   Dim Itm As clsChartItem

   For Each Itm In UltraChart.Items

      Itm.Value = CLng(Rnd * 2000)
   Next
    
   UltraChart.Refresh

End Sub

Private Sub Form_Load()

   Dim X As Integer, oChartItem As clsChartItem
   
   Randomize

   With UltraChart
      .Animate = True
      .RefreshOnChangeValue = False
      .TextDegree = 90
   End With

   For i = 7 To 9
      UltraChart.AddSeries "Y0" & i, "200" & i
   Next i
   
   Dim MyStr As String
   Dim Month As Variant
   
   MyStr = "Jan|Feb|Mar|Apr|May|Jun"
   Month = Split(MyStr, "|")
   
   For i = 1 To 6
      UltraChart.AddCategory "C" & i, Month(i - 1)
   Next i
   
   For X = 1 To 6
      For i = 7 To 9
      
         Set oChartItem = New clsChartItem

         With oChartItem
            .ItemID = X
            .Value = IIf(X = 1, 1900, CLng(Rnd * 2000))
            .Category = UltraChart.Category("C" & X)
            .Series = UltraChart.Series("Y0" & i)
                
            .Description = "Total Sale Of " & .Category & " " & .Series
         End With
            
         UltraChart.AddItem oChartItem
      Next i
   Next X

End Sub

Private Sub Form_Resize()
    
   UltraChart.Width = Me.ScaleWidth
   UltraChart.Height = Me.ScaleHeight
    
End Sub

