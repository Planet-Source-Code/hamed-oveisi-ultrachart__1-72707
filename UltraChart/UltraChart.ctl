VERSION 5.00
Begin VB.UserControl UltraChart 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   5580
   ScaleWidth      =   8400
   ToolboxBitmap   =   "UltraChart.ctx":0000
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   2565
   End
End
Attribute VB_Name = "UltraChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************
' UltraChart         By Hamed Oveisi
'                    Based On AnimatedChart

' What's New         Multi Series/Category
'                    Display data as ToolTip
'                    Animation On/Off
'                    Text Rotation Degree
'                    Full Animation on changing values
'                    Some more Enhancements ...
'
'
' Limitation         Still not supporting Negative values
'                    Missing Pie type;)
'                    Fixed coded colors
'***********************************************************************************

'Gradient Constants
Private Const GRADIENT_FILL_RECT_H   As Long = &H0
Private Const GRADIENT_FILL_RECT_V   As Long = &H1
Private Const GRADIENT_FILL_TRIANGLE As Long = &H2
Private Const GRADIENT_FILL_OP_FLAG  As Long = &HFF

Private Type TRIVERTEX          'For gradient Drawing

   X As Long
   Y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer

End Type

Private Type GRADIENT_RECT

   UPPERLEFT As Long
   LOWERRIGHT As Long

End Type

Enum GRADIENT_FILL_RECT

   FillHor = GRADIENT_FILL_RECT_H
   FillVer = GRADIENT_FILL_RECT_V

End Enum

Private Type GRADIENT_TRIANGLE

   Vertex1 As Long
   Vertex2 As Long
   Vertex3 As Long

End Type

Private Declare Function GradientFillRect _
                Lib "msimg32" _
                Alias "GradientFill" (ByVal hdc As Long, _
                                      pVertex As TRIVERTEX, _
                                      ByVal dwNumVertex As Long, _
                                      pMesh As GRADIENT_RECT, _
                                      ByVal dwNumMesh As Long, _
                                      ByVal dwMode As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetPixel _
                Lib "gdi32.dll" (ByVal hdc As Long, _
                                 ByVal X As Long, _
                                 ByVal Y As Long) As Long

Private uColumns()       As Double       'Array of column height values
'used to determine hittest feature.

Private uColWidth        As Double       'The calculated width of each column.
Private uRowHeight       As Double       'The calculated height of each column.
Private uTopMargin       As Double         '--------------------------------------
Private uBottomMargin    As Double         'Margins used around the chart content.
Private uLeftMargin      As Double         '
Private uRightMargin     As Double         '--------------------------------------
Private uContentBorder   As Boolean      'Border around the chart content?
Private uSelectable      As Boolean      'Marker indicating whether user can select a column.
Private uHotTracking     As Boolean      'Marker indicating use of hot tracking.
Private uSelectedColumn  As Double         'Marker indicating the selected column.
Private uOldSelection    As Double
Private uDisplayDescript As Boolean      'Display description when selectable
Private uChartTitle      As String       'Chart title
Private uChartSubTitle   As String       'Chart sub title
Private uDisplayCategory As Boolean      'Marker indicating display of x axis
Private uDisplayYAxis    As Boolean      'Marker indicating display of y axis
Private uColorBars       As Boolean      'Marker indicating use of different coloured bars
Private uIntersectMajor  As Double       'Major intersect value
Private uIntersectMinor  As Double       'Minor intersect value
Private uMaxYValue       As Double       'Default maximum y value
Private uXAxisLabel      As String       'Label to be displayed below the X-Axis
Private uYAxisLabel      As String       'Label to be displayed left of the Y-Axis

Public Items             As Collection   'Collection of chart items

Private offsetX          As Double
Private offsetY          As Double

Private bLegendAdded     As Boolean
Private bLegendClicked   As Boolean
Private bDisplayLegend   As Boolean
Private bResize          As Boolean

Private bProcessingOver  As Boolean      'Marker to speed up mouse over effects.

Public Enum Theme

   [ThemePersianGulf] = 0
   [ThemeSky] = 1
   [ThemeNeon] = 2
   [ThemeNormal] = 3

End Enum

Private m_ActiveTheme          As Theme
Private m_Animate              As Boolean
Private m_TextDegree           As Integer
Private m_RefreshOnChangeValue As Boolean
                           
Private IsDrawedOnce           As Boolean
Private IsInDrawMode           As Boolean

Private Colors(15, 1)          As Long
Private cItem()                As String
Private oToolTip               As New CTooltip
Private oSeries                As New Collection
Private oCategory              As New Collection

Private uCatSpace              As Long

Public Event ItemClick(cItem As clsChartItem)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Function AddItem(cItem As clsChartItem) As Boolean
   Items.Add cItem

   If cItem.Value > uMaxYValue Then
      uMaxYValue = cItem.Value
   End If

End Function

Public Function EditCopy() As Boolean
   Clipboard.SetData UserControl.Image
End Function

Public Property Let MarginTop(lMargin As Double)
   uTopMargin = lMargin * Screen.TwipsPerPixelY
   pDrawChart
End Property

Public Property Get MarginTop() As Double
   MarginTop = uTopMargin / Screen.TwipsPerPixelY
End Property

Public Property Let MarginBottom(lMargin As Double)
   uBottomMargin = lMargin * Screen.TwipsPerPixelY
   pDrawChart
End Property

Public Property Get MarginBottom() As Double
   MarginBottom = uBottomMargin / Screen.TwipsPerPixelY
End Property

Public Property Let MarginLeft(lMargin As Double)
   uLeftMargin = lMargin * Screen.TwipsPerPixelX
   pDrawChart
End Property

Public Property Get MarginLeft() As Double
   MarginLeft = uLeftMargin / Screen.TwipsPerPixelX
End Property

Public Property Let MarginRight(lMargin As Double)
   uRightMargin = lMargin * Screen.TwipsPerPixelX
   pDrawChart
End Property

Public Property Get MarginRight() As Double
   MarginRight = uRightMargin / Screen.TwipsPerPixelX
End Property

Public Property Let ContentBorder(DisplayBorder As Boolean)
   uContentBorder = DisplayBorder
   pDrawChart
End Property

Public Property Get ContentBorder() As Boolean
   ContentBorder = uContentBorder
End Property

Public Property Let Selectable(EnableSelection As Boolean)
Attribute Selectable.VB_Description = "If Column is selectable"
   uSelectable = EnableSelection
   pDrawChart
End Property

Public Property Get Selectable() As Boolean
   Selectable = uSelectable
End Property

Public Property Let HotTracking(UseHotTracking As Boolean)
   uHotTracking = UseHotTracking
   pDrawChart
End Property

Public Property Get HotTracking() As Boolean
   HotTracking = uHotTracking
End Property

Public Property Let SelectedColumn(ColNumber As Long)

   Dim ret   As Double
   Dim oItem As clsChartItem

   On Error Resume Next
    
   uSelectedColumn = ColNumber
   pDrawChart
    
   ret = uColumns(ColNumber)

   If Err.Number Then
      uSelectedColumn = -1
   Else
      Set oItem = Items(ColNumber + 1)
      RaiseEvent ItemClick(oItem)
   End If

End Property

Public Property Get SelectedColumn() As Long
Attribute SelectedColumn.VB_MemberFlags = "400"
   SelectedColumn = uSelectedColumn
End Property

Public Property Let ChartTitle(sTitle As String)
   uChartTitle = sTitle
   pDrawChart
End Property

Public Property Get ChartTitle() As String
   ChartTitle = uChartTitle
End Property

Public Property Let ChartSubTitle(sTitle As String)
   uChartSubTitle = sTitle
   pDrawChart
End Property

Public Property Get ChartSubTitle() As String
   ChartSubTitle = uChartSubTitle
End Property

Public Property Let IntersectMajor(ISValue As Double)
   uIntersectMajor = ISValue
   pDrawChart
End Property

Public Property Get IntersectMajor() As Double
   IntersectMajor = uIntersectMajor
End Property

Public Property Let IntersectMinor(ISValue As Double)
   uIntersectMinor = ISValue
   pDrawChart
End Property

Public Property Get IntersectMinor() As Double
   IntersectMinor = uIntersectMinor
End Property

Public Property Let DisplayYAxis(DisplayAxis As Boolean)
   uDisplayYAxis = DisplayAxis
   pDrawChart
End Property

Public Property Get DisplayYAxis() As Boolean
   DisplayYAxis = uDisplayYAxis
End Property

Public Property Let DisplayXAxis(DisplayAxis As Boolean)
   uDisplayCategory = DisplayAxis
   pDrawChart
End Property

Public Property Get DisplayXAxis() As Boolean
   DisplayXAxis = uDisplayCategory
End Property

Public Property Let MaxY(dMax As Double)
   uMaxYValue = dMax
   pDrawChart
End Property

Public Property Get MaxY() As Double
   MaxY = uMaxYValue
End Property

Public Property Let SelectionInformation(DisplayInfo As Boolean)
Attribute SelectionInformation.VB_Description = "Show Tooltip of information when a Column Selected"
   uDisplayDescript = DisplayInfo
   pDrawChart
End Property

Public Property Get SelectionInformation() As Boolean
   SelectionInformation = uDisplayDescript
End Property

Public Property Let AxisLabelY(sCaption As String)
   uYAxisLabel = sCaption
   pDrawChart
End Property

Public Property Get AxisLabelY() As String
   AxisLabelY = uYAxisLabel
End Property

Public Property Let AxisLabelX(sCaption As String)
   uXAxisLabel = sCaption
   pDrawChart
End Property

Public Property Get AxisLabelX() As String
   AxisLabelX = uXAxisLabel
End Property

Public Property Let BackColor(hColor As OLE_COLOR)
   UserControl.BackColor = hColor
   pDrawChart
End Property

Public Property Get BackColor() As OLE_COLOR
   BackColor = UserControl.BackColor
End Property

Public Property Let ForeColor(hColor As OLE_COLOR)
   UserControl.ForeColor = hColor
   pDrawChart
End Property

Public Property Get ForeColor() As OLE_COLOR

   ForeColor = UserControl.ForeColor
End Property

Public Property Let ColorBars(bUseColor As Boolean)
   uColorBars = bUseColor
   pDrawChart
End Property

Public Property Get ColorBars() As Boolean
   ColorBars = uColorBars
End Property

Private Sub tmrStart_Timer()
   IsDrawedOnce = False
   tmrStart.Enabled = False

   Call pDrawChart
End Sub

Private Sub UserControl_DragDrop(Source As Control, X As Single, Y As Single)
   Source.Left = X - offsetX
   Source.Top = Y - offsetY
End Sub

Private Sub UserControl_Initialize()
   Set Items = New Collection
   uCatSpace = 10
End Sub

Private Sub UserControl_InitProperties()

   Dim X          As Integer
   Dim oChartItem As clsChartItem
    
   uTopMargin = 50 * Screen.TwipsPerPixelY
   uBottomMargin = 55 * Screen.TwipsPerPixelY
   uLeftMargin = 55 * Screen.TwipsPerPixelX
   uRightMargin = 55 * Screen.TwipsPerPixelX
   uContentBorder = True
   uSelectable = False
   uHotTracking = False
   uSelectedColumn = -1
   uOldSelection = -1
   uChartTitle = UserControl.Name
   uChartSubTitle = "Ultra Chart"
   uDisplayYAxis = True
   uDisplayCategory = True
   uColorBars = False
   uIntersectMajor = 10
   uIntersectMinor = 2
   uMaxYValue = 100
    
   uCatSpace = 10
End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

        
   RaiseEvent MouseDown(Button, Shift, X, Y)
    
TrackExit:

   Exit Sub

End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

   Dim X1    As Long
   Dim oItem As clsChartItem

   X1 = (uColWidth)
    
   On Error GoTo TrackExit
    
   If IsInDrawMode Then GoTo TrackExit
    
   If uHotTracking Then
      If (Y <= UserControl.ScaleHeight - uBottomMargin) And uSelectable Then

         Dim oChartItm As clsChartItem

         For Each oChartItm In Items

            If oChartItm.Left <= X And oChartItm.Right >= X Then
               If oChartItm.Top <= Y Then
                  If Not bProcessingOver Then
                     bProcessingOver = True
                     uSelectedColumn = oChartItm.Loc  '(X - uLeftMargin) \ (X1)

                     If Not uSelectedColumn = uOldSelection Then
                        Cls
                        pDrawChart
                        uOldSelection = uSelectedColumn
                     End If

                     bProcessingOver = False
                  End If

                  Exit For

               Else

                  If Not bProcessingOver Then
                     bProcessingOver = True
                     uSelectedColumn = -1

                     If Not uSelectedColumn = uOldSelection Then
                        Cls
                        pDrawChart
                        uOldSelection = uSelectedColumn
                     End If

                     bProcessingOver = False
                  End If
               End If
            End If

         Next

      Else

         If Not bProcessingOver Then
            bProcessingOver = True
            uSelectedColumn = -1

            If Not uSelectedColumn = uOldSelection Then
               Cls
               pDrawChart
               uOldSelection = uSelectedColumn
            End If

            bProcessingOver = False
         End If
      End If
   End If

TrackExit:

   Exit Sub

End Sub

Public Sub Refresh()
   pDrawOldValToNewVal
End Sub

Public Sub Clear()

   Dim X As Integer
    
   For X = 1 To Items.Count
      Set Items(i) = Nothing
   Next X
    
   Set Items = Nothing
   Set Items = New Collection

   pDrawChart
End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
   RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   On Error Resume Next

   With PropBag
      uTopMargin = .ReadProperty("uTopMargin")
      uBottomMargin = .ReadProperty("uBottomMargin")
      uLeftMargin = .ReadProperty("uLeftMargin")
      uRightMargin = .ReadProperty("uRightMargin")
      uContentBorder = .ReadProperty("uContentBorder")
      uSelectable = .ReadProperty("uSelectable", False)
      uHotTracking = .ReadProperty("uHotTracking", False)
      uSelectedColumn = .ReadProperty("uSelectedColumn", -1)
      uChartTitle = .ReadProperty("uChartTitle", UserControl.Name)
      uChartSubTitle = .ReadProperty("uChartSubTitle", uChartSubTitle)
      uDisplayCategory = .ReadProperty("uDisplayCategory", uDisplayCategory)
      uDisplayYAxis = .ReadProperty("uDisplayYAxis", uDisplayYAxis)
      uColorBars = .ReadProperty("uColorBars", False)
      uIntersectMajor = .ReadProperty("uIntersectMajor", 10)
      uIntersectMinor = .ReadProperty("uIntersectMinor", 2)
      uMaxYValue = .ReadProperty("uMaxYValue", 100)
      uDisplayDescript = .ReadProperty("uDisplayDescript", False)
      uXAxisLabel = .ReadProperty("uXAxisLabel")
      uYAxisLabel = .ReadProperty("uYAxisLabel")
      UserControl.BackColor = .ReadProperty("BackColor")
      UserControl.ForeColor = .ReadProperty("ForeColor")
      uOldSelection = -1
      m_ActiveTheme = .ReadProperty("ActiveTheme", 0)
      m_Animate = .ReadProperty("Animate", True)
      m_RefreshOnChangeValue = .ReadProperty("RefreshOnChangeValue", True)
      m_TextDegree = .ReadProperty("TextDegree", 0)
   End With

End Sub

Private Sub UserControl_Resize()
    
   If IsDrawedOnce Then
      bResize = True
      pDrawChart
      bResize = False
   End If

End Sub

Private Sub UserControl_Show()
   'pDrawChart
   Call pSetStyle
   Call pSetColors
         
   UserControl.Cls
   pDrawBackTheme

   If m_Animate Then
      tmrStart.Enabled = True
   Else
      pDrawChart
   End If
   
End Sub

Private Sub UserControl_Terminate()

   On Error Resume Next

   Dim oChartItm As clsChartItem
    
   For Each oChartItm In Items

      Set oChartItm = Nothing
   Next

   Set Items = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "uTopMargin", uTopMargin
      .WriteProperty "uBottomMargin", uBottomMargin
      .WriteProperty "uLeftMargin", uLeftMargin
      .WriteProperty "uRightMargin", uRightMargin
      .WriteProperty "uContentBorder", uContentBorder
      .WriteProperty "uSelectable", uSelectable
      .WriteProperty "uHotTracking", uHotTracking
      .WriteProperty "uSelectedColumn", uSelectedColumn
      .WriteProperty "uChartTitle", uChartTitle
      .WriteProperty "uChartSubTitle", uChartSubTitle
      .WriteProperty "uDisplayCategory", uDisplayCategory
      .WriteProperty "uDisplayYAxis", uDisplayYAxis
      .WriteProperty "uColorBars", uColorBars
      .WriteProperty "uIntersectMajor", uIntersectMajor
      .WriteProperty "uIntersectMinor", uIntersectMinor
      .WriteProperty "uMaxYValue", uMaxYValue
      .WriteProperty "uDisplayDescript", uDisplayDescript
      .WriteProperty "uXAxisLabel", uXAxisLabel
      .WriteProperty "uYAxislabel", uYAxisLabel
      .WriteProperty "BackColor", UserControl.BackColor
      .WriteProperty "ForeColor", UserControl.ForeColor
      .WriteProperty "ActiveTheme", m_ActiveTheme
      .WriteProperty "Animate", m_Animate, True
      .WriteProperty "RefreshOnChangeValue", m_RefreshOnChangeValue, True
      .WriteProperty "TextDegree", m_TextDegree, 0
   End With

End Sub

Private Function pDoGradient(FromColor As Long, _
                             ToColor As Long, _
                             Optional DrawHorVer As GRADIENT_FILL_RECT = FillHor, _
                             Optional Left As Long = 0, _
                             Optional Top As Long = 0, _
                             Optional Width As Long = -1, _
                             Optional Height As Long = -1, _
                             Optional ByVal Drawhdc As Long = -1) As Boolean

   Dim Vert(1) As TRIVERTEX
   Dim gRect   As GRADIENT_RECT
   Dim r       As Byte, G As Byte, B As Byte
       
   pLong2RGB FromColor, r, G, B

   With Vert(0)
      .X = Left
      .Y = Top
      .Red = Val("&h" & Hex(r) & "00")
      .Green = Val("&h" & Hex(G) & "00")
      .Blue = Val("&h" & Hex(B) & "00")
      .Alpha = 0&
   End With
    
   pLong2RGB ToColor, r, G, B

   With Vert(1)
      .X = Left + Width
      .Y = Top + Height
      .Red = Val("&h" & Hex(r) & "00")
      .Green = Val("&h" & Hex(G) & "00")
      .Blue = Val("&h" & Hex(B) & "00")
      .Alpha = 0&
   End With

   gRect.UPPERLEFT = 0
   gRect.LOWERRIGHT = 1

   pDoGradient = GradientFillRect(IIf(Drawhdc = -1, UserControl.hdc, Drawhdc), Vert(0), 2, gRect, 1, DrawHorVer)
    
End Function

Private Function pLong2RGB(nColor As Long, Red As Byte, Green As Byte, Blue As Byte)
   Red = (nColor And &HFF&)
   Green = (nColor And &HFF00&) / &H100
   Blue = (nColor And &HFF0000) / &H10000
End Function

Private Sub pDrawItem(ByVal ColorOne As Long, _
                      ByVal ColorTwo As Long, _
                      ByVal Left As Long, _
                      ByVal Top As Long, _
                      ByVal Width As Long, _
                      ByVal Height As Long, _
                      Optional ByVal Animated As Boolean = False)

   Select Case m_ActiveTheme

      Case ThemePersianGulf
         pDoGradient ColorTwo, ColorOne, FillVer, Left, Top, Width, Height
         pDoGradient ColorOne, ColorTwo, FillVer, Left + 2, Top + 2, Width - 4, Height - 4

      Case Else
         pDoGradient ColorTwo, ColorOne, FillHor, Left, Top, Width, Height
         
         pDoGradient ColorOne, ColorTwo, FillHor, Left + 2, Top + 2, (Width / 3) * 2 - 4, Height - 4
         pDoGradient ColorTwo, ColorOne, FillHor, Left + (Width / 3) * 2 - 2, Top + 2, (Width / 3) - 1, Height - 4
   End Select

End Sub

Private Sub pDrawAllItems()

   Dim i      As Double
   Dim Down   As Long
  
   Dim lStep  As Long
   Dim Item   As clsChartItem
   Dim xTwips As Long

   xTwips = Screen.TwipsPerPixelX
   
   lStep = IIf(m_Animate, 10, 1)
   
   On Error GoTo Er

   For i = 1 To lStep

      For Each Item In Items

         With Item

            Down = .Top + .Height

            If .Height >= (.Height / lStep) * i Then
               pDrawItem .ColorOne, .ColorTwo, .Left / xTwips, (Down - ((.Height / lStep) * i)) / xTwips, (.Width) / xTwips, ((.Height / lStep) * i) / xTwips
               X1 = uLeftMargin + (2 * Screen.TwipsPerPixelX): X2 = UserControl.ScaleWidth - uRightMargin
               Y1 = (UserControl.ScaleHeight - uBottomMargin) - (X * uRowHeight)
               UserControl.Line (X1, Y1)-(X2 + 1, Y1 + 15), pGetThemeLineColor, BF
               .fSetValues
            End If

         End With

      Next
      
      Tim = Timer

      Do While Timer - Tim < 0.07: Loop
      
      UserControl.Refresh
      
   Next i
   
Er:
   IsDrawedOnce = True

End Sub

Private Sub pSetColors()
   Colors(0, 0) = RGB(185, 239, 255): Colors(0, 1) = RGB(30, 155, 230)
   Colors(1, 0) = RGB(255, 125, 79): Colors(1, 1) = RGB(129, 0, 0)
   Colors(2, 0) = RGB(0, 254, 0): Colors(2, 1) = RGB(0, 122, 0)
   Colors(3, 0) = RGB(233, 131, 255): Colors(3, 1) = RGB(214, 23, 255)
   Colors(4, 0) = RGB(95, 206, 255): Colors(4, 1) = RGB(0, 116, 210)
   Colors(5, 0) = RGB(255, 193, 66): Colors(5, 1) = RGB(185, 0, 0)
   Colors(6, 0) = RGB(215, 255, 168): Colors(6, 1) = RGB(99, 163, 23)
   Colors(7, 0) = RGB(201, 61, 154): Colors(7, 1) = RGB(153, 13, 106)
   Colors(8, 0) = RGB(0, 0, 254): Colors(8, 1) = RGB(0, 0, 122)
   Colors(9, 0) = RGB(255, 255, 160): Colors(9, 1) = RGB(250, 197, 12)
End Sub

Private Function pGetYTopLegend(ByVal MaxChartValue As Long) As Long

   Dim Text  As String
   Dim MyStr As String
   Dim Num   As Long
   
   Text = CStr(MaxChartValue)
   
   If Val(Text) > 10 Then
      MyStr = String(Len(Text) - 2, "0")
      
      Num = Val(Left(Text, 2))
      
      If Num Mod 10 = 0 Then
         MyStr = Num & MyStr
      ElseIf Num Mod 10 > 5 Then
         MyStr = CStr(Int(Num / 10) + 1) & "0" & MyStr
      Else
         MyStr = CStr(Int(Num / 10)) & "5" & MyStr
      End If

   Else
      MyStr = 10
   End If
   
   pGetYTopLegend = CLng(MyStr)
End Function

Public Property Get ActiveTheme() As Theme
   ActiveTheme = m_ActiveTheme
End Property

Public Property Let ActiveTheme(ByVal NewTheme As Theme)
   m_ActiveTheme = NewTheme
   
   pSetStyle

   PropertyChanged "ActiveTheme"
   
   pDrawChart
End Property

Private Sub pDrawChart()

   Dim CurrentColor As Integer
   Dim iCols        As Integer
   Dim X            As Double
   Dim X1           As Double
   Dim X2           As Double
   Dim Y1           As Double
   Dim y2           As Double
   Dim xTemp        As Double
   Dim yTemp        As Double
   Dim sDescription As String
   Dim oChartItem   As clsChartItem
   Dim lTopYValue   As Double
    
   If IsInDrawMode Then Exit Sub

   IsInDrawMode = True
    
   lTopYValue = pGetYTopLegend(uMaxYValue)
    
   uIntersectMajor = lTopYValue / 10
    
   iCols = Items.Count
    
   uRowHeight = lTopYValue

   For X = 1 To Items.Count
      Set oChartItem = Items(X)

      If uRowHeight - CDbl(oChartItem.Value) < 0 Then uRowHeight = CDbl(oChartItem.Value)
   Next X
    
   If uRowHeight = 0 Then uRowHeight = 0.001
    
   If uMaxYValue < uRowHeight Then uMaxYValue = uRowHeight
    
   With UserControl
      uRowHeight = ((.ScaleHeight - (uTopMargin + uBottomMargin)) / uRowHeight)

      If iCols Then uColWidth = (.ScaleWidth - (uLeftMargin + uRightMargin) - (oCategory.Count - 1) * uCatSpace * Screen.TwipsPerPixelX) / iCols    '/ 2.5
   End With
    
   pDrawBackTheme
        
   If iCols Then ReDim uColumns(iCols - 1)

   On Error Resume Next

   'Intersect lines
    
   With UserControl
      .CurrentX = (.ScaleWidth / 2) - (.TextWidth(uChartTitle) / 2)
      .CurrentY = 0
      .FontBold = True
      UserControl.Print uChartTitle
      .FontBold = False
        
      .FontSize = .FontSize - 2
      .CurrentX = (.ScaleWidth / 2) - (.TextWidth(uChartSubTitle) / 2)
      UserControl.Print uChartSubTitle
      .FontSize = .FontSize + 2
   End With
    
   If uDisplayYAxis Then

      Dim Counter  As Double
      Dim LastLine As Double

      For X = 0 To lTopYValue Step lTopYValue * 0.1
         X1 = uLeftMargin + (2 * Screen.TwipsPerPixelX): X2 = UserControl.ScaleWidth - uRightMargin
         Y1 = (UserControl.ScaleHeight - uBottomMargin) - (X * uRowHeight)

         If (X) Mod uIntersectMajor = 0 Then
            Counter = Counter + 1
                
            If Counter Mod 2 = 0 Then
               pDrawIntersect X1, Y1, X2, LastLine
            Else
               LastLine = Y1
            End If

            UserControl.Line (X1, Y1)-(X2 + 1, Y1 + 15), pGetThemeLineColor, BF
            UserControl.FontSize = UserControl.FontSize - 2
            UserControl.CurrentX = uLeftMargin - UserControl.TextWidth(X) - (5 * Screen.TwipsPerPixelX)
            UserControl.CurrentY = Y1 - (UserControl.TextHeight("0") / 2)
            UserControl.Print (X)
            UserControl.FontSize = UserControl.FontSize + 2
         End If

      Next X

   End If
   
   ReDim cItem(Items.Count - 1)
    
   'On Error GoTo 0
   If uContentBorder Then
      UserControl.Line (uLeftMargin - 15, uTopMargin)-(uLeftMargin, UserControl.ScaleHeight - uBottomMargin), pGetThemeLineColor, BF
   End If
   
   If oToolTip.IsShowed Then oToolTip.Destroy

   Dim S             As Long
   Dim sCategoryText As String
   Dim LastX         As Long
    
   X = 0

   For S = 1 To oCategory.Count
      sCategoryText = oCategory(S)
       
      For i = 0 To Items.Count - 1
         Set oChartItem = Items(i + 1)
           
         If oChartItem.Category = sCategoryText Then
              
            X1 = (X * uColWidth) + uLeftMargin + (S - 1) * (uCatSpace * Screen.TwipsPerPixelX)  '(2 * Screen.TwipsPerPixelX)
            X2 = X1 + uColWidth - (2 * Screen.TwipsPerPixelX)
             
            Y1 = (UserControl.ScaleHeight - uBottomMargin) - (CDbl(oChartItem.Value) * uRowHeight)
            y2 = UserControl.ScaleHeight - uBottomMargin
                      
            With oChartItem
               .Left = X1
               .Right = X2
               .Loc = X
               .Top = Y1
               .Height = y2 - Y1 - 1
               .Width = X2 - X1
               Set .Host = Me
            End With
              
            uColumns(X) = Y1
                           
            'Selected bar outline
            If X = uSelectedColumn And uSelectable And IsDrawedOnce Then
               pDrawItem RGB(252, 233, 179), RGB(244, 192, 51), (X1 + 1) / 15, Y1 / 15, (X2 - X1 - 1) / 15, (y2 - Y1) / 15, False
                                                            
               If uDisplayDescript Then

                  With oToolTip

                     If .IsShowed Then .Destroy
                     .Title = oChartItem.Category & "  " & oChartItem.Series & " : " & Format(oChartItem.Value, "#,0")
                     .TipText = oChartItem.Description
                     .Create UserControl.hWnd
                  End With

               End If
                  
            Else

               If oSeries.Count = 1 Then
                  CurrentColor = (oChartItem.ItemID - 1) Mod 10
               Else
                  CurrentColor = (pGetSeriesIndex(oChartItem.Series) - 1) Mod 10
               End If
                  
               If Not IsDrawedOnce Then

                  With oChartItem
                     .ColorOne = IIf(uColorBars, Colors(CurrentColor, 0), Colors(2, 0))
                     .ColorTwo = IIf(uColorBars, Colors(CurrentColor, 1), Colors(2, 1))
                  End With

               Else

                  Dim ColorOne As Long
                  Dim ColorTwo As Long

                  ColorOne = oChartItem.ColorOne  'IIf(uColorBars, Colors(CurrentColor, 0), Colors(2, 0))
                  ColorTwo = oChartItem.ColorTwo ' IIf(uColorBars, Colors(CurrentColor, 1), Colors(2, 1))
                  pDrawItem ColorOne, ColorTwo, (X1 + 1) / 15, Y1 / 15, (X2 - X1 - 1) / 15, (y2 - Y1) / 15, Not IsDrawedOnce
                     
               End If
      
            End If
              
            X = X + 1
         End If

      Next i
       
      If uDisplayCategory And X > LastX Then
         UserControl.FontSize = UserControl.FontSize - 1
          
         X1 = (LastX * uColWidth) + uLeftMargin + (S - 1) * (uCatSpace * Screen.TwipsPerPixelX)
         X2 = (X - LastX) * uColWidth + X1

         xTemp = (((X2 - X1) / 2) + X1) / Screen.TwipsPerPixelX 'X1 / Screen.TwipsPerPixelX
         yTemp = (UserControl.ScaleHeight - uBottomMargin + UserControl.TextWidth(oChartItem.Category) / 1.25) / Screen.TwipsPerPixelY
          
         PrintRotText UserControl.hdc, sCategoryText, xTemp, yTemp, m_TextDegree
          
         If LastX > 0 Then
            UserControl.Line (X1 - (uCatSpace / 2) * Screen.TwipsPerPixelX, y2)-(X1 - (uCatSpace / 2) * Screen.TwipsPerPixelX, y2 + UserControl.TextHeight(oChartItem.Category) / 2), pGetThemeLineColor
         End If
          
         UserControl.FontSize = UserControl.FontSize + 1
          
         LastX = X '+ 1
      End If
       
   Next S
    
   UserControl.CurrentY = (Height - TextHeight(oSeries.Count)) / 2
    
   For S = 1 To oSeries.Count
      CurrentColor = (S - 1) Mod 10
      
      With UserControl
         .CurrentX = (Width - uRightMargin + 400)
         Y1 = .CurrentY
         UserControl.Line (.CurrentX - 100, .CurrentY + 30)-(.CurrentX - 230, .CurrentY + TextHeight(oSeries(S)) - 30), Colors(CurrentColor, 0), BF
         .CurrentY = Y1
         .CurrentX = (Width - uRightMargin + 400)
      End With
      
      Print oSeries(S)
   Next S
    
   If Not IsDrawedOnce Then pDrawAllItems
    
   'Print the x axis label
   If Len(uXAxisLabel) Then
      UserControl.FontSize = UserControl.FontSize - 1
      UserControl.CurrentY = UserControl.ScaleHeight - UserControl.TextHeight(uXAxisLabel) * 1.5
      UserControl.CurrentX = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(uXAxisLabel) / 2)
      UserControl.Print uXAxisLabel
      UserControl.FontSize = UserControl.FontSize + 1
   End If
    
   'Print the y axis label
   If Len(uYAxisLabel) Then
      UserControl.FontSize = UserControl.FontSize - 1
      PrintRotText UserControl.hdc, uYAxisLabel, UserControl.TextHeight(uYAxisLabel) / Screen.TwipsPerPixelX, UserControl.ScaleHeight / 2 / Screen.TwipsPerPixelY, 90
      UserControl.FontSize = UserControl.FontSize + 1
   End If
    
   IsInDrawMode = False
    
End Sub

Private Sub pDrawBackTheme()

   Dim lWidth  As Long
   Dim lHeight As Long

   lWidth = (UserControl.ScaleWidth) / Screen.TwipsPerPixelX
   lHeight = (UserControl.ScaleHeight / Screen.TwipsPerPixelY)
   
   UserControl.Cls

   Select Case m_ActiveTheme

      Case ThemePersianGulf
         pDoGradient RGB(0, 3, 102), RGB(0, 100, 202), FillVer, 0, 0, lWidth, lHeight, UserControl.hdc

      Case ThemeSky
         pDoGradient RGB(158, 190, 230), RGB(185, 210, 239), FillVer, 0, 0, lWidth, lHeight

      Case ThemeNeon
         pDoGradient RGB(0, 0, 0), RGB(75, 75, 75), FillVer, 0, 0, lWidth, lHeight
   End Select

End Sub

Private Sub pDrawIntersect(ByVal X1 As Long, _
                           ByVal Y1 As Long, _
                           ByVal X2 As Long, _
                           ByVal LastLine As Long)

   Dim lHeight As Long
   Dim lWidth  As Long
   Dim lLeft   As Long
   Dim lTop    As Long
   
   lHeight = ((Y1 - LastLine) / 15) - 1
   lWidth = (X2 - X1) / 15 + 1
   lLeft = X1 / 15
   lTop = (Y1 / 15) + 1

   Select Case m_ActiveTheme

      Case ThemePersianGulf
         pDoGradient RGB(0, 54, 144), RGB(0, 59, 149), FillVer, lLeft, lTop, lWidth, lHeight
         UserControl.Line (X1, Y1)-(X2 + 1, Y1 + 30), RGB(0, 129, 199), BF

      Case ThemeSky
         pDoGradient RGB(227, 239, 255), RGB(201, 224, 255), FillVer, lLeft, lTop, lWidth, (lHeight / 2)
         pDoGradient RGB(183, 214, 255), RGB(190, 218, 255), FillVer, lLeft, lTop + (lHeight / 2), lWidth, lHeight - (lHeight / 2) + 1

      Case ThemeNeon
         pDoGradient RGB(66, 70, 81), RGB(58, 61, 69), FillVer, lLeft, lTop, lWidth, (lHeight / 8) * 3
         pDoGradient RGB(46, 47, 47), RGB(59, 59, 59), FillVer, lLeft, lTop + (lHeight / 8) * 3, lWidth, (lHeight / 8) * 4
         pDoGradient RGB(68, 68, 68), RGB(75, 75, 75), FillVer, lLeft, lTop + ((lHeight / 8) * 7) - 1, lWidth, lHeight - (lHeight / 8) * 7 + 1
   End Select

End Sub

Private Function pGetThemeLineColor() As Long

   Select Case m_ActiveTheme

      Case ThemePersianGulf
         pGetThemeLineColor = RGB(0, 129, 199)

      Case ThemeNeon
         pGetThemeLineColor = RGB(40, 40, 40)

      Case ThemeSky
         pGetThemeLineColor = RGB(141, 178, 227) 'RGB(173, 209, 255) '
   End Select

End Function

Private Sub pSetStyle()

   Select Case m_ActiveTheme

      Case ThemePersianGulf
         UserControl.ForeColor = vbWhite

      Case ThemeSky
         UserControl.ForeColor = vbBlack 'RGB(131, 200, 240)

      Case ThemeNeon
         UserControl.ForeColor = vbWhite
   End Select

End Sub

Public Sub AddSeries(ByVal Key As String, Optional ByVal Text As String)

   On Error Resume Next
   
   If Len(Text) = 0 Then Text = Key

   oSeries.Add Text, Key
End Sub

Public Function Series(ByVal Key As String) As String

   On Error Resume Next

   Series = oSeries(Key)
   
End Function

Public Sub AddCategory(ByVal Key As String, Optional ByVal Text As String)

   On Error Resume Next
   
   If Len(Text) = 0 Then Text = Key

   oCategory.Add Text, Key
End Sub

Public Function Category(ByVal Key As String) As String

   On Error Resume Next

   Category = oCategory(Key)
   
End Function

Private Function pGetSeriesIndex(ByVal Text As String) As Long

   Dim i As Long

   For i = 1 To oSeries.Count

      If oSeries(i) = Text Then pGetSeriesIndex = i: Exit Function
   Next i

End Function

Friend Sub fChartItemChanged(ByVal oChartItem As clsChartItem, ByVal OldValue As Long)

   Dim i     As Double
   Dim lStep As Long
   
   If m_RefreshOnChangeValue Then
      If m_Animate Then
         lStep = (oChartItem.Value - OldValue) / 5
         
         oChartItem.fSetValue OldValue

         For i = 1 To 5
            oChartItem.fSetValue OldValue + lStep * i
            pDrawChart
            Tim = Timer

            Do While Timer - Tim < 0.07: Loop
      
            UserControl.Refresh
         Next i

      End If

      oChartItem.fSetValues
      pDrawChart
   End If

End Sub

Public Property Get Animate() As Boolean
Attribute Animate.VB_Description = "using Animation"
   Animate = m_Animate
End Property

Public Property Let Animate(ByVal vNewValue As Boolean)
   m_Animate = vNewValue
   PropertyChanged "Animate"
End Property

Public Property Get TextDegree() As Integer
Attribute TextDegree.VB_Description = "Text Rotation Degree"
   TextDegree = m_TextDegree
End Property

Public Property Let TextDegree(ByVal iNewValue As Integer)

   If iNewValue >= 0 And iNewValue <= 360 Then
      m_TextDegree = iNewValue
      PropertyChanged "TextDegree"
   End If

End Property

Private Sub pDrawOldValToNewVal()
   
   Dim oChartItm As clsChartItem
   
   Dim i         As Integer
   Dim j         As Integer
   Dim lStep()   As Long
   Dim OldValue  As Double
   
   ReDim lStep(Items.Count)
   
   If m_Animate Then

      For i = 1 To 5
         For j = 1 To Items.Count
            Set oChartItm = Items(j)
            OldValue = oChartItm.OldValue
            
            If i = 1 Then
               If OldValue <> oChartItm.Value Then lStep(j) = (oChartItm.Value - OldValue) / 5
            End If
            
            If lStep(j) <> 0 Then
               oChartItm.fSetValue OldValue + lStep(j) * i
            End If

         Next
         
         pDrawChart
         Tim = Timer

         Do While Timer - Tim < 0.07: Loop
   
         UserControl.Refresh
      Next i

   End If
   
   For Each oChartItm In Items

      oChartItm.fSetValues
   Next

   pDrawChart
   UserControl.Refresh
   
   Set oChartItm = Nothing
End Sub

Public Property Get RefreshOnChangeValue() As Boolean
Attribute RefreshOnChangeValue.VB_Description = "Refresh automatically when Value changes"
   RefreshOnChangeValue = m_RefreshOnChangeValue
End Property

Public Property Let RefreshOnChangeValue(ByVal bNewValue As Boolean)
   m_RefreshOnChangeValue = bNewValue
   PropertyChanged "RefreshOnChangeValue"
End Property

Public Sub SaveChart(ByVal FileName As String)

   On Error Resume Next

   SavePicture UserControl.Image, FileName
End Sub

Public Property Get Picture() As StdPicture
   Set Picture = UserControl.Image
End Property

