Attribute VB_Name = "BorderMod"
Option Explicit
Private clsBorders As cBorders

Public Sub SetCustomBorderControll(Style As BorderrStyle, MyControll As Control, RGB_Color As String)
'          Dim MyControll As Control
          Dim MyUserControl As UserControl, RGBArray() As String, Red As Integer, Green As Integer, Blue As Integer
          
21320     If RGB_Color <> "" Then
21330         RGBArray = Split(RGB_Color, ",")
21340         Red = val(RGBArray(0))
21350         Green = val(RGBArray(1))
21360         Blue = val(RGBArray(2))
21370     End If
      '    On Error GoTo Fout


          ' In the SetBorder call below, following parameters should be explained
          
          ' ///// Border Styles \\\\\
          ' bsFlat1Color. Solid 1-pixel border, 1 color (i.e., flat).
          '       Uses Shadow only
          ' bsFlat2Color. Left/Top borders are 1 color, right/bottom are another
          '       Uses Shadow & Highlight only
          ' bsSunken. Left/Top outer border are Shadow, Right/Bottom outer are HighLight
          '           Left/Top inner border are DarkShadow, Right/Bottom inner are LightShadow
          ' bsRaised. Left/Top outer border are HighLight, Right/Bottom outer are DarkShadow
          '           Left/Top inner border are LightShadow, Right/Bottom outer are Shadow
          
          ' ///// Colors \\\\\ vb system colors can be passed
          ' Shadow: 2nd darkest of 4 color borders; color for a single color border
          ' DarkShadow: the darkest of 4 color borders
          ' LightShadow: 2nd lightest of 4 color borders
          ' Highlight: lightest of 4 color borders
          ' Special values for the above 4 colors
          '   -1 = AutoShade. DarkShadow, LightShadow & Highlight are shades of Shadow
          '           DarkShadow = Shadow darkened by 90% or black whichever is greater
          '           LightShadow = Shadow lightened by 85% of its lightest value (white)
          '           Highlight = Shadow lightened by 100% (or vbWhite)
          '   -2 = System Colors: vb3DDKShadow, vbButtonShadow, vb3DLight, vbHighlight respectively
          '   -3 & -4 (Reserved) are used by the class to fake single borders on combo boxes
          
          ' ///// Control Type \\\\\
          ' Some controls have their borders drawn be VB on the control's client area whereas
          '   others are drawn in the non-client area as expected. Think of a form with no
          '   borders but you want borders so you draw it on the form (non-client area).
          '   VB combo boxes are very much like that scenario. Therefore, the control type
          '   needs to be known in advance so the class can handle those special cases.
          ' ctComboBox: use for comboboxes and drivecombo
          ' ctImageCombo: use for the image combobox
          ' ctListBox: use for listboxes and file listboxes
          ' ctOther: use for other controls like treeview, listview, progressbar, etc
          
          
          Dim borderType As BorderStyleOptions
          Dim Colors(0 To 3) As Long
          
          ' for all custom color samples, we will let the class autoshade the
          '   necessary colors based off of the passed primary color (vbBlue-ish)
          '   however, you can dictate any of the 4 colors
21380     Colors(0) = &HFF8080
21390     Colors(1) = bsAutoShade: Colors(2) = bsAutoShade: Colors(3) = bsAutoShade

21400     Select Case Style
              Case FLAT_SINGLE_COLOR   ' flat single color
21410             borderType = bsFlat1Color
21420         Case FLAT_TWO_COLOR  ' flat two color
21430             borderType = bsFlat2Color
21440         Case SUNKEN_CUSTOM  ' sunken custom
21450             borderType = bsSunken
21460         Case RAISED_CUSTOM  ' raised custom
21470             borderType = bsRaised
21480         Case FLAT_SYSTEM_COLOR ' flat dual using system colors
21490             Colors(0) = bsSysDefault
21500             Colors(3) = bsSysDefault
21510             borderType = bsFlat2Color
21520         Case FLAT_CUSTOM_COLOR
21530             Colors(0) = RGB(Red, Green, Blue) ' bsSysDefault
21540             Colors(2) = RGB(Red, Green, Blue) ' bsSysDefault
21550             borderType = bsFlat2Color
21560         Case RAISED_SYSTEM_COLOR ' raised using system colors
21570             Colors(0) = bsSysDefault: Colors(1) = bsSysDefault
21580             Colors(2) = bsSysDefault: Colors(3) = bsSysDefault
21590             borderType = bsRaised
21600         Case SUNKEN_SYSTEM_COLOR ' sunken using system colors
21610             Colors(0) = bsSysDefault: Colors(1) = bsSysDefault
21620             Colors(2) = bsSysDefault: Colors(3) = bsSysDefault
21630             borderType = bsSunken
21640     End Select
          
21650     If clsBorders Is Nothing Then
21660         Set clsBorders = New cBorders
21670     End If
          

          

'21700         If TypeName(MyControll) = SpecifiekeControllType Or SpecifiekeControllType = "" Then
                       If TypeName(MyControll) = "Label" Or TypeName(MyControll) = "CheckBox" Or TypeName(MyControll) = "OptionButton" Or TypeName(MyControll) = "Frame" Then
21710             ElseIf TypeName(MyControll) = "TextBox" Then
21720                 clsBorders.SetBorder MyControll.hwnd, borderType, ctTextBox, Colors(0), Colors(1), Colors(2), Colors(3)
21730             ElseIf TypeName(MyControll) = "ListBox" Or TypeName(MyControll) = "FileListBox" Then
21740                 clsBorders.SetBorder MyControll.hwnd, borderType, ctListBox, Colors(0), Colors(1), Colors(2), Colors(3)
21750             ElseIf TypeName(MyControll) = "ComboBox" Or TypeName(MyControll) = "DriveListBox" Or TypeName(MyControll) = "DTPicker" Then
21760                 clsBorders.SetBorder MyControll.hwnd, borderType, ctComboBox, Colors(0), Colors(1), Colors(2), Colors(3)
21770             ElseIf TypeName(MyControll) = "ImageCombo" Then
21780                 clsBorders.SetBorder MyControll.hwnd, borderType, ctImageCombo, Colors(0), Colors(1), Colors(2), Colors(3)
21790             ElseIf TypeName(MyControll) = "vbalImageList" Or TypeName(MyControll) = "Timer" Or TypeName(MyControll) = "VistaForm" Or TypeName(MyControll) = "Shape" Or TypeName(MyControll) = "Label" Or TypeName(MyControll) = "Image" Or TypeName(MyControll) = "ImageList" Or TypeName(MyControll) = "WebBrowser" Or TypeName(MyControll) = "vbalDTabControlX" Then
21800             Else
21810                     clsBorders.SetBorder MyControll.hwnd, borderType, , Colors(0), Colors(1), Colors(2), Colors(3)
21820             End If
'21830         End If

              
              
End Sub
Public Sub SetCustomBorderControllSpecifiekeControll(Style As String, SpecifiekeControllType As String, RGB_Color As String, SpecifiekeControllName As String)
          Dim MyControll As Control
          Dim MyUserControl As UserControl, RGBArray() As String, Red As Integer, Green As Integer, Blue As Integer
          
21850     If RGB_Color <> "" Then
21860         RGBArray = Split(RGB_Color, ",")
21870         Red = val(RGBArray(0))
21880         Green = val(RGBArray(1))
21890         Blue = val(RGBArray(2))
21900     End If
      '    On Error GoTo Fout
          
          
          Dim borderType As BorderStyleOptions
          Dim Colors(0 To 3) As Long
          
          ' for all custom color samples, we will let the class autoshade the
          '   necessary colors based off of the passed primary color (vbBlue-ish)
          '   however, you can dictate any of the 4 colors
21910     Colors(0) = &HFF8080
21920     Colors(1) = bsAutoShade: Colors(2) = bsAutoShade: Colors(3) = bsAutoShade
          
21930     Select Case Style
              Case "FLAT_SINGLE_COLOR" ' flat single color
21940             borderType = bsFlat1Color
21950         Case "FLAT_TWO_COLOR"  ' flat two color
21960             borderType = bsFlat2Color
21970         Case "SUNKEN_CUSTOM"  ' sunken custom
21980             borderType = bsSunken
21990         Case "RAISED_CUSTOM"  ' raised custom
22000             borderType = bsRaised
22010         Case "FLAT_SYSTEM_COLOR" ' flat dual using system colors
22020             Colors(0) = bsSysDefault
22030             Colors(3) = bsSysDefault
22040             borderType = bsFlat2Color
22050         Case "FLAT_CUSTOM_COLOR"
22060             Colors(0) = RGB(Red, Green, Blue) ' bsSysDefault
22070             Colors(2) = RGB(Red, Green, Blue) ' bsSysDefault
22080             borderType = bsFlat2Color
22090         Case "RAISED_SYSTEM_COLOR" ' raised using system colors
22100             Colors(0) = bsSysDefault: Colors(1) = bsSysDefault
22110             Colors(2) = bsSysDefault: Colors(3) = bsSysDefault
22120             borderType = bsRaised
22130         Case "SUNKEN_SYSTEM_COLOR" ' sunken using system colors
22140             Colors(0) = bsSysDefault: Colors(1) = bsSysDefault
22150             Colors(2) = bsSysDefault: Colors(3) = bsSysDefault
22160             borderType = bsSunken
22170     End Select
          
22180     If clsBorders Is Nothing Then
22190         Set clsBorders = New cBorders
22200     End If
                    

22230         If MyControll.Name = SpecifiekeControllName Then
22240             If TypeName(MyControll) = "TextBox" Then
22250                 clsBorders.SetBorder MyControll.hwnd, borderType, ctTextBox, Colors(0), Colors(1), Colors(2), Colors(3)
22260             ElseIf TypeName(MyControll) = "ListBox" Or TypeName(MyControll) = "FileListBox" Then
22270                 clsBorders.SetBorder MyControll.hwnd, borderType, ctListBox, Colors(0), Colors(1), Colors(2), Colors(3)
22280             ElseIf TypeName(MyControll) = "ComboBox" Or TypeName(MyControll) = "DriveListBox" Or TypeName(MyControll) = "DTPicker" Then
22290                 clsBorders.SetBorder MyControll.hwnd, borderType, ctComboBox, Colors(0), Colors(1), Colors(2), Colors(3)
22300             ElseIf TypeName(MyControll) = "ImageCombo" Then
22310                 clsBorders.SetBorder MyControll.hwnd, borderType, ctImageCombo, Colors(0), Colors(1), Colors(2), Colors(3)
22320             ElseIf TypeName(MyControll) = "vbalImageList" Or TypeName(MyControll) = "Timer" Or TypeName(MyControll) = "VistaForm" Or TypeName(MyControll) = "Shape" Or TypeName(MyControll) = "Label" Or TypeName(MyControll) = "Image" Or TypeName(MyControll) = "ImageList" Or TypeName(MyControll) = "WebBrowser" Or TypeName(MyControll) = "vbalDTabControlX" Then
22330             Else
22340                     clsBorders.SetBorder MyControll.hwnd, borderType, , Colors(0), Colors(1), Colors(2), Colors(3)
22350             End If
22370         End If

              
End Sub

