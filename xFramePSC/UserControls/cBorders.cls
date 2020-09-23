VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' VB IDE Safe Subclassing provided by...
'*************************************************************************************************
'* cSelfSub - self-subclassing class template
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
'* v1.1 VirtualAlloc memory to prevent Data Execution Prevention faults on Win64......... 20060324
'* v1.2 Thunk redesigned to handle unsubclassing and memory release...................... 20060325
'* v1.3 Data array scrapped in favour of property accessors.............................. 20060405
'*************************************************************************************************

' While I assisted troubleshooting one of TerriTop's apps, I came up with one solution
' that didn't quite work for him, but he suggested I share it with PSC. So here it is.

' Purpose: Provide custom border colors for some common VB controls. The controls this
'   should work for are those that have static borders. In other words, during normal,
'   click, and enable events, the border does not change. VB Command Buttons are a
'   prime example of borders that do change, therefore, this wouldn't work for them.
'   Not all static border controls have been tested. The ones in this example have been
'   and I could include others by request, but the borders must be static, not dynamic.

' The complex part of this project was trying to determine how borders were drawn on
'   the controls. Some controls have actual border window styles (i.e., WS_Border,
'   WS_EX_ClientEdge, etc) and those can be drawn over during a WM_NCPaint message.
'   But others have their borders drawn on their client area or one of their children's
'   client area. Another exception is the ImageCombo which is an owner-drawn control.


Option Explicit

'-Selfsub declarations----------------------------------------------------------------------------
Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
End Enum

Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
Private Const CODE_LEN      As Long = 240                                   'Thunk length in bytes
Private Const WNDPROC_OFF   As Long = &H30                                  'WndProc execution offset
Private Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))    'Bytes to allocate per thunk, data + code + msg tables
Private Const PAGE_RWX      As Long = &H40                                  'Allocate executable memory
Private Const MEM_COMMIT    As Long = &H1000                                'Commit allocated memory
Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 1                                     'Shutdown flag data index
Private Const IDX_HWND      As Long = 2                                     'hWnd data index
Private Const IDX_EBMODE    As Long = 3                                     'EbMode data index
Private Const IDX_CWP       As Long = 4                                     'CallWindowProc data index
Private Const IDX_SWL       As Long = 5                                     'SetWindowsLong data index
Private Const IDX_FREE      As Long = 6                                     'VirtualFree data index
Private Const IDX_ME        As Long = 7                                     'Owner data index
Private Const IDX_WNDPROC   As Long = 8                                     'Original WndProc data index
Private Const IDX_CALLBACK  As Long = 9                                     'zWndProc data index
Private Const IDX_BTABLE    As Long = 10                                    'Before table data index
Private Const IDX_ATABLE    As Long = 11                                    'After table data index
Private Const IDX_EBX       As Long = 14                                    'Data code index

Private z_Base              As Long                                         'Data pointer base
Private z_Funk              As Collection                                   'hWnd/thunk-address collection
Private z_TblEnd            As Long                                         'End of the vTable
Private z_Code(29)          As Currency                                     'Thunk machine-code initialised here

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'-------------------------------------------------------------------------------------------------


Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As RECT) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetBkColor Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByRef lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long

Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOSENDCHANGING As Long = &H400
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_BORDER As Long = &H800000
Private Const WM_NCPAINT As Long = &H85
Private Const WM_PAINT As Long = &HF&
Private Const WM_DRAWITEM As Long = &H2B
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const SM_CYBORDER As Long = 6
Private Const SM_CYEDGE As Long = 46
Private Const COLOR_BTNHIGHLIGHT As Long = 20
Private Const COLOR_BTNSHADOW As Long = 16
Private Const COLOR_3DDKSHADOW As Long = 21
Private Const COLOR_3DLIGHT As Long = 22
Private Const COLOR_WINDOWFRAME As Long = 6
Private Const LBS_NOINTEGRALHEIGHT As Long = &H100&
Private Const CBS_NOINTEGRALHEIGHT As Long = &H400&
Private Const WM_CTLCOLOREDIT As Long = &H133
Private Const WM_CTLCOLORLISTBOX As Long = &H134
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Settings() As Long
Public Enum Specialcolors 'sneaky way of exposing public "constants" from a class
    bsSysDefault = -2
    bsAutoShade = -1
End Enum
Public Enum BorderStyleOptions
    [_bs_Max] = 3
    [_bsBackColor] = -3
    bsFlat1Color = 0    ' 1 color, 1 pixel border around control
    bsFlat2Color = 1    ' 2 color, 1 pixel border
    bsSunken = 2        ' 4 color, 2 pixel border
    bsRaised = 3        ' 4 color, 2 pixel border
'    bsNone = 4          ' no borders at all -- not coded/not sure it will be
End Enum
Public Enum vbControlType
    ctOther = 0         ' controls that have a true border (non-client area)
    ctComboBox = 1      ' some draw on client and some on non-client
    ctImageCombo = 2    ' owner-drawn by VB
    ctListBox = 3       ' control height can change when changing border styles
    ctTextBox = 4       ' when flat style, VB draws border on client area
End Enum

Private Borders As Collection

Public Sub SetBorder(ByVal hwnd As Long, BorderStyle As BorderStyleOptions, _
                    Optional ByVal ctrlType As vbControlType = ctOther, _
                    Optional ByVal Shadow As Long = vbButtonShadow, _
                    Optional ByVal DarkShadow As Long = bsAutoShade, _
                    Optional ByVal Highlight As Long = bsAutoShade, _
                    Optional ByVal LightShadow As Long = bsAutoShade)
          
          ' The control will be subclassed to allow custom borders. Therefore it is
          ' best to call ResetBorder on any control loaded with Controls.Add before
          ' that control is removed if you sent that added control here.
          
          ' ///// Border Styles \\\\\
          ' bsFlat1Color. Solid 1-pixel border, 1 color (i.e., flat).
          '       Uses Shadow only
          ' bsFlat2Color. Left/Top borders are 1 color, right/bottom are another
          '       Uses Shadow & Highlight only
          ' bsSunken. Left/Top outer border are Shadow, Right/Bottom outer are HighLight
          '           Left/Top inner border are DarkShadow, Right/Bottom inner are LightShadow
          ' bsRaised. Left/Top outer border are HighLight, Right/Bottom outer are DarkShadow
          '           Left/Top inner border are LightShadow, Right/Bottom outer are Shadow
          
          ' ///// colors \\\\\ vb system colors can be passed
          ' Shadow: 2nd darkest of 4 color borders; color for a single color border
          ' DarkShadow: the darkest of 4 color borders
          ' LightShadow: 2nd lightest of 4 color borders
          ' Highlight: lightest of 4 color borders
          ' Special values for the above 4 colors
          '   -1 = AutoShade. DarkShadow, LightShadow & Highlight are shades of Shadow
          '           DarkShadow = Shadow darkened to 15% from black
          '           LightShadow = Shadow lightened by 85% of its lightest value (white)
          '           Highlight = Shadow lightened by 100% (or vbWhite)
          '   -2 = System colors: vb3DDKShadow, vbButtonShadow, vb3DLight, vbHighlight respectively
          '   -3 & -4 (Reserved) are used by the class to fake single borders on combo boxes
          
          ' ///// Control Type \\\\\
          ' Some controls have their borders drawn by VB on the control's client area whereas
          '   others are drawn in the non-client area as expected. Think of a form with no
          '   borders but you want borders so you draw it on the form (non-client area).
          '   VB combo boxes are very much like that scenario. Therefore, the control type
          '   needs to be known in advance so the class can handle those special cases.
          '   There are more special cases too & those known are handled herein
          ' ctComboBox: use for comboboxes and drivecombo
          ' ctImageCombo: use for the image combobox
          ' ctListBox: use for listboxes and file listboxes
          ' ctTextBox: use for the textbox control
          ' ctOther: use for other controls like treeview, listview, progressbar, etc
          
          
          ' sanity checks first & don't allow user to pass a reserved color code
24010     If BorderStyle < bsFlat1Color Or BorderStyle > [_bs_Max] Then Exit Sub
24020     If DarkShadow = -3 Or DarkShadow = -4 Then DarkShadow = 0
24030     If LightShadow = -3 Or LightShadow = -4 Then LightShadow = 0
24040     If Shadow = -3 Or Shadow = -4 Then Shadow = 0
24050     If Highlight = -3 Or Highlight = -4 Then Highlight = 0
          
          Dim Index As Long, lFlags As Long, isSubclassed As Boolean
          Dim lExStyle As Long, lStyle As Long, lOldStyle As Long
          Dim cHwnd As Long, wRect As RECT, cRect As RECT
          Dim borderSize As Long, borderSizeNew As Long
          ' flags for special case control scenarios
          Dim bIntegralHT As Boolean, bFlatTextBox As Boolean, bHasBorder As Boolean
          
24060     On Error Resume Next
24070     If Borders Is Nothing Then          ' first time thru
24080         Set Borders = New Collection
24090         Index = 1
24100         ReDim Settings(1 To 8)
24110         Borders.Add Index, "h" & hwnd
24120     Else
24130         Index = Borders("h" & hwnd)     ' been here before, is this hWnd already ours?
24140         If Index = 0 Then               ' nope, set it up
24150             Index = UBound(Settings) + 1
24160             ReDim Preserve Settings(1 To Index + 7)
24170             Borders.Add Index, "h" & hwnd
24180             If Err Then Err.Clear
24190         Else
24200             isSubclassed = True         ' yep, use cached settings
24210         End If
24220     End If
          
          ' cache colors & style settings
24230     Settings(Index) = CLng(BorderStyle Or (ctrlType * &H100&))
24240     Settings(Index + 1) = Shadow
24250     Settings(Index + 2) = Highlight
24260     Settings(Index + 3) = DarkShadow
24270     Settings(Index + 4) = LightShadow
24280     Settings(Index + 7) = hwnd        ' needed when re-indexing if control is unsubclassed
24290     If isSubclassed Then
24300         lStyle = Settings(Index + 5)
24310         lExStyle = Settings(Index + 6)
24320     Else
24330         lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
24340         lStyle = GetWindowLong(hwnd, GWL_STYLE)
24350         Settings(Index + 5) = lStyle
24360         Settings(Index + 6) = lExStyle
24370     End If
          ' setwindowpos flags
24380     lFlags = SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOSIZE Or SWP_NOZORDER
          
          ' special case for combo boxes. We can't remove the borders on these, because
          ' they really aren't borders for the most part; they are drawn on the client
          ' ComboBox.Style=1 is an exception. It has borders come from its child's borders,
          ' but to make things a little bit easier we will treat all combos the same
24390     If ctrlType = ctComboBox Or ctrlType = ctImageCombo Then
24400         If BorderStyle = bsFlat2Color Or BorderStyle = bsFlat1Color Then
                  ' to fake a 1-pixel border, we will draw the inner level of borders
                  ' in the same backcolor as the control. Special flags are set for this.
                  ' Note: GetBkColor does not return the control's BackColor property.
                  ' These controls are sent a wm_ctlcoloredit or wm_ctlcolorlistbox message
                  ' forcing the parent object to provide the bacground color brush
24410             Settings(Index + 3) = [_bsBackColor]
24420             Settings(Index + 4) = [_bsBackColor] - 1
                  ' normally a 1 pixel border only requires 2 colors and only an outer border
                  ' however, we'll tweak so 4 colors are used & also an inner border
24430             Settings(Index) = CLng(bsSunken Or (ctrlType * &H100&))
24440             If BorderStyle = bsFlat1Color Then Settings(Index + 2) = Settings(Index + 1)
24450         End If
24460     Else
              ' combo styles are not modified, but color tweaked above to appear so
24470         Select Case BorderStyle
              Case bsFlat1Color, bsFlat2Color
24480             If ctrlType = ctTextBox Then
                      ' another exception: textbox with flat border style draws border on client
24490                 If (lExStyle And WS_EX_CLIENTEDGE) = 0 Then ' flat?
24500                     lStyle = lStyle And Not WS_BORDER
24510                     bFlatTextBox = True
24520                 Else
24530                     lStyle = lStyle Or WS_BORDER
24540                 End If
24550             Else
24560                 lStyle = lStyle Or WS_BORDER
24570             End If
24580             lExStyle = lExStyle And Not WS_EX_CLIENTEDGE
24590         Case Else
24600             If ctrlType = ctTextBox Then
                      ' another exception: textbox with flat border style draws border on client
24610                 If (Settings(Index + 6) And WS_EX_CLIENTEDGE) = 0 Then ' flat?
                          ' we need 2 borders, one drawn on client & we provide the 2nd one via WS_Border
24620                     lStyle = lStyle Or WS_BORDER
24630                     lExStyle = lExStyle And Not WS_EX_CLIENTEDGE
24640                     bFlatTextBox = True
24650                 Else
24660                     lStyle = lStyle And Not WS_BORDER
24670                     lExStyle = lExStyle Or WS_EX_CLIENTEDGE
24680                 End If
24690             Else
24700                 lStyle = lStyle And Not WS_BORDER
24710                 lExStyle = lExStyle Or WS_EX_CLIENTEDGE
24720             End If
24730         End Select
              
              ' now should a user be setting a listbox (IntegralHeight=True) border to
              ' raised/sunken when it was previously set to flat, simply changing the borderstyles
              ' will result in the listbox shrinking in size. We should account for that
              
              ' get current bordersize
24740         lOldStyle = GetWindowLong(hwnd, GWL_STYLE)
              ' get current bordersize
24750         If (lOldStyle And WS_BORDER) = WS_BORDER Then
24760             borderSize = GetSystemMetrics(SM_CYBORDER)
24770             bHasBorder = True
24780         End If
24790         If (GetWindowLong(hwnd, GWL_EXSTYLE) And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then
24800             borderSize = borderSize + GetSystemMetrics(SM_CYEDGE)
24810         Else
24820             bHasBorder = False
24830         End If
24840         GetWindowRect hwnd, wRect
24850         cRect = wRect
24860         OffsetRect wRect, -wRect.Left, -wRect.Top

24870         If ctrlType = ctListBox Then
24880             bIntegralHT = ((lOldStyle And LBS_NOINTEGRALHEIGHT) = 0)
                  
                  ' get adjusted bordersize
24890             If (lStyle And WS_BORDER) = WS_BORDER Then borderSizeNew = GetSystemMetrics(SM_CYBORDER)
24900             If (lExStyle And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then borderSizeNew = borderSizeNew + GetSystemMetrics(SM_CYEDGE)
24910             If bIntegralHT = True Then
24920                 If borderSizeNew > borderSize Then
                          ' adjust height as needed
24930                     wRect.Bottom = wRect.Bottom + (borderSizeNew - borderSize) * 2
24940                     lFlags = lFlags And Not SWP_NOSIZE
24950                 End If
24960             End If
24970         End If
24980         If bHasBorder Then
                  ' some controls (FileList, Treeview, ListView) have an extra border.
                  ' It is same color as the form's back color an is the far outer border.
                  ' We will be removing that but to ensure the control looks as though it
                  ' stays at the same X,Y coordinates after we remove it, we will shift
                  ' the control 1 pixel
24990             ScreenToClient GetParent(hwnd), cRect
25000             OffsetRect cRect, 1, 1
25010             lFlags = lFlags And Not SWP_NOMOVE
25020         Else
25030             cRect.Left = 0: cRect.Top = 0
25040         End If
25050         SetWindowLong hwnd, GWL_EXSTYLE, lExStyle
25060         SetWindowLong hwnd, GWL_STYLE, lStyle
25070     End If
          
25080     If Not isSubclassed Then    ' haven't subclassed this hWnd yet
25090         sc_Subclass hwnd
25100         If ctrlType = ctImageCombo Then
25110             sc_AddMsg hwnd, WM_DRAWITEM, MSG_AFTER ' ownerdrawn control
25120             sc_AddMsg hwnd, WM_CTLCOLOREDIT, MSG_BEFORE
25130         ElseIf ctrlType = ctComboBox Or bFlatTextBox = True Then
                  ' draw border on client
25140             sc_AddMsg hwnd, WM_PAINT, MSG_AFTER
25150             sc_AddMsg hwnd, WM_CTLCOLOREDIT, MSG_BEFORE
25160             sc_AddMsg hwnd, WM_CTLCOLORLISTBOX, MSG_BEFORE
25170         Else
25180             sc_AddMsg hwnd, WM_NCPAINT, MSG_AFTER ' draw border on non-client
25190         End If
25200     End If
          
          ' force any border changes now
25210     SetWindowPos hwnd, 0, cRect.Left, cRect.Top, wRect.Right, wRect.Bottom, lFlags
              
          ' force repaint. SetWindowPos seems not to do this for all controls
25220     InvalidateRect hwnd, ByVal 0&, True
          
End Sub

Public Sub ReSetBorder(ByVal hwnd As Long, Optional Refresh As Boolean = True)

          ' function removes an custom borders & resets to system defaults
          ' For extra info on comments, see SetBorder. This is basically the
          ' opposite logic that was applied in SetBorder
          
25230     If Borders Is Nothing Then Exit Sub
          
          Dim Index As Long
          Dim wRect As RECT, cRect As RECT
          Dim lStyle As Long, lStyleEX As Long
          Dim borderSize As Long, borderSizeNew As Long
          Dim lFlags As Long
          Dim ctrlType As Long
          Dim bIntegralHT As Boolean, bHasBorder As Boolean
          
25240     On Error Resume Next
25250     Index = Borders("h" & hwnd) ' did we subclass this one?
25260     If Index = 0 Then
25270         If Err Then Err.Clear   ' nope, nothing to do
25280         Exit Sub
25290     End If
25300     On Error GoTo 0

25310     sc_UnSubclass hwnd  ' unsubclass it
          ' these will be the flags used for SetWindowPos
25320     lFlags = SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOSIZE
          
25330     ctrlType = ((Settings(Index) \ &HFF) And &HFF)
          
          ' comboboxes go through un-modified, don't need to run extra code for those
25340     If Not (ctrlType = ctComboBox Or ctrlType = ctImageCombo) Then
              ' listboxes and comboboxes could shrink in height if there IntegralHeight
              ' property is set to True. This is because leaving the control the same
              ' size but adding extra border pixels to it would reduce the total client
              ' area and thus force the control to resize itself smaller. Account for this.
              ' Additionally, other controls like FileList, Treeview, ListView had
              ' and extra border level we compensated for by shifting its X,Y position.
              ' Re-shift it if needed.
              
25350         GetWindowRect hwnd, wRect
25360         cRect = wRect
25370         OffsetRect wRect, -wRect.Left, -wRect.Top
          
              ' calculate the original bordersize
25380         If (Settings(Index + 5) And WS_BORDER) = WS_BORDER Then
25390             borderSizeNew = GetSystemMetrics(SM_CYBORDER)
25400             bHasBorder = True
25410         End If
25420         If (Settings(Index + 6) And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then
25430             borderSizeNew = borderSizeNew + GetSystemMetrics(SM_CYEDGE)
25440         Else
25450             bHasBorder = False
25460         End If
          
25470         If ctrlType = ctListBox Then
25480             bIntegralHT = ((Settings(Index + 5) And LBS_NOINTEGRALHEIGHT) = 0)
                  
                  ' calculate the current vertical bordersize
25490             lStyle = GetWindowLong(hwnd, GWL_STYLE)
25500             lStyleEX = GetWindowLong(hwnd, GWL_EXSTYLE)
25510             If (lStyle And WS_BORDER) = WS_BORDER Then borderSize = GetSystemMetrics(SM_CYBORDER)
25520             If (lStyleEX And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE Then borderSize = borderSize + GetSystemMetrics(SM_CYEDGE)
                  
                  ' now calculate any difference in height
25530             If bIntegralHT Then ' our control doesn't have IntegralHeight=False
25540                 If borderSizeNew > borderSize Then
                          ' need to resize the control when we re-apply the original borders
                          ' add the extra height needed so the client area can remain same size
25550                     wRect.Bottom = wRect.Bottom + (borderSizeNew - borderSize) * 2
25560                     lFlags = lFlags And Not SWP_NOSIZE ' we want to resize the control
25570                 End If
25580             End If
25590         End If
25600         If bHasBorder Then
25610             ScreenToClient GetParent(hwnd), cRect
25620             OffsetRect cRect, -1, -1
25630             lFlags = lFlags And Not SWP_NOMOVE
25640         Else
25650             cRect.Left = 0: cRect.Top = 0
25660         End If
          
25670     End If
          ' set the new border styles as needed, then force a resize/redraw
25680     SetWindowLong hwnd, GWL_STYLE, Settings(Index + 5)
25690     SetWindowLong hwnd, GWL_EXSTYLE, Settings(Index + 6)
25700     SetWindowPos hwnd, 0, cRect.Left, cRect.Top, wRect.Right, wRect.Bottom, lFlags

          ' now clean up our collection
25710     If Borders.Count = 1 Then
25720         Set Borders = Nothing
25730         Erase Settings
25740     Else
25750         If Index + 8 < UBound(Settings) Then
                  ' move the last item in our collection over the unsubclassed item
25760             lFlags = Borders.Item(Borders.Count)
25770             RtlMoveMemory VarPtr(Settings(Index)), VarPtr(Settings(lFlags)), 32&
                  ' now rebuild the collection item
25780             Borders.Remove "h" & Settings(Index + 7)
25790             Borders.Add Index, "h" & Settings(Index + 7), "h" & hwnd
25800         End If
25810         Borders.Remove "h" & hwnd
25820         ReDim Preserve Settings(1 To UBound(Settings) - 8)
25830     End If

          ' force repaint. SetWindowPos seems not to do this for all controls
25840     If Refresh Then InvalidateRect hwnd, ByVal 0&, True

End Sub

Private Sub RGBToHSL(inColor As Long, Hue As Single, _
                Luminance As Single, Saturation As Single)

      ' various versions on the net. This one needed some tweaking but will accurately
      ' return the HSL +/- 2 on a single RGB value compared to the color picker.

      ' The HSL values are ratios btwn 0:100 & can be applied to any maximum value with a
      ' minimum value of zero. For example: Luminance is a ratio btwn 0:100 and can be
      ' multiplied against 240 to return a minimum/maximum Luminance of 0 to 240 similar
      ' to how the color picker displays it.  It can also be applied against 256, the
      ' maximum number of shades of light.
          Dim maxRGB As Single, minRGB As Single, hueDelta As Single
          Dim Red As Single, Green As Single, Blue As Single
          Dim bRGB(0 To 3) As Byte
          
          ' get long color into RGB bytes & convert to a ratio
25850     bRGB(0) = inColor Mod &H100
25860     bRGB(1) = (inColor \ &H100) Mod &H100
25870     bRGB(2) = (inColor \ &H10000) Mod &H100
          
25880     Red = bRGB(0) / 255
25890     Green = bRGB(1) / 255
25900     Blue = bRGB(2) / 255
          
          ' simple little routine to get the largest/smallest of the RGB bytes
25910     If Red > Green Then
25920         minRGB = Green
25930         maxRGB = Red
25940     Else
25950         minRGB = Red
25960         maxRGB = Green
25970     End If
25980     If Blue > maxRGB Then
25990         maxRGB = Blue
26000     Else
26010         If Blue < minRGB Then minRGB = Blue
26020     End If

          ' calculate Luminance or Light as some know it
26030     Luminance = (maxRGB + minRGB) / 2

          ' calculate Saturation
26040     If maxRGB = minRGB Then
              ' grayscale, only Luminance is used
26050         Saturation = 0#
26060         Hue = 0#
26070     Else
26080         If Luminance <= 0.5 Then
26090            Saturation = (maxRGB - minRGB) / (maxRGB + minRGB)
26100         Else
26110            Saturation = (maxRGB - minRGB) / (2 - maxRGB - minRGB)
26120         End If
              ' now for the Hue; much more complicated
              
              ' get the difference between the largest/smallest color range
26130         hueDelta = maxRGB - minRGB
              
              ' this is a bit complicated, but Hue is calculated here as a relation
              ' to a 360 degree circle.
              
              ' The IF's below determine between what 2 of the 3 primary colors this color falls
              ' in order to calculate the Hue
26140         If Red = maxRGB Then
                 ' falls betweeen green & blue
26150             Hue = (Green - Blue) / hueDelta
26160         ElseIf Green = maxRGB Then
                 ' falls between blue and red
26170             Hue = 2# + (Blue - Red) / hueDelta
26180         Else    'Blue = maxRGB
                 ' falls between red & green
26190             Hue = 4# + (Red - Green) / hueDelta
26200         End If
              ' convert to degrees
26210         Hue = Hue * 60
              ' check for reverse rotation & adjust by adding 360
26220         If Hue < 0# Then Hue = Hue + 360#
              ' now we can convert Hue & Saturation to a simple ratio before returning
26230         Hue = Hue / 360# * 100#
26240         Saturation = Saturation * 100#
26250     End If
          ' convert to proper ratio before returning
26260     Luminance = Luminance * 100#
          ' done: return result.
End Sub

Private Function HSLToRGB(ByVal Hue As Single, _
        ByVal Luminance As Single, ByVal Saturation As Single) As Long

      ' various versions on the net. This one needed some tweaking but will accurately
      ' return the RGB +/- 2 on a single HSL value when compared to the color picker.

      ' When the parameters are passed as percentages calculated from the RGBtoHSL
      ' routine, this routine will convert HSL to RGB with 100% accuracy.

      ' function will return the RGB as a long:
      ' 1st byte=Red, 2nd byte=Blue , 3rd byte=Green, 4th byte is unused

          Dim Red As Single, Green As Single, Blue As Single
          Dim minRGB As Single, maxRGB As Single
          Dim bRGB(0 To 3) As Byte
          
          ' ensure passed parameters are within 0:100 range
26270         If Saturation < 0 Then
26280             Saturation = 0
26290         ElseIf Saturation > 100 Then
26300             Saturation = 100
26310         End If
26320         If Luminance < 0 Then
26330             Luminance = 0
26340         ElseIf Luminance > 100 Then
26350             Luminance = 100
26360         End If
26370         If Hue < 0 Then
26380             Hue = 0
26390         ElseIf Hue > 100 Then
26400             Hue = 100
26410         End If
          ' now convert ratios to a 0:1 range
26420     Saturation = Saturation / 100
26430     Luminance = Luminance / 100

26440     If Saturation = 0 Then
             ' grayscale
26450        Red = Luminance
26460        Green = Luminance
26470        Blue = Luminance
          
26480     Else
26490         Hue = Hue / 100
              ' calculate the least RGB value
26500         If Luminance <= 0.5 Then
26510             minRGB = Luminance * (1 - Saturation)
26520         Else
26530             minRGB = Luminance - Saturation * (1 - Luminance)
26540         End If
              ' calculate the greatest RGB value
26550         maxRGB = 2 * Luminance - minRGB
                
              ' Hue is related to a circle or quadrants in this case
              ' We will use ratio on 360 degree to determine reverse rotation
26560         Hue = Hue * 360
              ' Any angle calculated > 299.999 is a reverse rotation and the ratio
              ' is actually a negative number. We need to get this negative value.
              ' Note: This section checked by running all possible RGB combinations
              ' thru the code to ensure RGB > HSL > RGB returned original color.
26570         If Hue >= 300 Then Hue = Hue - 360
              ' now to get the final 0:1 ratio...
26580         Hue = Hue / 60
              
              ' this is the reverse of the RGBtoHSL routine & is where the magic happens
26590         If (Hue < 1) Then
26600            Red = maxRGB
26610            If (Hue < 0) Then
26620               Green = minRGB
26630               Blue = Green - Hue * (maxRGB - minRGB)
26640            Else
26650               Blue = minRGB
26660               Green = Hue * (maxRGB - minRGB) + Blue
26670            End If
26680         ElseIf (Hue < 3) Then
26690            Green = maxRGB
26700            If (Hue < 2) Then
26710               Blue = minRGB
26720               Red = Blue - (Hue - 2) * (maxRGB - minRGB)
26730            Else
26740               Red = minRGB
26750               Blue = (Hue - 2) * (maxRGB - minRGB) + Red
26760            End If
26770         Else
26780            Blue = maxRGB
26790            If (Hue < 4) Then
26800               Red = minRGB
26810               Green = Red - (Hue - 4) * (maxRGB - minRGB)
26820            Else
26830               Green = minRGB
26840               Red = (Hue - 4) * (maxRGB - minRGB) + Green
26850            End If
          
26860         End If
                   
26870     End If
          
          ' convert the ratios to RGB & return the result
26880     bRGB(0) = CByte(Red * 255): bRGB(1) = Green * 255: bRGB(2) = Blue * 255
26890     HSLToRGB = RGB(bRGB(0), bRGB(1), bRGB(2))

End Function

Private Function ConvertColor(tColor As Long) As Long

      ' Converts VB color constants to real color values
26900     If tColor < 0 Then
26910         ConvertColor = GetSysColor(tColor And &HFF&)
26920     Else
26930         ConvertColor = tColor
26940     End If
End Function


Private Sub Class_Terminate()
        'Terminate all subclassing
        
        Dim i As Integer
26950   If Not Borders Is Nothing Then
26960     For i = Borders.Count To 1 Step -1
26970         ReSetBorder Settings(Borders.Item(i) + 7), False
26980     Next
26990   End If
27000   sc_Terminate
27010   Set Borders = Nothing
27020   Erase Settings
End Sub

'-SelfSub code------------------------------------------------------------------------------------
Private Function sc_Subclass(ByVal lng_hWnd As Long) As Boolean             'Subclass the specified window handle
        Dim nAddr As Long
        
27030   If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
27040     zError "sc_Subclass", "Invalid window handle"
27050   End If

27060   If z_Funk Is Nothing Then                                                 'If first time
27070     nAddr = zAddressOf(1)                                                   'Get the address of the final private method
          
27080     If nAddr = 0 Then
27090       zError "sc_Subclass", "Callback method not found"
27100     End If
          
27110     Set z_Funk = New Collection                                             'Create the hWnd/thunk-memory-address collection
          
27120     z_Code(6) = -490736517001394.5807@: z_Code(7) = 484417356483292.94@: z_Code(8) = -171798741966746.6996@: z_Code(9) = 843649688964536.7412@: z_Code(10) = -330085705188364.0817@: z_Code(11) = 41621208.9739@: z_Code(12) = -900372920033759.9903@: z_Code(13) = 291516653989344.1016@: z_Code(14) = -621553923181.6984@: z_Code(15) = 291551690021556.6453@: z_Code(16) = 28798458374890.8543@: z_Code(17) = 86444073845629.4399@: z_Code(18) = 636540268579660.4789@: z_Code(19) = 60911183420250.2143@: z_Code(20) = 846934495644380.8767@: z_Code(21) = 14073829823.4668@: z_Code(22) = 501055845239149.5051@: z_Code(23) = 175724720056981.1236@: z_Code(24) = 75457451135513.7931@: z_Code(25) = -576850389355798.3357@: z_Code(26) = 146298060653075.5445@: z_Code(27) = 850256350680294.7583@: z_Code(28) = -4888724176660.092@: z_Code(29) = 21456079546.6867@
          
27130     z_Base = VarPtr(z_Code(0))                                              'Map the address of z_Code()'s first element to the zData() array
27140     zData(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                           'Store the EbMode function address in the thunk data
27150     zData(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                   'Store CallWindowProc function address in the thunk data
27160     zData(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                    'Store the SetWindowLong function address in the thunk data
27170     zData(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                    'Store the VirtualFree function address in the thunk data
27180     zData(IDX_ME) = ObjPtr(Me)                                              'Store my object address in the thunk data
27190     zData(IDX_CALLBACK) = nAddr                                             'Store the zWndProc address in the thunk data
27200   End If
        
27210   z_Base = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                   'Allocate executable memory
27220   RtlMoveMemory z_Base, VarPtr(z_Code(0)), CODE_LEN                         'Copy the thunk to the allocated memory

27230   On Error GoTo Catch                                                       'Catch double subclassing
27240     z_Funk.Add z_Base, "h" & lng_hWnd                                       'Add the hWnd/thunk-address to the collection
27250   On Error GoTo 0

27260   zData(IDX_EBX) = z_Base                                                   'Patch the data address
27270   zData(IDX_HWND) = lng_hWnd                                                'Store the window handle in the thunk data
27280   zData(IDX_BTABLE) = z_Base + CODE_LEN                                     'Store the address of the before table in the thunk data
27290   zData(IDX_ATABLE) = zData(IDX_BTABLE) + ((MSG_ENTRIES + 1) * 4)           'Store the address of the after table in the thunk data
27300   zData(IDX_WNDPROC) = _
                      SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_Base + WNDPROC_OFF) 'Set the new WndProc and store the original WndProc in the thunk data
27310   sc_Subclass = True                                                        'Indicate success
27320   Exit Function                                                             'Exit

Catch:
27330   zError "sc_Subclass", "Window handle is already subclassed"
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
        Dim i     As Long
        Dim nAddr As Long

27340   If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
27350     zError "sc_UnSubAll", "Subclassing hasn't been started", False
27360   Else
27370     With z_Funk
27380       For i = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
27390         nAddr = .Item(i)                                                    'Map zData() to the hWnd thunk address
27400         If IsBadCodePtr(nAddr) = 0 Then                                     'Ensure that the thunk hasn't already freed itself
27410           z_Base = nAddr                                                    'Map the thunk memory to the zData() array
27420           sc_UnSubclass zData(IDX_HWND)                                     'UnSubclass
27430         End If
27440       Next i                                                                'Next member of the collection
27450     End With
          
27460     Set z_Funk = Nothing                                                    'Destroy the hWnd/thunk-address collection
27470   End If
End Sub

'UnSubclass the specified window handle
Public Sub sc_UnSubclass(ByVal lng_hWnd As Long)
27480   If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
27490     zError "UnSubclass", "Subclassing hasn't been started", False
27500   Else
27510     zDelMsg lng_hWnd, ALL_MESSAGES, IDX_BTABLE                              'Delete all before messages
27520     zDelMsg lng_hWnd, ALL_MESSAGES, IDX_ATABLE                              'Delete all after messages
27530     z_Base = zMap_hWnd(lng_hWnd)                                            'Map the thunk memory to the zData() array
27540     zData(IDX_SHUTDOWN) = -1                                                'Set the shutdown indicator
27550     z_Funk.Remove "h" & lng_hWnd                                            'Remove the specified window handle from the collection
27560   End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
27570   If When And MSG_BEFORE Then                                               'If the message is to be added to the before original WndProc table...
27580     zAddMsg lng_hWnd, uMsg, IDX_BTABLE                                      'Add the message to the before table
27590   End If

27600   If When And MSG_AFTER Then                                                'If message is to be added to the after original WndProc table...
27610     zAddMsg lng_hWnd, uMsg, IDX_ATABLE                                      'Add the message to the after table
27620   End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
27630   If When And MSG_BEFORE Then                                               'If the message is to be deleted from the before original WndProc table...
27640     zDelMsg lng_hWnd, uMsg, IDX_BTABLE                                      'Delete the message from the before table
27650   End If

27660   If When And MSG_AFTER Then                                                'If the message is to be deleted from the after original WndProc table...
27670     zDelMsg lng_hWnd, uMsg, IDX_ATABLE                                      'Delete the message from the after table
27680   End If
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
27690   z_Base = zMap_hWnd(lng_hWnd)                                              'Map zData() to the thunk of the specified window handle
27700   sc_CallOrigWndProc = _
              CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
End Function

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal nTable As Long)
        Dim nCount As Long                                                        'Table entry count
        Dim i      As Long                                                        'Loop index

27710   z_Base = zMap_hWnd(lng_hWnd)                                              'Map zData() to the thunk of the specified window handle
27720   z_Base = zData(nTable)                                                    'Map zData() to the table address

27730   If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
27740     nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
27750   Else
27760     nCount = zData(0)                                                       'Get the current table entry count

27770     If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
27780       zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values", False
27790       Exit Sub
27800     End If

27810     For i = 1 To nCount                                                     'Loop through the table entries
27820       If zData(i) = 0 Then                                                  'If the element is free...
27830         zData(i) = uMsg                                                     'Use this element
27840         Exit Sub                                                            'Bail
27850       ElseIf zData(i) = uMsg Then                                           'If the message is already in the table...
27860         Exit Sub                                                            'Bail
27870       End If
27880     Next i                                                                  'Next message table entry

27890     nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
27900     zData(nCount) = uMsg                                                    'Store the message in the appended table entry
27910   End If

27920   zData(0) = nCount                                                         'Store the new table entry count
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal nTable As Long)
        Dim nCount As Long                                                        'Table entry count
        Dim i      As Long                                                        'Loop index

27930   z_Base = zMap_hWnd(lng_hWnd)                                              'Map zData() to the thunk of the specified window handle
27940   z_Base = zData(nTable)                                                    'Map zData() to the table address

27950   If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
27960     zData(0) = 0                                                            'Zero the table entry count
27970   Else
27980     nCount = zData(0)                                                       'Get the table entry count
          
27990     For i = 1 To nCount                                                     'Loop through the table entries
28000       If zData(i) = uMsg Then                                               'If the message is found...
28010         zData(i) = 0                                                        'Null the msg value -- also frees the element for re-use
28020         Exit Sub                                                            'Exit
28030       End If
28040     Next i                                                                  'Next message table entry
          
28050     zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table", False
28060   End If
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String, Optional ByVal bEnd As Boolean = True)
28070   App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
        
28080   MsgBox sMsg & ".", IIf(bEnd, vbCritical, vbExclamation) + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
        
28090   If bEnd Then
28100     End
28110   End If
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
28120   zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
28130   Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
28140   If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
28150     zError "z_Base = zMap_hWnd", "Subclassing hasn't been started", True
28160   Else
28170     On Error GoTo Catch                                                     'Catch unsubclassed window handles
28180     zMap_hWnd = z_Funk("h" & lng_hWnd)                                      'Get the thunk address
28190     z_Base = zMap_hWnd                                                      'Map zData() to the thunk address
28200   End If
        
28210   Exit Function                                                             'Exit returning the thunk address

Catch:
28220   zError "z_Base = zMap_hWnd", "Window handle isn't subclassed"
End Function

'Return the address of the specified ordinal private method, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(Optional ByVal nOrdinal As Long = 1) As Long
        Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
        Dim bVal  As Byte
        Dim nAddr As Long                                                         'Address of the vTable
        Dim i     As Long                                                         'Loop index
        Dim J     As Long                                                         'Loop limit
        
28230   If z_TblEnd = 0 Then                                                      'First time through...
28240     RtlMoveMemory VarPtr(nAddr), ObjPtr(Me), 4                              'Get the address of this object instance
        
28250     If Not zProbe(nAddr + &H1C, i, bSub) Then                               'Probe for a Class method
28260       If Not zProbe(nAddr + &H6F8, i, bSub) Then                            'Probe for a Form method
28270         If Not zProbe(nAddr + &H7A4, i, bSub) Then                          'Probe for a UserControl method
28280           Exit Function                                                     'Bail...
28290         End If
28300       End If
28310     End If
          
28320     i = i + 4                                                               'Bump to the next entry
28330     J = i + 1024                                                            'Set a reasonable limit, scan 256 vTable entries
          
28340     Do While i < J
28350       RtlMoveMemory VarPtr(nAddr), i, 4                                     'Get the address stored in this vTable entry
            
28360       If IsBadCodePtr(nAddr) Then                                           'Is the entry an invalid code address?
28370         z_TblEnd = i                                                        'Cache the vTable end-point
28380         GoTo found                                                          'Bad method signature, quit loop
28390       End If
        
28400       RtlMoveMemory VarPtr(bVal), nAddr, 1                                  'Get the byte pointed to by the vTable entry
28410       If bVal <> bSub Then                                                  'If the byte doesn't match the expected value...
28420         z_TblEnd = i                                                        'Cache the vTable end-point
28430         GoTo found                                                          'Bad method signature, quit loop
28440       End If
            
28450       i = i + 4                                                             'Next vTable entry
28460     Loop
          
28470     Exit Function                                                           'Final method not found
28480   End If
        
found:
28490   RtlMoveMemory VarPtr(zAddressOf), z_TblEnd - (nOrdinal * 4), 4            'Return the specified vTable entry address
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
        Dim bVal    As Byte
        Dim nAddr   As Long
        Dim nLimit  As Long
        Dim nEntry  As Long
        
28500   nAddr = nStart                                                            'Start address
28510   nLimit = nAddr + 32                                                       'Probe eight entries
        
28520   Do While nAddr < nLimit                                                   'While we've not reached our probe depth
28530     RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
          
28540     If nEntry <> 0 Then                                                     'If not an implemented interface
28550       RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
28560       If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
28570         nMethod = nAddr                                                     'Store the vTable entry
28580         bSub = bVal                                                         'Store the found method signature
28590         zProbe = True                                                       'Indicate success
28600         Exit Function                                                       'Return
28610       End If
28620     End If
          
28630     nAddr = nAddr + 4                                                       'Next vTable entry
28640   Loop
End Function

Private Property Get zData(ByVal nIndex As Long) As Long
28650   RtlMoveMemory VarPtr(zData), z_Base + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
28660   RtlMoveMemory z_Base + (nIndex * 4), VarPtr(nValue), 4
End Property

'-Subclass callback: must be private and the last method in the source file-----------------------
Private Sub zWndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
      '*************************************************************************************************
      '* bBefore  - Indicates whether the callback is before or after the original WndProc. Usually you
      '*            will know unless the callback for the uMsg value is specified as MSG_BEFORE_AFTER
      '*            (both before and after the original WndProc).
      '* bHandled - In a before original WndProc callback, setting bHandled to True will prevent the
      '*            message being passed to the original WndProc and (if set to do so) the after
      '*            original WndProc callback.
      '* lReturn  - WndProc return value. Set as per the MSDN documentation for the message value,
      '*            and/or, in an after the original WndProc callback, act on the return value as set
      '*            by the original WndProc.
      '* hWnd     - Window handle.
      '* uMsg     - Message value.
      '* wParam   - Message related data.
      '* lParam   - Message related data.
      '*************************************************************************************************
        
          Dim hBrushPen As Long
          Dim wRect As RECT
          Dim hDC As Long, cHwnd As Long, cHDC As Long
          Dim HLS(0 To 3) As Single
          Dim Color As Long
          Dim colIndex As Long
          Dim i As Integer, nrSteps As Integer
          
28670     colIndex = Borders("h" & hwnd)  ' index into the Settings array
28680     GetWindowRect hwnd, wRect       ' get window position/size
28690     OffsetRect wRect, -wRect.Left, -wRect.Top ' offset to 0,0
          
28700     Select Case uMsg
          
          Case WM_CTLCOLOREDIT, WM_CTLCOLORLISTBOX
              ' sent by combo boxes being subclassed
              Dim lgBrush As LOGBRUSH
28710         If Settings(colIndex + 4) = [_bsBackColor] - 1 Then
                  ' 1-pixel border, but 2nd level border is drawn by VB
                  ' We will overdraw it with back color. These controls send
                  ' parent a request for a colored brush to use for background
                  ' Intercept the brush and get its color
28720             hBrushPen = sc_CallOrigWndProc(hwnd, uMsg, wParam, lParam)
28730             GetGDIObject hBrushPen, Len(lgBrush), lgBrush
28740             Settings(colIndex + 3) = lgBrush.lbColor
                  ' now draw the 2nd level border
28750             hDC = GetWindowDC(hwnd)
28760             InflateRect wRect, -1, -1
28770             FrameRect hDC, wRect, hBrushPen
28780             bHandled = True
28790             lReturn = hBrushPen
28800         End If
          
28810     Case Else ' WM_DrawItem, WM_Pain, WM_NCPaint depending on type of control subclassed
28820         If uMsg = WM_DRAWITEM Then hwnd = FindWindowEx(hwnd, 0, vbNullString, vbNullString)
              '^^ image combo borders are custom drawn on child, get that child window
28830         hDC = GetWindowDC(hwnd)
              
28840         If (Settings(colIndex) And &HFF) = bsFlat1Color Then
                  ' single color, 1 pixel borders (non-combos). Simply use FrameRect
28850             If Settings(colIndex + 1) = bsSysDefault Then
28860                 Color = GetSysColor(COLOR_WINDOWFRAME)
28870             ElseIf Settings(colIndex + 1) = bsAutoShade Then
28880                 Color = GetSysColor(COLOR_WINDOWFRAME)
28890             Else
28900                 Color = ConvertColor(Settings(colIndex + 1))
28910             End If
28920             hBrushPen = CreateSolidBrush(Color)
28930             FrameRect hDC, wRect, hBrushPen
28940             DeleteObject hBrushPen
28950         Else
                  ' determine number of passes. 2 for 1-pixel border, 4 for 2-pixels
28960             If (Settings(colIndex) And &HFF) = bsFlat2Color Then nrSteps = 2 Else nrSteps = 4
28970             Select Case Settings(colIndex + 1)
                      ' get base color (Shadow) and convert it to HSL
                      Case bsAutoShade
28980                     Color = HSLToRGB(HLS(0), 100, HLS(2))
28990                 Case bsSysDefault, [_bsBackColor]
29000                     Color = GetSysColor(COLOR_BTNSHADOW)
29010                 Case Else
29020                     Color = ConvertColor(Settings(colIndex + 1))
29030             End Select
29040             RGBToHSL Color, HLS(0), HLS(1), HLS(2)
                  
29050             For i = 1 To nrSteps
                      ' select appropriate color, create pen & draw
29060                 If (Settings(colIndex) And &HFF) = bsRaised Then
                          ' raised borders use different color scheme
29070                     Select Case i
                          Case 1
29080                         Select Case Settings(colIndex + 2)
                              Case bsAutoShade
29090                             Color = HSLToRGB(HLS(0), 100, HLS(2))
29100                         Case bsSysDefault, [_bsBackColor]
29110                             Color = GetSysColor(COLOR_BTNHIGHLIGHT)
29120                         Case Else
29130                             Color = ConvertColor(Settings(colIndex + 2))
29140                         End Select
29150                     Case 2
29160                         Select Case Settings(colIndex + 3)
                              Case bsAutoShade
29170                             If HLS(1) < 20 Then
29180                                 Color = vbBlack
29190                             Else
29200                                 Color = HSLToRGB(HLS(0), 15, HLS(2))
29210                             End If
29220                         Case bsSysDefault
29230                             Color = GetSysColor(COLOR_3DDKSHADOW)
29240                         Case [_bsBackColor]
29250                             Color = GetBkColor(hDC)
29260                         Case Else
29270                             Color = ConvertColor(Settings(colIndex + 3))
29280                         End Select
29290                     Case 3
29300                         Select Case Settings(colIndex + 4)
                              Case bsAutoShade
29310                             Color = HSLToRGB(HLS(0), 85, HLS(2))
29320                         Case bsSysDefault
29330                             Color = GetSysColor(COLOR_3DLIGHT)
29340                         Case [_bsBackColor]
29350                             Color = GetBkColor(hDC)
29360                         Case [_bsBackColor] - 1
29370                             If Settings(colIndex + 3) = [_bsBackColor] Then
29380                                 Color = GetBkColor(hDC)
29390                             Else
29400                                 Color = Settings(colIndex + 3)
29410                             End If
29420                         Case Else
29430                             Color = ConvertColor(Settings(colIndex + 4))
29440                         End Select
29450                     Case 4
29460                         Select Case Settings(colIndex + 1)
                              Case bsAutoShade
29470                             Color = HSLToRGB(HLS(0), 100, HLS(2))
29480                         Case bsSysDefault, [_bsBackColor]
29490                             Color = GetSysColor(COLOR_BTNSHADOW)
29500                         Case Else
29510                             Color = ConvertColor(Settings(colIndex + 1))
29520                         End Select
29530                     End Select
29540                 Else                ' sunken borders
29550                     Select Case i
                          Case 1
29560                     Case 2
29570                         Select Case Settings(colIndex + i)
                              Case bsAutoShade
29580                             Color = HSLToRGB(HLS(0), 100, HLS(2))
29590                         Case bsSysDefault, [_bsBackColor]
29600                             Color = GetSysColor(COLOR_BTNHIGHLIGHT)
29610                         Case Else
29620                             Color = ConvertColor(Settings(colIndex + i))
29630                         End Select
29640                     Case 3
29650                         Select Case Settings(colIndex + i)
                              Case bsAutoShade
29660                             Color = HSLToRGB(HLS(0), HLS(1) - 15, HLS(2))
29670                         Case bsSysDefault
29680                             Color = GetSysColor(COLOR_3DDKSHADOW)
29690                         Case [_bsBackColor]
29700                             Color = GetBkColor(hDC)
29710                         Case Else
29720                             Color = ConvertColor(Settings(colIndex + i))
29730                         End Select
29740                     Case 4
29750                         Select Case Settings(colIndex + i)
                              Case bsAutoShade
29760                             Color = HSLToRGB(HLS(0), 85, HLS(2))
29770                         Case bsSysDefault
29780                             Color = GetSysColor(COLOR_3DLIGHT)
29790                         Case [_bsBackColor]
29800                             Color = GetBkColor(hDC)
29810                         Case [_bsBackColor] - 1
29820                         Case Else
29830                             Color = ConvertColor(Settings(colIndex + i))
29840                         End Select
29850                     End Select
29860                 End If
                      
29870                 hBrushPen = SelectObject(hDC, CreatePen(0, 1, Color))
29880                 Select Case i
                      Case 1
29890                     MoveToEx hDC, 0, wRect.Bottom - 2, ByVal 0&
29900                     LineTo hDC, 0, 0
29910                     LineTo hDC, wRect.Right - 1, 0
29920                 Case 2
29930                     LineTo hDC, wRect.Right - 1, 0
29940                     LineTo hDC, wRect.Right - 1, wRect.Bottom - 1
29950                     LineTo hDC, -1, wRect.Bottom - 1
29960                 Case 3
29970                     MoveToEx hDC, 1, wRect.Bottom - 3, ByVal 0&
29980                     LineTo hDC, 1, 1
29990                     LineTo hDC, wRect.Right - 2, 1
30000                 Case 4
30010                     LineTo hDC, wRect.Right - 2, 1
30020                     LineTo hDC, wRect.Right - 2, wRect.Bottom - 2
30030                     LineTo hDC, 0, wRect.Bottom - 2
30040                 End Select
30050                 DeleteObject SelectObject(hDC, hBrushPen)
30060             Next
30070         End If
30080     End Select
          ' release DC as needed
30090     If Not hDC = 0 Then ReleaseDC hwnd, hDC
End Sub


