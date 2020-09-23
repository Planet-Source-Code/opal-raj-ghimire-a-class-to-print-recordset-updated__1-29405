VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RsPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'#############################################################################
' #
' #                                              Recordset Printing Class-2 (ADO DAO)
' #                                                              Updated on 02.02.2002
' #                                      ---------------------------------------------------------------------
' #                                        Author: Opal Raj Ghimire, buna48@hotmail.com
' #                                                    http://geocities.com/opalraj/vb
' #                                      ---------------------------------------------------------------------
' #         You are free to use this class in your projects but read readme.txt.
' #         If you use it, a notification is appreciated. Don't forget to credit me at appropriate place
' #         This class comes with absolutely NO WARranty ! Use it at your own risk !
' #         May posses errors !! Inform me if you find one I will try to fix it.
' #
'############################################################################
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
  
Private BorX As Long, BorY As Long
Private BorDX As Long, BorDY As Long    ' Bor=Boarder
Private bFieldPrinted As Boolean               ' Fields Names printed or not
Private RecSource As Object                       ' Rocordset object
Private Cols() As String                                 ' Col width
Private Aligns() As String                              ' Alignments
Private bFPage As Boolean                         ' Final page flag
Private Num_of_Cols
Private Type POINTAPI
        x As Long
        Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'********************Publics*************************

Public PosTop As Long: Public PosLeft As Long
Public VSpace As Long: Public HSpace As Long
Public GridPrint As Boolean: Public GridColor As Long: Public GridStyle As Long: Public GridWidth As Long
Public HorLines As Boolean 'Prints a line between the rows
Public VerLines As Boolean 'Prints a line between cols
Public RowsPerPage As Integer: Public RowHeight As Long
Public BoarderPrint As Boolean: Public BoarderWidth As Long: Public BoarderColor As Long
Public BoarderDistance As Long: Public BoarderStyle As Long
Public WordWrap As Boolean
Public DrawEmptyRect As Boolean
Public RoundX As Long: Public RoundY As Long
Public FillColor As Long: Public TextColor As Long
Public CurX As Long: Public CurY As Long


'*************Const****************
Const DT_LEFT = 0
Const DT_CENTER = 1
Const DT_RIGHT = 2
Const DT_VCENTER = 4
Const DT_WORDBREAK = 16
Const DT_SINGLELINE = 32
Const DT_NOPREFIX = 2048
Const DT_END_ELLIPSIS = 32768
Const DT_MODIFYSTRING = 65536
Const DT_WORD_ELLIPSIS = 262144
Const DT_RTLREADING = 131072  'For RightToLeft people
Const DT_CALCRECT = 1024
Private Sub Class_Initialize()
'------------------<<Default Values >>-----------------
PosTop = 50: PosLeft = 50: VSpace = 5: HSpace = 5
GridPrint = True: GridColor = vbBlack: GridStyle = 0: GridWidth = 1
RowsPerPage = 10: RowHeight = 250
BoarderPrint = True: BoarderWidth = 1: BoarderColor = vbBlack: BoarderDistance = 5: BoarderStyle = 0
Cols = Split("1000,1000,1500,1500", ",")
Aligns = Split("0,0,0,0", ",")
RoundX = 0: RoundY = 0: TextColor = vbBlack: FillColor = vbWhite
BorX = 0: BorY = 0
HorLines = False
VerLines = False
DrawEmptyRect = True


End Sub

Public Property Let ColWidthStr(NewValue As String)
Cols = Split(NewValue, ",")
End Property

Public Property Let ColAlignStr(NewValue As String)
Aligns = Split(NewValue, ",")
End Property

Public Sub PrintOut(obj As Object, ByVal PageNumber As Long)
On Error Resume Next
Dim Kount As Long, Box As RECT, Box2 As RECT, Rows As Long, sText As Variant
Dim Left As Long, Top As Long, RowHt As Long, Col As Long, bEmptyRect As Boolean
Dim DT_Code As Long, MulX As Single, MulY As Single, OldPen As Long, NewPen As Long
Dim OldBrush As Long, NewBrush As Long, opal As POINTAPI, DT_Code2 As Long, BackUpPosTop As Long
Dim HalfBorDis As Long, TmpW As Long, A As Long, bValueChanged As Boolean
Dim Box2Bt As Long, Box2Rt As Long, VLines() As Long, bFirstTime As Boolean, b1stTime As Boolean
''------------------<<OO >>-----------------
If TypeName(RecSource) <> "Recordset" Then Exit Sub
PageNumber = PageNumber - 1: PageNumber = Abs(PageNumber) 'to avoid page # 0
obj.ScaleMode = vbPixels
obj.Print
MulX = 1: MulY = 1
If TypeName(obj) = "Printer" Then
'Following lines assumes printer's pixels always smaller than screen's and Printer is installed
             MulX = Screen.TwipsPerPixelX / Printer.TwipsPerPixelX
             MulY = Screen.TwipsPerPixelY / Printer.TwipsPerPixelY
End If  'TypeName(obj) = "Printer"
RowHt = obj.ScaleY(RowHeight, vbTwips, vbPixels)
'Getting Number of fields of recordset
If Not bFieldPrinted Then
            Num_of_Cols = UBound(Cols)
            If Val(Cols(UBound(Cols))) = 0 Then
                    bFirstTime = True
                                Do While bFirstTime
                                            Num_of_Cols = Num_of_Cols - 1
                                            If Val(Cols(Num_of_Cols)) <> 0 Then bFirstTime = False:
                                            If Num_of_Cols < 0 Then bFirstTime = False: Num_of_Cols = 0
                                 Loop
            End If  'Val(Cols(UBound(Cols))) = 0

            ReDim Preserve Cols(Num_of_Cols) ''Reserving space for width of Box and for alignments
            ReDim Preserve Aligns(Num_of_Cols)
                  For Kount = 0 To Num_of_Cols ' UBound(Cols) 'Putting 0 into empty ones
                       If Aligns(Kount) = Empty Then Aligns(Kount) = "0"
                  Next Kount
            b1stTime = True
End If 'Not bFieldPrinted

If Num_of_Cols = 0 Then Exit Sub ' no cols to print
If VerLines Then ReDim Preserve VLines(Num_of_Cols): bFirstTime = True
If WordWrap Then 'Setting default alignment
             DT_Code = DT_LEFT Or DT_WORDBREAK Or DT_NOPREFIX Or DT_RTLREADING
Else
             DT_Code = DT_LEFT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_VCENTER Or DT_RTLREADING
End If

SetTextColor obj.hDC, TextColor
Left = PosLeft
Top = PosTop
HalfBorDis = BoarderDistance / 2 'for Boarders

If BoarderPrint Then
              Left = PosLeft + (HalfBorDis * MulX): BorX = PosLeft
              Top = PosTop + (HalfBorDis * MulY): BorY = PosTop
End If

If bFieldPrinted Then
              Top = PosTop + RowHt + ((VSpace - GridWidth) * MulY): bFieldPrinted = False
              If BoarderPrint Then Top = Top + (HalfBorDis * MulY)
End If
RecSource.MoveFirst 'Preparing the page
If Err = 3021 Then
             MsgBox Err.Description  'Either EOF of BOF True
End If
RecSource.Move RowsPerPage * PageNumber
If GridPrint Or HorLines Or VerLines Then   'Making Pen
              NewPen = CreatePen(GridStyle, GridWidth * MulX, GridColor)
              NewBrush = CreateSolidBrush(FillColor)
              OldPen = SelectObject(obj.hDC, NewPen)
              OldBrush = SelectObject(obj.hDC, NewBrush)
End If  'GridPrint Or HorLines
For Rows = 0 To RowsPerPage - 1
            For Col = 0 To Num_of_Cols
                    Box.Left = Left:            Box.Top = Top
                    Box.Right = Box.Left + obj.ScaleX(Val(Cols(Col)), vbTwips, vbPixels)
                    Box.Bottom = Box.Top + RowHt
                    If (VerLines And bFirstTime) Then VLines(Col) = Box.Right
                    sText = "":             If Not IsNull(RecSource.Fields(Col).Value) Then sText = RecSource.Fields(Col).Value
                    bEmptyRect = True: If Not DrawEmptyRect Then bEmptyRect = IIf(sText <> "", True, False)
                    If (GridPrint And bEmptyRect) Then RoundRect obj.hDC, Box.Left, Box.Top, Box.Right, Box.Bottom, obj.ScaleX(RoundX, vbTwips, vbPixels), obj.ScaleY(RoundY, vbTwips, vbPixels)
                    Box.Top = Box.Top + (3 * MulY): Box.Right = Box.Right - (3 * MulX)
                    Box.Left = Box.Left + (3 * MulX): Box.Bottom = Box.Bottom - (3 * MulY)
                    DT_Code2 = DT_Code Or CLng(Aligns(Col))
                    
                    If WordWrap Then
                            Box2 = Box
                            Box2Bt = Box.Bottom - Box.Top
                            A = DrawText(obj.hDC, sText, Len(sText), Box2, DT_Code2 Or DT_CALCRECT)
                            If A < Box2Bt Then
                            'Text height is smaller than the Box height, so center it Vertically
                                    Box2 = Box: Box2.Bottom = Box2.Top + A: Box2Rt = Box2.Bottom - Box2.Top
                                    Box2.Top = Box2.Top + (Box2Bt - Box2Rt) / 2
                                    Box2.Bottom = Box2.Top + A
                                    DrawText obj.hDC, sText, Len(sText), Box2, DT_Code2
                            Else
                                    'Do nothing
                                    DrawText obj.hDC, sText, Len(sText), Box, DT_Code2
                            End If  'A < Box2Bt
                    Else
                            DrawText obj.hDC, sText, Len(sText), Box, DT_Code2
                    End If  'WordWrap
                    DT_Code2 = 0
                    Left = Left + obj.ScaleX(Val(Cols(Col)), vbTwips, vbPixels) + ((HSpace - GridWidth) * MulX)
                    'Reduce the space in between columns
                    If Val(Cols(Col)) = 0 Then Left = Left + obj.ScaleX(Val(Cols(Col)), vbTwips, vbPixels) - ((HSpace - GridWidth) * MulX)
            Next Col
            bFirstTime = False
            Top = Top + RowHt + ((VSpace - GridWidth) * MulY):      Left = PosLeft
            If BoarderPrint Then Left = Left + (HalfBorDis * MulX)
            RecSource.MoveNext
            bFPage = False  'bFPage  Final Page indicator
            If Err = 3021 Then bFPage = True
            
            If HorLines Then  'Prints Horizontle lines'<<~~~~~~~~~~~~~~~~   HORIZONTLE LINES ~~~~~~~~<<

                    A = 0: If BoarderPrint Then A = HalfBorDis
           
                    If b1stTime Then
                        Box2Rt = Box.Right + (3 * MulX)
                        Box2Bt = Box.Top - (3 * MulY) ' - (((VSpace - GridWidth) / 2) * MulY)
                        MoveToEx obj.hDC, PosLeft + (A * MulX), Box2Bt, opal
                        LineTo obj.hDC, Box2Rt, Box2Bt
                        b1stTime = False
                    End If
                    
                    Box2Rt = Box.Right + (3 * MulX)
                    Box2Bt = Box.Bottom + (3 * MulY) + (((VSpace - GridWidth) / 2) * MulY)
                    If (Rows = (RowsPerPage - 1) Or bFPage) Then Box2Bt = Box.Bottom + (3 * MulY)
                    MoveToEx obj.hDC, PosLeft + (A * MulX), Box2Bt, opal
                    LineTo obj.hDC, Box2Rt, Box2Bt
            End If  'Horizontle lines <<~~~~~~~~~~~~~~~~~~~~~END OF HORIZONTLE LINE ~~~~~~<<
            If bFPage Then Exit For
Next Rows

If VerLines Then 'Print Vertical lines betn rows<<~~~~~~~~~~~ VERTICLE LINE~~~~~~~<<
            For Kount = 0 To UBound(Cols):    If Val(Cols(Kount)) = 0 Then VLines(Kount) = 0
            Next
            Box2Bt = Box.Bottom + (3 * MulY)
            A = 0: If BoarderPrint Then A = HalfBorDis
            MoveToEx obj.hDC, PosLeft + (A * MulX), PosTop + (A * MulY), opal ' Firstline
            LineTo obj.hDC, PosLeft + (A * MulX), Box2Bt
            For Kount = 0 To Num_of_Cols - 1
                    If VLines(Kount) <> 0 Then
                                Box2Rt = VLines(Kount) + (((HSpace - GridWidth) / 2) * MulX)
                                MoveToEx obj.hDC, Box2Rt, PosTop + (A * MulY), opal
                                LineTo obj.hDC, Box2Rt, Box2Bt
                    End If
          Next
          Box2Rt = Box.Right + (3 * MulX)
          MoveToEx obj.hDC, Box2Rt, PosTop + (A * MulY), opal
          LineTo obj.hDC, Box2Rt, Box2Bt
         
End If  'VerLines '<<~~~~~~~~~~~~~~~~~~~~ END OF VERTICLE LINE ~~~~~~~~~~~~~~~<<

If GridPrint Or HorLines Or VerLines Then  'Killing Pen
           SelectObject obj.hDC, OldPen: SelectObject obj.hDC, OldBrush
           DeleteObject NewPen: DeleteObject NewBrush
End If  'GridPrint Or HorLines
If Val(Cols(UBound(Cols))) = 0 Then Box.Right = Box.Right - ((HSpace - GridWidth) * MulX)
Box.Bottom = Box.Bottom + (3 * MulY): Box.Right = Box.Right + (3 * MulX)
BorDX = Box.Right + (HalfBorDis * MulX): BorDY = Box.Bottom + (HalfBorDis * MulY)
CurX = Box.Right: CurY = Box.Bottom
If BoarderPrint Then '<<~~~~~~~~~~~~~~ BOARDERS ~~~~~~~~~~~~~~~~~~<<
            CurX = BorDX:    CurY = BorDY
            NewPen = CreatePen(BoarderStyle, BoarderWidth * MulX, BoarderColor)
            OldPen = SelectObject(obj.hDC, NewPen)
            MoveToEx obj.hDC, BorDX, BorDY, opal   'Starting point Dx Dy
            LineTo obj.hDC, BorX, BorDY                      'MoveRight
            LineTo obj.hDC, BorX, BorY                         'MoveUp
            LineTo obj.hDC, BorDX, BorY                      'Moveleft
            LineTo obj.hDC, BorDX, BorDY                   'MoveDown to starting point
            SelectObject obj.hDC, OldPen
            DeleteObject NewPen
End If  'BoarderPrint
End Sub

Public Property Set RowSource(ByVal rsNewValue As Object)
'RowSource is ADO or DAO recordset object
Set RecSource = rsNewValue
End Property

Private Sub Class_Terminate()
Set RecSource = Nothing

End Sub

Public Sub PrintFields(obj As Object)
Dim Kount As Long, Box As RECT, Box2 As RECT, sText As String
Dim Left As Long, Top As Long, RowHt As Long, Col As Long
Dim DT_Code As Long, MulX As Single, MulY As Single, OldPen As Long, NewPen As Long
Dim OldBrush As Long, NewBrush As Long, opal As POINTAPI, Box2Rt As Long, Box2Bt As Long
Dim DT_Code2 As Long, bFirstTime As Boolean, HalfBorDis As Long, A As Long
'--------------------------------------
If TypeName(RecSource) <> "Recordset" Then Exit Sub
MulX = 1: MulY = 1
'Mul is multiplyer
obj.ScaleMode = vbPixels
obj.Print

If TypeName(obj) = "Printer" Then
            'Following lines assumes printer's resolution always smaller than screen
            MulX = Screen.TwipsPerPixelX / Printer.TwipsPerPixelX
            MulY = Screen.TwipsPerPixelY / Printer.TwipsPerPixelY
End If  'TypeName(obj) = "Printer"

Left = PosLeft
Top = PosTop
HalfBorDis = BoarderDistance / 2

If BoarderPrint Then  '<<~~~~~~~~~~~~~~~~~~~~~~~~~~ BOARDER ------<<
            BorX = PosLeft:            BorY = PosTop
            Left = PosLeft + (HalfBorDis * MulX):        Top = PosTop + (HalfBorDis * MulY)
End If
RowHt = obj.ScaleY(RowHeight, vbTwips, vbPixels)

Num_of_Cols = UBound(Cols) 'Getting Number of fields of recordset
If Val(Cols(UBound(Cols))) = 0 Then
             bFirstTime = True
                                Do While bFirstTime
                                            Num_of_Cols = Num_of_Cols - 1
                                            If Val(Cols(Num_of_Cols)) <> 0 Then bFirstTime = False:
                                            If Num_of_Cols < 0 Then bFirstTime = False: Num_of_Cols = 0
                                 Loop
End If  'Val(Cols(UBound(Cols))) = 0
If Num_of_Cols = 0 Then Exit Sub ' no cols to print'Setting default alignment
If WordWrap Then
            DT_Code = DT_LEFT Or DT_WORDBREAK Or DT_NOPREFIX Or DT_RTLREADING
Else
            DT_Code = DT_LEFT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_VCENTER Or DT_RTLREADING
End If
'Reserving space for width of Box and for alignments

ReDim Preserve Cols(Num_of_Cols): ReDim Preserve Aligns(Num_of_Cols)
'Putting 0 into empty ones
For Kount = 0 To Num_of_Cols 'UBound(Cols)
            If Aligns(Kount) = Empty Then Aligns(Kount) = "0"
Next Kount

SetTextColor obj.hDC, TextColor

If GridPrint Or HorLines Then
             NewPen = CreatePen(GridStyle, GridWidth * MulX, GridColor)
             NewBrush = CreateSolidBrush(FillColor)
             OldPen = SelectObject(obj.hDC, NewPen)
             OldBrush = SelectObject(obj.hDC, NewBrush)
End If  'GridPrint or HorLines

For Kount = 0 To Num_of_Cols  'RecSource.Fields.Count - 1
            sText = RecSource.Fields(Kount).Name
            Box.Left = Left:              Box.Top = Top
            Box.Right = Box.Left + obj.ScaleX(Val(Cols(Kount)), vbTwips, vbPixels)
            Box.Bottom = Box.Top + RowHt
            If GridPrint Then RoundRect obj.hDC, Box.Left, Box.Top, Box.Right, Box.Bottom, obj.ScaleX(RoundX, vbTwips, vbPixels), obj.ScaleY(RoundY, vbTwips, vbPixels)
    
            Box.Top = Box.Top + (3 * MulY): Box.Right = Box.Right - (3 * MulX)
            Box.Left = Box.Left + (3 * MulX): Box.Bottom = Box.Bottom - (3 * MulY)
            DT_Code2 = DT_Code Or Val(Aligns(Kount))
            If WordWrap Then
                            Box2 = Box
                            Box2Bt = Box.Bottom - Box.Top
                            A = DrawText(obj.hDC, sText, Len(sText), Box2, DT_Code2 Or DT_CALCRECT)
                            If A < Box2Bt Then
                            'Text height is smaller than the Box height, so center it Vertically
                                    Box2 = Box: Box2.Bottom = Box2.Top + A: Box2Rt = Box2.Bottom - Box2.Top
                                    Box2.Top = Box2.Top + (Box2Bt - Box2Rt) / 2
                                    Box2.Bottom = Box2.Top + A
                                    DrawText obj.hDC, sText, Len(sText), Box2, DT_Code2
                            Else
                                    'Do nothing
                                    DrawText obj.hDC, sText, Len(sText), Box, DT_Code2
                            End If  'A < Box2Bt
                    Else
                            DrawText obj.hDC, sText, Len(sText), Box, DT_Code2
            End If  'WordWrap
                    
            DT_Code2 = 0: sText = ""
            Left = Left + obj.ScaleX(Val(Cols(Kount)), vbTwips, vbPixels) + ((HSpace - GridWidth) * MulX)
            'Reduce the space in between columns if cols width=0
            If Val(Cols(Kount)) = 0 Then Left = Left + obj.ScaleX(Val(Cols(Kount)), vbTwips, vbPixels) - ((HSpace - GridWidth) * MulX)
   
Next Kount

If HorLines Then '<<~~~~~~~~~~~~~~~~~~~~~HORIZONTAL LINE ~~~~~~~~~~~~<<
           A = 0: If BoarderPrint Then A = HalfBorDis
           Box2Rt = Box.Right + (3 * MulX)
                       
           Box2Bt = Box.Top - (3 * MulY)
           MoveToEx obj.hDC, PosLeft + (A * MulX), Box2Bt, opal
           LineTo obj.hDC, Box2Rt, Box2Bt
                       
           Box2Bt = Box.Bottom + (3 * MulY) + (((VSpace - GridWidth) / 2) * MulY)
           MoveToEx obj.hDC, PosLeft + (A * MulX), Box2Bt, opal
           LineTo obj.hDC, Box2Rt, Box2Bt
End If  'HorLines
         
If GridPrint Or HorLines Then
           SelectObject obj.hDC, OldPen: SelectObject obj.hDC, OldBrush
           DeleteObject NewPen: DeleteObject NewBrush
End If  'GridPrint
   
bFieldPrinted = True
End Sub

Public Property Get FinalPage() As Boolean
'Flag to indicate Last page is displayed or Printed
FinalPage = bFPage
End Property

'Returns width of output in PIXELS
Public Property Get GetWidth(obj As Object) As Long
Dim Kount As Long, Mul As Single
Dim TotWidth As Long
Dim TotHSpace As Long, ToT As Long
Dim TotZeroCol As Long
Dim TotCols As Long
Mul = 1
obj.ScaleMode = vbPixels
obj.Print
If TypeName(obj) = "Printer" Then Mul = Screen.TwipsPerPixelX / Printer.TwipsPerPixelX
TotCols = UBound(Cols)
For Kount = 0 To TotCols
        TotWidth = TotWidth + obj.ScaleX(Val(Cols(Kount)), vbTwips, vbPixels)
        If Val(Cols(Kount)) = 0 Then TotZeroCol = TotZeroCol + 1
Next Kount
ToT = TotCols - TotZeroCol
TotHSpace = ToT * ((HSpace * Mul) - (GridWidth * Mul))

GetWidth = TotWidth + TotHSpace

If BoarderPrint Then GetWidth = GetWidth + (BoarderDistance * Mul)
End Property

'Returns height of output in PIXELS
Public Property Get GetHeight(obj As Object, Including_FieldNames As Boolean) As Long
Dim RHeight As Single, Mul As Single
Dim ToTHeight As Long, TotVspace As Long
Dim Num_of_Rows As Long
Num_of_Rows = RowsPerPage
Mul = 1
obj.ScaleMode = vbPixels
obj.Print
If TypeName(obj) = "Printer" Then Mul = Screen.TwipsPerPixelY / Printer.TwipsPerPixelY

If Including_FieldNames Then Num_of_Rows = Num_of_Rows + 1

RHeight = obj.ScaleY(RowHeight, vbTwips, vbPixels) + ((VSpace * Mul) - (GridWidth * Mul))
ToTHeight = (RHeight * Num_of_Rows) - ((VSpace * Mul) - (GridWidth * Mul))

If BoarderPrint Then ToTHeight = ToTHeight + (BoarderDistance * Mul)
GetHeight = ToTHeight
End Property

