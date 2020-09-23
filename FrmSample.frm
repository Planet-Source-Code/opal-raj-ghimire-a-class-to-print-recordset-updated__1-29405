VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sample Demo"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11220
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   ScaleHeight     =   511
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   748
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hor Lines"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9810
      TabIndex        =   20
      Top             =   810
      Value           =   1  'Checked
      Width           =   1320
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ver Lines"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   9810
      TabIndex        =   19
      Top             =   450
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "RoundRect"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   9810
      TabIndex        =   18
      Top             =   90
      Width           =   1275
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9810
      TabIndex        =   16
      Top             =   7065
      Width           =   1365
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   375
      Left            =   10440
      Max             =   50
      Min             =   3
      TabIndex        =   15
      Top             =   2295
      Value           =   3
      Width           =   420
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   240
      LargeChange     =   25
      Left            =   9810
      Max             =   1200
      Min             =   200
      TabIndex        =   13
      Top             =   6660
      Value           =   200
      Width           =   1365
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "WordWrap"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9810
      TabIndex        =   12
      Top             =   1575
      Width           =   1185
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   240
      Left            =   9810
      Max             =   30
      Min             =   -1
      TabIndex        =   8
      Top             =   6165
      Width           =   1365
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   240
      Left            =   9810
      Max             =   30
      Min             =   -1
      TabIndex        =   7
      Top             =   5670
      Width           =   1365
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      Left            =   9810
      Max             =   200
      SmallChange     =   50
      TabIndex        =   6
      Top             =   5175
      Width           =   1365
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Boarder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9810
      TabIndex        =   5
      Top             =   1215
      Value           =   1  'Checked
      Width           =   1050
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9810
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "25"
      Top             =   2295
      Width           =   645
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<<Page Down"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9765
      TabIndex        =   2
      Top             =   3285
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Page Up>>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9765
      TabIndex        =   1
      Top             =   2745
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9765
      TabIndex        =   0
      ToolTipText     =   "        Prints the current page     "
      Top             =   4365
      Width           =   1365
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9855
      TabIndex        =   17
      Top             =   3915
      Width           =   1230
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Row Height"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9855
      TabIndex        =   14
      Top             =   6435
      Width           =   1230
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "H Space"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9900
      TabIndex        =   11
      Top             =   5940
      Width           =   1230
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "V Space"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9855
      TabIndex        =   10
      Top             =   5445
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Round Rect"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9810
      TabIndex        =   9
      Top             =   4950
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of lines per page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9810
      TabIndex        =   3
      Top             =   1845
      Width           =   1365
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long


Dim Res As ADODB.Recordset
Dim Con As ADODB.Connection
Dim M As RsPrinter
Dim PageNumber As Long

Private Sub Check1_Click()
M.GridPrint = IIf(Check1.Value, True, False)
End Sub

Private Sub Check2_Click()
M.BoarderPrint = IIf(Check2.Value, True, False)
End Sub

Private Sub Check3_Click()
M.WordWrap = IIf(Check3.Value, True, False)
End Sub

Private Sub Check4_Click()
M.VerLines = IIf(Check4.Value, True, False)
End Sub

Private Sub Check5_Click()
M.HorLines = IIf(Check5.Value, True, False)
End Sub

Private Sub Command1_Click()
On Error Resume Next
Printer.Print
If Err = 482 Then
    'printer is not installed
    MsgBox Err.Description
    Exit Sub
End If
'Printer.Orientation = vbPRORLandscape
Printer.Orientation = vbPRORPortrait


Printer.FontName = Form1.FontName
Printer.FontSize = Form1.FontSize

M.PosLeft = (Printer.ScaleWidth - M.GetWidth(Printer)) / 2            'Align center left-right
M.PosTop = (Printer.ScaleHeight - M.GetHeight(Printer, True)) / 2  'Align center top-bottom


M.PrintFields Printer                     'Sending fieldnames
M.PrintOut Printer, PageNumber  'Sending content

Printer.CurrentY = M.CurY + 10     ' 10 px below
'Calculating to center the page number
Printer.CurrentX = (M.PosLeft + M.CurX - Printer.TextWidth("Page No " + Str(PageNumber))) / 2
Printer.Print "Page No. "; PageNumber

Printer.EndDoc
'Re setting the Left and top for screen
'because we are using same instance of rsPrinter class for printing and previewing
M.PosLeft = 20
M.PosTop = 20


End Sub

Private Sub Command2_Click()
Cls
PageNumber = PageNumber + 1
M.FillColor = &HC0C0C0

M.PrintFields Form1
M.FillColor = vbWhite

M.PrintOut Form1, PageNumber
Form1.CurrentY = M.CurY + 10
Form1.CurrentX = (M.PosLeft + M.CurX - TextWidth("Page No " + Str(PageNumber))) / 2


Form1.Print "Page No. "; PageNumber
If M.FinalPage Then Command2.Enabled = False

If PageNumber > 1 Then
Command3.Enabled = True
End If


End Sub

Private Sub Command3_Click()
Cls
PageNumber = PageNumber - 1
M.FillColor = &HC0C0C0
M.PrintFields Form1
M.FillColor = vbWhite
M.PrintOut Form1, PageNumber

Form1.CurrentY = M.CurY + 10
Form1.CurrentX = (M.PosLeft + M.CurX - TextWidth("Page No " + Str(PageNumber))) / 2
Command2.Enabled = True
Form1.Print "Page No. "; PageNumber

If PageNumber = 1 Then Command3.Enabled = False

End Sub










Private Sub Command4_Click()
'This is just for fun.....
'-----------------------------------------
'    code in here are generated using my other project.
'    to find it visit http://geocities.com/opalraj/vb
'-----------------------------------------
Cls
Form1.ScaleMode = 3 ' Pixels
Form1.DrawWidth = 1
Form1.FillColor = 0
Form1.ForeColor = 0
Form1.FillStyle = 1
Form1.BackColor = 16777215
'-----------------------------------------
Form1.Line (47, 259)-(47, 411), 0
Form1.Line (33, 400)-(259, 400), 0
Form1.Line (46, 271)-(54, 267), 0
Form1.Line (54, 267)-(62, 268), 0
Form1.Line (63, 268)-(69, 274), 0
Form1.Line (69, 274)-(74, 281), 0
Form1.Line (74, 281)-(74, 288), 0
Form1.Line (74, 288)-(79, 297), 0
Form1.Line (79, 297)-(73, 299), 0
Form1.Line (73, 299)-(73, 301), 0
Form1.Line (73, 301)-(70, 301), 0
Form1.Line (70, 301)-(74, 303), 0
Form1.Line (74, 303)-(74, 307), 0
Form1.Line (74, 307)-(71, 310), 0
Form1.Line (71, 310)-(66, 310), 0
Form1.Line (66, 310)-(60, 308), 0
Form1.Line (67, 310)-(67, 316), 0
Form1.Line (67, 316)-(72, 322), 0
Form1.Line (72, 322)-(78, 326), 0
Form1.Line (78, 326)-(84, 339), 0
Form1.Line (84, 339)-(87, 354), 0
Form1.Line (87, 354)-(90, 362), 0
Form1.Line (90, 362)-(163, 371), 0
Form1.Line (164, 372)-(168, 376), 0
Form1.Line (168, 376)-(191, 376), 0
Form1.Line (191, 376)-(194, 401), 0
Form1.Line (192, 382)-(208, 384), 0
Form1.Line (208, 384)-(208, 402), 0
Form1.Line (208, 385)-(219, 371), 0
Form1.Line (219, 371)-(220, 365), 0
Form1.Line (220, 365)-(220, 360), 0
Form1.Line (220, 360)-(229, 360), 0
Form1.Line (229, 360)-(229, 375), 0
Form1.Line (229, 375)-(223, 383), 0
Form1.Line (223, 383)-(228, 386), 0
Form1.Line (228, 386)-(228, 398), 0
Form1.Line (228, 398)-(225, 401), 0
Form1.Line (67, 317)-(47, 321), 0
Form1.Line (91, 361)-(91, 370), 0
Form1.Line (91, 370)-(80, 376), 0
Form1.Line (80, 376)-(72, 387), 0
Form1.Line (72, 387)-(67, 388), 0
Form1.Line (67, 388)-(67, 400), 0
Form1.Line (71, 277)-(65, 277), 0
Form1.Line (65, 277)-(58, 277), 0
Form1.Line (58, 277)-(58, 285), 0
Form1.Line (58, 285)-(62, 293), 0
Form1.Line (62, 293)-(59, 297), 0
Form1.Line (59, 297)-(55, 290), 0
Form1.Line (55, 290)-(54, 296), 0
Form1.Line (54, 296)-(57, 301), 0
Form1.Line (57, 301)-(47, 309), 0
Form1.Line (65, 287)-(65, 287), 0
Form1.Line (65, 287)-(70, 284), 0
Form1.Line (69, 284)-(72, 285), 0
Form1.Line (71, 288)-(68, 288), 0
Form1.Line (162, 371)-(165, 372), 0
Form1.FillStyle = 2
Form1.FillColor = 0
ExtFloodFill Form1.hDC, 122, 383, 16777215, 1
Form1.FillStyle = 0
Form1.FillColor = 8421504
ExtFloodFill Form1.hDC, 110, 392, 16777215, 1
ExtFloodFill Form1.hDC, 111, 373, 16777215, 1
ExtFloodFill Form1.hDC, 221, 389, 16777215, 1
Form1.FillStyle = 1
Form1.Line (58, 267)-(67, 269), 0
Form1.FillColor = 12632256
Form1.FillStyle = 0
ExtFloodFill Form1.hDC, 52, 280, 16777215, 1
Form1.FillStyle = 6
ExtFloodFill Form1.hDC, 63, 353, 16777215, 1
'-----------------------------------------
Form1.FillColor = 12632256
Form1.ForeColor = 0
Form1.FillStyle = 1
Form1.BackColor = 16777215
'-----------------------------------------
Form1.Line (434, 265)-(428, 284), 0
Form1.Line (428, 284)-(437, 307), 0
Form1.Line (437, 307)-(435, 311), 0
Form1.Line (435, 311)-(424, 308), 0
Form1.Line (424, 308)-(423, 314), 0
Form1.Line (423, 314)-(414, 314), 0
Form1.Line (414, 314)-(423, 318), 0
Form1.Line (423, 318)-(423, 320), 0
Form1.Line (423, 320)-(420, 322), 0
Form1.Line (420, 322)-(417, 332), 0
Form1.Line (417, 332)-(405, 332), 0
Form1.Line (405, 332)-(394, 327), 0
Form1.Line (394, 327)-(389, 319), 0
Form1.Line (437, 307)-(452, 288), 0
Form1.Line (452, 288)-(455, 271), 0
Form1.Line (437, 307)-(441, 311), 0
Form1.Line (441, 311)-(448, 314), 0
Form1.Line (448, 314)-(445, 317), 0
Form1.Line (445, 317)-(451, 323), 0
Form1.Line (451, 323)-(445, 323), 0
Form1.Line (445, 323)-(445, 325), 0
Form1.Line (445, 325)-(447, 326), 0
Form1.Line (447, 328)-(445, 336), 0
Form1.Line (445, 336)-(450, 340), 0
Form1.Line (450, 340)-(459, 339), 0
Form1.Line (459, 339)-(480, 334), 0
Form1.Line (446, 324)-(448, 328), 0
Form1.Line (459, 284)-(468, 287), 0
Form1.Line (461, 284)-(461, 288), 0
Form1.Line (457, 288)-(467, 288), 0
Form1.Line (460, 279)-(468, 283), 0
Form1.Line (419, 281)-(420, 277), 0
Form1.Line (421, 278)-(416, 278), 0
Form1.Line (414, 277)-(421, 282), 0
Form1.Line (421, 274)-(418, 274), 0
Form1.Line (418, 274)-(414, 274), 0
Form1.Line (414, 274)-(411, 276), 0
Form1.Line (455, 272)-(448, 262), 0
Form1.Line (448, 262)-(451, 252), 0
Form1.Line (451, 252)-(469, 248), 0
Form1.Line (469, 247)-(499, 256), 0
Form1.Line (499, 256)-(506, 271), 0
Form1.Line (506, 271)-(512, 289), 0
Form1.Line (512, 289)-(515, 309), 0
Form1.Line (453, 269)-(470, 268), 0
Form1.Line (470, 268)-(477, 278), 0
Form1.Line (477, 278)-(477, 286), 0
Form1.Line (477, 286)-(482, 283), 0
Form1.Line (482, 283)-(483, 291), 0
Form1.Line (483, 291)-(480, 298), 0
Form1.Line (480, 298)-(472, 300), 0
Form1.Line (472, 300)-(483, 307), 0
Form1.Line (483, 307)-(490, 302), 0
Form1.Line (490, 302)-(494, 296), 0
Form1.Line (494, 296)-(489, 312), 0
Form1.Line (489, 312)-(480, 316), 0
Form1.Line (480, 316)-(472, 312), 0
Form1.Line (472, 312)-(477, 320), 0
Form1.Line (477, 320)-(488, 327), 0
Form1.Line (488, 327)-(499, 324), 0
Form1.Line (499, 324)-(498, 333), 0
Form1.Line (498, 333)-(510, 337), 0
Form1.Line (510, 337)-(523, 330), 0
Form1.Line (523, 330)-(533, 312), 0
Form1.Line (533, 312)-(526, 303), 0
Form1.Line (525, 303)-(517, 304), 0
Form1.Line (517, 304)-(514, 308), 0
Form1.Line (523, 303)-(533, 308), 0
Form1.Line (434, 266)-(442, 261), 0
Form1.Line (442, 260)-(444, 249), 0
Form1.Line (444, 249)-(422, 237), 0
Form1.Line (422, 237)-(402, 237), 0
Form1.Line (402, 237)-(389, 238), 0
Form1.Line (389, 238)-(378, 248), 0
Form1.Line (378, 248)-(376, 264), 0
Form1.Line (376, 264)-(371, 275), 0
Form1.Line (371, 275)-(369, 286), 0
Form1.Line (369, 286)-(357, 300), 0
Form1.Line (357, 300)-(380, 310), 0
Form1.Line (380, 310)-(387, 307), 0
Form1.Line (387, 307)-(388, 289), 0
Form1.Line (388, 289)-(391, 282), 0
Form1.Line (391, 282)-(395, 276), 0
Form1.Line (395, 276)-(399, 285), 0
Form1.Line (399, 285)-(404, 285), 0
Form1.Line (404, 285)-(401, 274), 0
Form1.Line (401, 274)-(401, 263), 0
Form1.Line (401, 263)-(410, 259), 0
Form1.Line (410, 259)-(425, 260), 0
Form1.Line (425, 260)-(436, 268), 0
Form1.Line (401, 331)-(396, 340), 0
Form1.Line (469, 338)-(475, 348), 0
Ellipse Form1.hDC, 337, 208, 550, 385
Form1.Line (397, 338)-(413, 366), 0
Form1.Line (413, 366)-(413, 381), 0
Form1.Line (474, 347)-(474, 360), 0
Form1.Line (474, 360)-(458, 384), 0
Form1.Line (513, 335)-(527, 352), 0
Form1.Line (368, 305)-(350, 340), 0
Form1.Line (442, 259)-(441, 263), 0
Form1.FillStyle = 0
ExtFloodFill Form1.hDC, 464, 231, 16777215, 1
ExtFloodFill Form1.hDC, 426, 357, 16777215, 1
Form1.FillColor = 8421504
ExtFloodFill Form1.hDC, 392, 253, 16777215, 1
ExtFloodFill Form1.hDC, 488, 274, 16777215, 1
Form1.FillStyle = 1
Form1.Circle (102, 271), 10
Form1.Circle (183, 264), 15
Form1.Circle (271, 255), 25

'-----------------------------------------

Form1.FillColor = 0
Form1.ForeColor = 0
Form1.FillStyle = 1
Form1.BackColor = 16777215
'-----------------------------------------
Form1.Line (112, 110)-(93, 179), 0
Form1.Line (93, 179)-(128, 182), 0
Form1.Line (128, 182)-(111, 111), 0
Form1.Line (194, 128)-(178, 201), 0
Form1.Line (178, 201)-(223, 199), 0
Form1.Line (223, 199)-(193, 127), 0
Form1.Line (193, 130)-(196, 156), 0
Form1.Line (196, 156)-(190, 173), 0
Form1.Line (190, 173)-(198, 201), 0
Form1.Line (112, 113)-(111, 141), 0
Form1.Line (111, 141)-(106, 153), 0
Form1.Line (106, 153)-(112, 167), 0
Form1.Line (112, 167)-(106, 180), 0
Form1.Line (106, 179)-(105, 206), 0
Form1.Line (105, 206)-(109, 206), 0
Form1.Line (109, 206)-(113, 180), 0
Form1.Line (192, 198)-(194, 227), 0
Form1.Line (194, 227)-(198, 227), 0
Form1.Line (198, 227)-(195, 199), 0
Ellipse Form1.hDC, 458, 75, 508, 130
Form1.Line (445, 103)-(400, 103), 0
Form1.Line (519, 102)-(561, 102), 0
Form1.Line (484, 140)-(484, 177), 0
Form1.Line (481, 63)-(481, 31), 0
Form1.Line (459, 77)-(437, 57), 0
Form1.Line (507, 78)-(532, 55), 0
Form1.Line (458, 128)-(439, 150), 0
Form1.Line (510, 130)-(535, 151), 0
Form1.Line (245, 144)-(251, 140), 0
Form1.Line (251, 140)-(254, 140), 0
Form1.Line (254, 140)-(257, 143), 0
Form1.Line (257, 143)-(260, 142), 0
Form1.Line (260, 142)-(264, 143), 0
Form1.Line (254, 146)-(261, 140), 0
Form1.Line (277, 155)-(287, 151), 0
Form1.Line (287, 151)-(291, 154), 0
Form1.Line (291, 154)-(294, 153), 0
Form1.Line (294, 153)-(303, 160), 0
Form1.Line (286, 157)-(295, 151), 0
Form1.Line (158, 126)-(165, 124), 0
Form1.Line (165, 124)-(167, 120), 0
Form1.Line (175, 128)-(182, 128), 0
Form1.Line (182, 128)-(188, 124), 0
Form1.Line (179, 124)-(183, 131), 0
Form1.Line (70, 120)-(75, 120), 0
Form1.Line (75, 120)-(75, 122), 0
Form1.Line (75, 122)-(80, 121), 0
Form1.Line (80, 121)-(83, 119), 0
Form1.Line (78, 119)-(73, 126), 0
Form1.Line (29, 136)-(36, 132), 0
Form1.Line (36, 132)-(39, 133), 0
Form1.Line (39, 133)-(43, 132), 0
Form1.Line (43, 132)-(48, 136), 0
Form1.Line (43, 128)-(36, 135), 0
Form1.Line (58, 152)-(61, 151), 0
Form1.Line (61, 151)-(65, 154), 0
Form1.Line (65, 154)-(69, 153), 0
Form1.Line (68, 151)-(72, 152), 0
Form1.Line (65, 149)-(66, 158), 0
Form1.FillColor = 65280
Form1.FillStyle = 0
ExtFloodFill Form1.hDC, 120, 169, 16777215, 1
ExtFloodFill Form1.hDC, 206, 181, 16777215, 1
Form1.FillColor = 32768
ExtFloodFill Form1.hDC, 106, 169, 16777215, 1
ExtFloodFill Form1.hDC, 187, 182, 16777215, 1
Form1.FillColor = 65535
ExtFloodFill Form1.hDC, 482, 95, 16777215, 1

CurrentX = 46
CurrentY = 420
FontSize = 10
Print "Recordset Printing Class by Opal R Ghimire"
CurrentX = 46
Print "Updated on 02.02.2002"
CurrentX = 46
Print "buna48@hotmail.com"

CurrentX = 46
Print "Nepal"

FontSize = 8


End Sub

Private Sub Form_Load()

PageNumber = 0
Set Res = New ADODB.Recordset
Set Con = New ADODB.Connection

Con.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" + App.Path + "\bb.mdb;Jet OLEDB;"
Con.Open

Res.Open "publishers", Con, adOpenKeyset, adLockPessimistic

'set up the class
ClassSetup
cover

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'MsgBox Str(x) & " " & Str(Y)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set M = Nothing
Res.Close
Con.Close
Set Res = Nothing
Set Con = Nothing

End Sub


Private Sub ClassSetup()

Set M = New RsPrinter                          'a copy of the class
With M

Set .RowSource = Res                       'Supply of the recordset object
'Note
'To use DAO or ADO data control
'set .RowSource=Data1.Recordset
'But one should not use this in form's load event
'because the data control populates recordset after the load event
.RowsPerPage = 25                       'Number of lines(rows) to display
.RowHeight = 240                           'Cell height, it should be tall enough to fit the text
'Following line sets width of each coloumn in twips supplying 0 value will not print that column
.ColWidthStr = "500,1500,2000,1000,1000,000,000,800,700,700"
.GridWidth = 1
.VSpace = 0                                          'Vertical space bet'n cols
.HSpace = 3                                          'Horizontal space
.RoundX = 100                                      'X  Value to round the cornor of rect
.RoundY = 100                                      'Ditto but Y
.BoarderStyle = 0
.BoarderWidth = 1
.BoarderDistance = 6
.BoarderPrint = True

.WordWrap = False ' if you make it true, make enough rowheight  too
.DrawEmptyRect = True ' draws rects even if there is no text, effective only when .GridPrint=True
.GridPrint = False 'this is to draw round rect around text
.HorLines = True  'Draws horizontal lines between rows
.VerLines = True   'Draws vertical lines between cols
'following line sets the alignmets of cols
'0= Left  1=Center  2= Right
.ColAlignStr = "1,0,0,0,0,1,0"             'meaning all left aligned but 1st and 6th center
.PosLeft = 20                                        'LeftPos for printer should not be this small
.PosTop = 20                                        'TopPos for printer should not be this small

End With

'control values
HScroll1.Value = 50
HScroll2.Value = 0 'vspace
HScroll3.Value = 3 'hspace
HScroll4.Value = 240
VScroll1.Value = 25
End Sub

Private Sub HScroll1_Change()
M.RoundX = HScroll1.Value
M.RoundY = HScroll1.Value
Label6.ToolTipText = "  Round Rect   "
End Sub

Private Sub HScroll1_Scroll()
Label6 = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
M.VSpace = HScroll2.Value
Label6.ToolTipText = "  V Space   "
End Sub

Private Sub HScroll2_Scroll()
Label6 = HScroll2.Value
End Sub

Private Sub HScroll3_Change()
M.HSpace = HScroll3.Value
Label6.ToolTipText = "  H Space   "
End Sub

Private Sub HScroll3_Scroll()
Label6 = HScroll3.Value
End Sub

Private Sub HScroll4_Change()
M.RowHeight = HScroll4.Value
Label6.ToolTipText = "  Row Height   "
End Sub

Private Sub HScroll4_Scroll()
Label6 = HScroll4.Value
End Sub


Private Sub VScroll1_Change()
M.RowsPerPage = VScroll1.Value
Text2 = VScroll1.Value

End Sub


Sub cover()
FontSize = 12
CurrentX = 400: CurrentY = 170
FontBold = True
Print "Recordset Printing Class"
FontBold = False
CurrentX = 400:
Print "Click [Page up]  to start"
Print
CurrentX = 400:: CurrentY = 470
Print "Pls read readme.txt b4 use"


CurrentX = 53: CurrentY = 370
Print "You are free to use this class in your projects. "
CurrentX = 53
Print "The class comes with absolutely NO"
CurrentX = 53
Print "warranty ! Use it at your own risk !"
FontSize = 8

End Sub
