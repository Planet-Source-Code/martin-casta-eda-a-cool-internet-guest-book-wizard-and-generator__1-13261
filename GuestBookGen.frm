VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guest Book Generator"
   ClientHeight    =   6135
   ClientLeft      =   390
   ClientTop       =   675
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   6840
      TabIndex        =   30
      Text            =   "Combo3"
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5400
      TabIndex        =   29
      Text            =   "Combo2"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Border"
      Height          =   255
      Left            =   3600
      TabIndex        =   26
      Top             =   240
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   9
      Left            =   3600
      TabIndex        =   25
      Text            =   "Combo1"
      Top             =   5520
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   8
      Left            =   3600
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   5040
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   7
      Left            =   3600
      TabIndex        =   23
      Text            =   "Combo1"
      Top             =   4560
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   6
      Left            =   3600
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   4080
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   5
      Left            =   3600
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   3600
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   4
      Left            =   3600
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   3120
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   3
      Left            =   3600
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   2640
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   3600
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   2160
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   3600
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   3600
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Initialize Guest Book"
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Guest Book"
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   5520
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   5040
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   4560
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   3600
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   4920
      ScaleHeight     =   3195
      ScaleWidth      =   3435
      TabIndex        =   27
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label6 
      Caption         =   "Background"
      Height          =   255
      Left            =   6840
      TabIndex        =   32
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Font"
      Height          =   255
      Left            =   5400
      TabIndex        =   31
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Layout"
      Height          =   255
      Left            =   4920
      TabIndex        =   28
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Field Size"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Field Name"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub Layout()
     Dim xCoord As Double
     Dim yCoord As Double
     Dim Ctr As Integer
     Picture1.Cls
     Picture1.Cls
     Picture1.CurrentX = 1500 - (Len(Text1.Text) * 15)
     Picture1.CurrentY = 300
     Picture1.Print Text1.Text
     For Ctr = 0 To 9
          If Text2(Ctr).Text <> "" Then
               xCoord = 200
               yCoord = (200 * Ctr) + 800
               Picture1.CurrentX = xCoord
               Picture1.CurrentY = yCoord
               Picture1.Print Text2(Ctr).Text
               Picture1.Line (xCoord + 1500, yCoord)-(xCoord + 1500 + (Val(Combo1(Ctr).Text) * 15), yCoord + 100), , B
          End If
     Next Ctr
End Sub
Sub GenerateASP2()
     Dim ASPString As String
     Dim Ctr As Integer
     ASPString = ASPString + "<%" + vbCrLf
     ASPString = ASPString + "Const ForReading = 1, ForWriting = 2, ForAppending = 8" + vbCrLf
     ASPString = ASPString + "Dim ThisLine" + vbCrLf
     ASPString = ASPString + "Dim PrintLine" + vbCrLf
     ASPString = ASPString + "Set FileObject = Server.CreateObject(" + Chr$(34) + "Scripting.FileSystemObject" + Chr$(34) + ")" + vbCrLf
     ASPString = ASPString + "GuestBookFile = Server.MapPath(" + Chr$(34) + "guestbook.txt" + Chr$(34) + ")" + vbCrLf
     ASPString = ASPString + "Set InputStream = FileObject.OpenTextFile (GuestBookFile, ForReading, False)" + vbCrLf
     ASPString = ASPString + "Do While not InputStream.AtEndOfStream" + vbCrLf
     ASPString = ASPString + "ThisLine = InputStream.ReadLine" + vbCrLf
     ASPString = ASPString + "PrintLine = PrintLine + ThisLine" + " + " + Chr$(34) + "<br>" + Chr$(34) + vbCrLf
     ASPString = ASPString + "Loop" + vbCrLf
     ASPString = ASPString + "InputStream.Close" + vbCrLf
     ASPString = ASPString + "Set OutputStream = Nothing" + vbCrLf
     ASPString = ASPString + "Set FileObject = Nothing" + vbCrLf
     ASPString = ASPString + "Response.Write PrintLine" + vbCrLf
     ASPString = ASPString + "%>" + vbCrLf
     If Len(App.Path) > 3 Then
          Open App.Path + "\guestbook\viewguestbook.asp" For Output As #1
     Else
          Open App.Path + "guestbook\viewguestbook.asp" For Output As #1
     End If
     Print #1, "<html><body bgcolor=" + Combo3.Text + "><center>"
     Print #1, "<font size=6 color=" + Combo2.Text + ">" + Text1.Text + "</font><br><br>"
     Print #1, "</center>" + ASPString + "<center>"
     Print #1, "<br><br>"
     Print #1, "<a href=default.htm>Main</a><br><br>"
     Print #1, "</center></body></html>"
     Close #1
End Sub




Private Sub Combo1_Click(Index As Integer)
     Call Layout
End Sub

Sub BackgroundColor()
     Picture1.BackColor = QBColor(Val(Combo3.Text))
End Sub







Private Sub Combo2_Click()
     Call FontColor
     Call Layout
End Sub
Sub BGColor()
     If Combo3.Text = "AQUA" Then Picture1.BackColor = &HFFFF00
     If Combo3.Text = "BLACK" Then Picture1.BackColor = &H0&
     If Combo3.Text = "BLUE" Then Picture1.BackColor = &HFF0000
     If Combo3.Text = "FUCHSIA" Then Picture1.BackColor = &HFF00FF
     If Combo3.Text = "GRAY" Then Picture1.BackColor = &H808080
     If Combo3.Text = "GREEN" Then Picture1.BackColor = &H8000&
     If Combo3.Text = "LIME" Then Picture1.BackColor = &HFF00&
     If Combo3.Text = "MAROON" Then Picture1.BackColor = &H80&
     If Combo3.Text = "NAVY" Then Picture1.BackColor = &H800000
     If Combo3.Text = "OLIVE" Then Picture1.BackColor = &H8080&
     If Combo3.Text = "PURPLE" Then Picture1.BackColor = &H800080
     If Combo3.Text = "RED" Then Picture1.BackColor = &HFF&
     If Combo3.Text = "SILVER" Then Picture1.BackColor = &HC0C0C0
     If Combo3.Text = "TEAL" Then Picture1.BackColor = &H808000
     If Combo3.Text = "WHITE" Then Picture1.BackColor = &HFFFFFF
     If Combo3.Text = "YELLOW" Then Picture1.BackColor = &HFFFF&
End Sub



Private Sub Combo3_Click()
     Call BGColor
     Call Layout
End Sub

Private Sub Command1_Click()
     Call GenerateHTML
     Call GenerateASP
     Call GenerateASP2
     If Text2(0).Text <> "" Then
          MsgBox ("Guest Book Generated at " + App.Path + "\guestbook")
     Else
          MsgBox ("Field Name must not be Blank")
     End If
End Sub

Private Sub Command2_Click()
     Call Initialize
End Sub
Sub GenerateASP()
     Dim ASPString As String
     Dim Ctr As Integer
     ASPString = ASPString + "<%" + vbCrLf
     ASPString = ASPString + "Const ForReading = 1, ForWriting = 2, ForAppending = 8" + vbCrLf
     ASPString = ASPString + "Set FileObject = Server.CreateObject(" + Chr$(34) + "Scripting.FileSystemObject" + Chr$(34) + ")" + vbCrLf
     ASPString = ASPString + "GuestBookFile = Server.MapPath(" + Chr$(34) + "guestbook.txt" + Chr$(34) + ")" + vbCrLf
     ASPString = ASPString + "Set OutputStream = FileObject.OpenTextFile (GuestBookFile, ForAppending, True)" + vbCrLf
     ASPString = ASPString + "OutputStream.WriteLine " + Chr$(34) + "Date : " + Chr$(34) + " + cstr(Now())" + vbCrLf
     For Ctr = 0 To 9
          If Text2(Ctr).Text <> "" Then
               ASPString = ASPString + "OutputStream.WriteLine " + Chr$(34) + Text2(Ctr).Text + " : " + Chr$(34) + " + " + " request.form(" + Chr$(34) + RemoveSpace(Text2(Ctr).Text) + Chr$(34) + ")" + vbCrLf
          End If
     Next Ctr
     ASPString = ASPString + "OutputStream.WriteLine " + Chr$(34) + " " + Chr$(34) + vbCrLf
     ASPString = ASPString + "OutputStream.Close" + vbCrLf
     ASPString = ASPString + "Set OutputStream = Nothing" + vbCrLf
     ASPString = ASPString + "Set FileObject = Nothing" + vbCrLf
     ASPString = ASPString + "%>" + vbCrLf
     If Len(App.Path) > 3 Then
          Open App.Path + "\guestbook\addguestbook.asp" For Output As #1
     Else
          Open App.Path + "guestbook\addguestbook.asp" For Output As #1
     End If
     Print #1, ASPString
     Print #1, "<html><body bgcolor=" + Combo3.Text + "><center>"
     Print #1, "<font size=6 color=" + Combo2.Text + ">" + Text1.Text + "</font><br><br>"
     Print #1, "<table cellpadding=5 border=" + Str$(Check1.Value) + ">"
     For Ctr = 0 To 9
          If Text2(Ctr).Text <> "" Then
               Print #1, "<tr>"
               Print #1, "<td><font color=" + Combo2.Text + ">"
               Print #1, Text2(Ctr).Text
               Print #1, "</font></td><td>"
               Print #1, "<%=request.form(" + Chr$(34) + RemoveSpace(Text2(Ctr).Text) + Chr$(34) + ")%>"
               Print #1, "</td>"
               Print #1, "</tr>"
          End If
     Next Ctr
     Print #1, "<tr>"
     Print #1, "<td>"
     Print #1, "</table><br><br>"
     Print #1, "Successfully Saved<br><br>"
     Print #1, "<a href=default.htm>Main</a><br><br>"
     Print #1, "<a href=viewguestbook.asp>View</a>"
     Print #1, "</center></body></html>"
     Close #1
End Sub
Sub FontColor()
     If Combo2.Text = "AQUA" Then Picture1.ForeColor = &HFFFF00
     If Combo2.Text = "BLACK" Then Picture1.ForeColor = &H0&
     If Combo2.Text = "BLUE" Then Picture1.ForeColor = &HFF0000
     If Combo2.Text = "FUCHSIA" Then Picture1.ForeColor = &HFF00FF
     If Combo2.Text = "GRAY" Then Picture1.ForeColor = &H808080
     If Combo2.Text = "GREEN" Then Picture1.ForeColor = &H8000&
     If Combo2.Text = "LIME" Then Picture1.ForeColor = &HFF00&
     If Combo2.Text = "MAROON" Then Picture1.ForeColor = &H80&
     If Combo2.Text = "NAVY" Then Picture1.ForeColor = &H800000
     If Combo2.Text = "OLIVE" Then Picture1.ForeColor = &H8080&
     If Combo2.Text = "PURPLE" Then Picture1.ForeColor = &H800080
     If Combo2.Text = "RED" Then Picture1.ForeColor = &HFF&
     If Combo2.Text = "SILVER" Then Picture1.ForeColor = &HC0C0C0
     If Combo2.Text = "TEAL" Then Picture1.ForeColor = &H808000
     If Combo2.Text = "WHITE" Then Picture1.ForeColor = &HFFFFFF
     If Combo2.Text = "YELLOW" Then Picture1.ForeColor = &HFFFF&
End Sub
Private Sub Form_Load()
     Call Initialize
     On Error Resume Next
     If Len(App.Path) > 3 Then
          MkDir App.Path + "\guestbook\"
     Else
          MkDir App.Path + "guestbook\"
     End If
     Call PopulateSelection
     Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
End Sub
Sub GenerateHTML()
     Dim Ctr As Integer
     If Len(App.Path) > 3 Then
          Open App.Path + "\guestbook\default.htm" For Output As #1
     Else
          Open App.Path + "guestbook\default.htm" For Output As #1
     End If
     Print #1, "<html><body bgcolor=" + Combo3.Text + "><center>"
     Print #1, "<font size=6 color=" + Combo2.Text + ">" + Text1.Text + "</font><br><br>"
     Print #1, "<form method=post action=addguestbook.asp>"
     Print #1, "<table cellpadding=5 border=" + Str$(Check1.Value) + ">"
     For Ctr = 0 To 9
          If Text2(Ctr).Text <> "" Then
               Print #1, "<tr>"
               Print #1, "<td><font color=" + Combo2.Text + ">"
               Print #1, Text2(Ctr).Text
               Print #1, "</font></td><td>"
               Print #1, "<input name=" + RemoveSpace(Text2(Ctr).Text) + " size=" + Combo1(Ctr).Text + ">"
               Print #1, "</td>"
               Print #1, "</tr>"
          End If
     Next Ctr
     Print #1, "<tr>"
     Print #1, "<td>"
     Print #1, "</td><td>"
     Print #1, "<input type=reset>"
     Print #1, "<input type=submit>"
     Print #1, "</td>"
     Print #1, "</tr>"
     Print #1, "</table>"
     Print #1, "</form><br>"
     Print #1, "<a href=viewguestbook.asp>View</a>"
     Print #1, "</center></body></html>"
     Close #1
End Sub

Sub Initialize()
     Text1.Text = ""
     Dim Ctr As Integer
     For Ctr = 0 To 9
          Text2(Ctr).Text = ""
     Next Ctr
     For Ctr = 0 To 9
          Combo1(Ctr).Text = "50"
     Next Ctr
     Combo2.Text = "BLACK"
     Combo3.Text = "SILVER"
     Check1.Value = 0
End Sub
Sub PopulateSelection()
     Dim Ctr1 As Integer
     Dim Ctr2 As Integer
     For Ctr = 0 To 9
          For Ctr2 = 10 To 100 Step 10
               Combo1(Ctr).AddItem Str$(Ctr2)
          Next Ctr2
     Next Ctr
     Combo2.AddItem "AQUA"
     Combo2.AddItem "BLACK"
     Combo2.AddItem "BLUE"
     Combo2.AddItem "FUCHSIA"
     Combo2.AddItem "GRAY"
     Combo2.AddItem "GREEN"
     Combo2.AddItem "LIME"
     Combo2.AddItem "MAROON"
     Combo2.AddItem "NAVY"
     Combo2.AddItem "OLIVE"
     Combo2.AddItem "PURPLE"
     Combo2.AddItem "RED"
     Combo2.AddItem "SILVER"
     Combo2.AddItem "TEAL"
     Combo2.AddItem "WHITE"
     Combo2.AddItem "YELLOW"
     Combo3.AddItem "AQUA"
     Combo3.AddItem "BLACK"
     Combo3.AddItem "BLUE"
     Combo3.AddItem "FUCHSIA"
     Combo3.AddItem "GRAY"
     Combo3.AddItem "GREEN"
     Combo3.AddItem "LIME"
     Combo3.AddItem "MAROON"
     Combo3.AddItem "NAVY"
     Combo3.AddItem "OLIVE"
     Combo3.AddItem "PURPLE"
     Combo3.AddItem "RED"
     Combo3.AddItem "SILVER"
     Combo3.AddItem "TEAL"
     Combo3.AddItem "WHITE"
     Combo3.AddItem "YELLOW"
End Sub
Function RemoveSpace(FieldName As String) As String
     Dim TempStorage As String
     Dim Ctr As Integer
     For Ctr = 1 To Len(FieldName)
          If Mid$(FieldName, Ctr, 1) <> " " Then
               TempStorage = TempStorage + Mid$(FieldName, Ctr, 1)
          End If
     Next Ctr
     RemoveSpace = TempStorage
End Function

Private Sub Text1_Change()
      Call Layout
End Sub

Private Sub Text2_Change(Index As Integer)
     Call Layout
End Sub
