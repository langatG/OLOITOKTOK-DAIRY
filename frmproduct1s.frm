VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmproduct1s 
   Caption         =   "PRODUCTS UPDATE"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   Icon            =   "frmproduct1s.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7875
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker DTPexdate 
      Height          =   375
      Left            =   1680
      TabIndex        =   39
      Top             =   5520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51642369
      CurrentDate     =   43660
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New Supplier"
      Height          =   375
      Left            =   5160
      TabIndex        =   37
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdproductaging 
      Caption         =   "Aging Products"
      Height          =   375
      Left            =   5880
      TabIndex        =   28
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txtRLevel 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   26
      Text            =   "5"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtsellingprice 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   25
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtpprice 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   23
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtpassit 
      Appearance      =   0  'Flat
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   21
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame fra1 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   3240
      TabIndex        =   20
      Top             =   3840
      Width           =   4335
      Begin VB.TextBox txtcracc 
         Height          =   375
         Left            =   1680
         TabIndex        =   36
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtdracc 
         Height          =   375
         Left            =   1680
         TabIndex        =   35
         Top             =   480
         Width           =   2535
      End
      Begin VB.PictureBox Picture3 
         Height          =   255
         Left            =   1320
         Picture         =   "frmproduct1s.frx":0442
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   34
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   1320
         Picture         =   "frmproduct1s.frx":0D0C
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   33
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Craccno"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "DrAccNo"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblcracc 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lbldracc 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkserialrequired 
      Caption         =   "Serial Required"
      Height          =   375
      Left            =   5760
      TabIndex        =   19
      Top             =   6000
      Width           =   1455
   End
   Begin VB.ComboBox cbosupplier 
      Height          =   315
      ItemData        =   "frmproduct1s.frx":15D6
      Left            =   1680
      List            =   "frmproduct1s.frx":15D8
      TabIndex        =   2
      Top             =   1800
      Width           =   4455
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   6000
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   3600
      Picture         =   "frmproduct1s.frx":15DA
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   16
      Top             =   360
      Width           =   240
   End
   Begin VB.TextBox txtbalance 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   3120
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker txtdateenterered 
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51642369
      CurrentDate     =   38814
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtserialno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtpname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox txtpcode 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "Expiry date"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Re-Order Level"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Selling Price "
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Purchase Price "
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Date Entered"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Balance In Store"
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Serial No"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Product Code"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmproduct1s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Provider As String
Dim seria As Integer
Private Sub CHKSERIALIZED_Click()
'If CHKSERIALIZED = vbChecked Then
'frmserialization.Show vbModal
'End If
End Sub

Private Sub chkserialrequired_Click()
If chkserialrequired = vbChecked Then
txtserialno.Visible = True
seria = 1
Else
seria = 0
txtserialno.Visible = False
End If
End Sub

'Public sel As String
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
txtpassit.Visible = True
fra1.Visible = True
If txtpassit = "" Then
MsgBox "Please enter Password on the text above", vbInformation
Exit Sub
End If
Dim rsp As Recordset
Set cn = CreateObject("adodb.connection")
Provider = cn
Set cn = New ADODB.Connection
Provider = "Maziwa"
cn.Open Provider
Set rsp = CreateObject("adodb.recordset")
sql = "select *  from useraccounts where UserLoginID='" & User & "' "
rsp.Open sql, cn
Dim pass As String

If Not rsp.EOF Then
pass = modsecurity.Encript_String(txtpassit)
If pass = rsp.Fields("password") Then
'txtpassit.Visible = False
Else
MsgBox "You are not allowed to delete  . Consult administrator only", vbInformation
Exit Sub

End If
Else
MsgBox "You are not allowed to delete . Consult administrator only", vbInformation
Exit Sub

End If
'*****************************************
Set rst = New Recordset
Dim bo As Boolean
'Dim cn As Connection

Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider

sql = ""
sql = "delete from ag_products where p_code='" & txtpcode & "'"
cn.Execute sql

'// delete all the details in the stock balance

sql = ""
sql = "select * from ag_stockbalance where p_code='" & txtpcode & "' order by trackid"
Set rst = New ADODB.Recordset
rst.Open sql, cn, adOpenKeyset, adLockOptimistic

If Not rst.EOF Then
While Not rst.EOF
sql = ""
sql = "delete from ag_stockbalance where trackid=" & rst.Fields("trackid") & ""
cn.Execute sql

rst.MoveNext
Wend
End If

MsgBox "You have successfully deleted product code", vbInformation
txtbalance = ""
txtpcode = ""
txtpname = ""
txtserialno = ""
txtquantity = ""

End Sub

Private Sub cmdNew_Click()

Set rs = oSaccoMaster.GetRecordset("d_sp_PNO")
If Not rs.EOF Then
txtpcode = rs.Fields(0) + 1
Else
txtpcode = 1
Exit Sub
End If

txtpassit = ""
txtsellingprice = ""
txtpprice = ""
txtquantity = ""
cbosupplier = ""
txtpname = ""
txtbalance = ""
txtserialno = ""
End Sub

Private Sub cmdproductaging_Click()
Dim lastdate As Date
Dim lastdateofsale As Date
Dim pcode As String
Dim dy As Integer
Dim grade As String
Dim rsd As New ADODB.Recordset
'//truncate this table

sql = "truncate table ag_paging"
oSaccoMaster.ExecuteThis (sql)
'select pcode,ldate,dy,auditdate,audit,grade from ag_paging
'//we look for all the products.
sql = ""
sql = "select * From ag_products order by p_code asc"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
pcode = rs.Fields("p_code")
'//get the last date
Set rst = oSaccoMaster.GetRecordset("set dateformat dmy select top 1 * from ag_stockbalance where p_code='" & pcode & "' order by transdate desc,trackid desc")
If Not rst.EOF Then
lastdate = rst.Fields("transdate")
End If
'//get the last date sold
sql = ""
sql = "select * from ag_receipts  where p_code='" & pcode & "' order by  t_date desc, r_id desc"
Set rsd = oSaccoMaster.GetRecordset(sql)
If Not rsd.EOF Then
lastdateofsale = rsd.Fields("t_date")
Else
lastdateofsale = Format(Get_Server_Date, "dd/mm/yyyy")
End If
If lastdate = "12:00:00 AM" Then
lastdate = Format(Get_Server_Date, "dd/mm/yyyy")
End If
dy = DateDiff("d", lastdate, lastdateofsale)
If dy < 0 Then
grade = "Normal"
dy = 0
GoTo shadi
End If
'0-30 days normal
If dy > 0 And dy < 30 Then
grade = "Normal"
dy = dy
GoTo shadi
End If
'31-60 substandard
If dy > 31 And dy < 60 Then
grade = "Substandard"
dy = dy
GoTo shadi
End If
'60- 90 watch
If dy > 61 And dy < 90 Then
grade = "Watch"
dy = dy
GoTo shadi
End If
'90- &&& risk
If dy > 90 Then
grade = "Risk"
dy = dy
GoTo shadi
End If
shadi:

'select pcode,ldate,dy,auditdate,audit,grade from ag_paging
sql = ""
sql = "set dateformat dmy insert into ag_paging (pcode,ldate,ltdate,dy,auditdate,audit,grade)"
sql = sql & "values('" & pcode & "','" & lastdate & "','" & lastdateofsale & "'," & dy & ",'" & Get_Server_Date & "','" & User & "','" & grade & "') "
oSaccoMaster.ExecuteThis (sql)
dy = 0
rs.MoveNext
Wend
MsgBox "Records successfully done", vbInformation

'//give him the report here
'agrovetagingreport
reportname = "agrovetagingreport.rpt"

 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
'//we look for receipts tables
'//get the number of days
'/// insert into the number of days
'//give us a report



End Sub

Private Sub cmdsave_Click()
Set rst = New Recordset
If lbldracc = "" Then MsgBox "select the account to Debit": Exit Sub

If lblcracc = "" Then MsgBox "select the account to credit": Exit Sub


'
Dim unsera As Integer
'Dim cn As Connection
If Trim(txtquantity) = "" Then
MsgBox "Quantity cannot be Zero", vbInformation
Exit Sub

End If
If Trim(txtbalance) = "" Then txtbalance = 0
If chkserialrequired = vbChecked Then

seria = 1
unsera = txtquantity

'// should only be one item
If txtquantity > 1 Then
MsgBox "Serialized items should only be added as one", vbCritical
Exit Sub
End If
Else
seria = 0
unsera = 0
End If

Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,qout,unserialized from ag_products where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If rs.EOF Then
'// insert into ag_products
If txtserialno = "" Then txtserialno = 0
sql = ""
sql = "set dateformat dmy insert into  ag_products(p_code,p_name,s_no,qin,qout,date_entered,last_d_updated,user_id,audit_date,o_bal,supplierid,serialized,unserialized,seria,pprice,sprice,Expirydate )"
sql = sql & "  values('" & txtpcode.Text & "','" & txtpname.Text & "'," & txtserialno.Text & "," & txtquantity.Text & "," & txtbalance.Text + txtquantity.Text & ",'" & txtdateenterered.value & "','" & txtdateenterered.value & "','Admin','" & Date & "'," & txtquantity.Text & ",'" & cbosupplier & "',0," & unsera & "," & seria & "," & txtpprice & "," & txtsellingprice & ",'" & DTPexdate.value & "')"
cn.Execute sql


If txtsellingprice = "" Then txtsellingprice = 0
If txtpprice = "" Then txtpprice = 0

sql = ""
sql = "set dateformat DMY INSERT INTO ag_stockbalance"
sql = sql & " (p_code, productname, openningstock, changeinstock, stockbalance, transdate,companyid,pprice,sprice,RLevel,Expirydate)"
sql = sql & " VALUES     ('" & txtpcode.Text & "','" & txtpname & "', " & txtbalance & ", " & txtquantity & ", " & txtbalance.Text + txtquantity.Text & ", '" & txtdateenterered & "',1," & txtpprice & "," & txtsellingprice & "," & txtRLevel & ",'" & DTPexdate.value & "')"
cn.Execute sql



Else
Dim d As Double
If Not IsNull(rs.Fields(2)) Then d = rs.Fields(2)
sql = ""
sql = "set dateformat DMY update ag_products set p_name='" & txtpname & "',qin=" & txtquantity.Text & ",qout=" & txtquantity.Text + rs.Fields("qout") & ",o_bal=" & txtquantity.Text + rs.Fields("qout") & ",last_d_updated='" & Date & "',user_id='" & User & "',audit_date='" & Date & "',unserialized=" & unsera + d & ",SERIA=" & seria & ",pprice=" & txtpprice & ",sprice=" & txtsellingprice & " where p_code='" & txtpcode.Text & "'"
cn.Execute sql

Dim rsst As Recordset
sql = ""
sql = "set dateformat DMY select top 1 * from ag_stockbalance where p_code='" & txtpcode & "' order by trackid desc "
Set rsst = New ADODB.Recordset
rsst.Open sql, cn
If Not rsst.EOF Then
sql = ""
sql = "set dateformat DMY INSERT INTO ag_stockbalance"
sql = sql & " (p_code, productname, openningstock, changeinstock, stockbalance, transdate,companyid)"
sql = sql & " VALUES     ('" & txtpcode & "', '" & txtpname & "', '" & txtbalance & "', '" & txtquantity & "', '" & txtquantity.Text + rs.Fields("qout") & "', '" & txtdateenterered & "',1)"
cn.Execute sql


'Else
'sql = "Update ag_stockbalance"
'sql = sql & " SET              productname = '" & txtpname & "', openningstock = " & txtbalance & ", changeinstock = " & txtquantity & ", ag_stockbalance = " & txtquantity.Text + rs.Fields("qout") & ", transdate = '" & txtdateenterered & "'"
'sql = sql & " WHERE     (p_code = '" & txtpcode & "') AND trackid=" & rsst.Fields("trackid") & ""
'cn.Execute sql
End If
'// update serialno database

'' ///update gl


End If
If seria = 1 Then
Set rst = Nothing
    sql = ""
   sql = "select * from serialno where serialno='" & txtserialno & "' AND P_CODE='" & txtpcode & "' and used=0"
   Set rst = New ADODB.Recordset
   rst.Open sql, cn, adOpenKeyset, adLockOptimistic

If rst.EOF Then
sql = ""
sql = "set dateformat DMY INSERT INTO serialno(serialno,p_code,used)"
sql = sql & " values('" & txtserialno & "','" & txtpcode & "',0)"
cn.Execute sql
Else
MsgBox "Item is in place and not yet used", vbInformation
Exit Sub
End If
End If
'****************'
sql = ""
If txtserialno = "" Then
txtserialno = 0
End If

'sql = "set dateformat dmy insert into  ag_products3(p_code,p_name,s_no,qin,qout,date_entered,last_d_updated,user_id,audit_date,o_bal,supplierid,serialized,unserialized,seria,pprice,sprice )"
'sql = sql & "  values('" & txtpcode.Text & "','" & txtpname.Text & "'," & txtSERIALNO.Text & "," & txtquantity.Text & "," & txtBalance.Text + txtquantity.Text & ",'" & txtdateenterered.value & "','" & txtdateenterered.value & "','Admin','" & Date & "'," & txtquantity.Text & ",'" & cbosupplier & "',0," & unsera & "," & seria & "," & txtpprice & "," & txtsellingprice & ")"
'cn.Execute sql

sql = ""
sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtdateenterered & "'," & txtquantity & " *" & txtpprice & ",'" & lbldracc & "','" & lblcracc & "','stock intake','" & cbosupplier & "' ,'stock intake','" & User & "',0,0)"
oSaccoMaster.ExecuteThis (sql)

txtbalance = ""
txtpcode = ""
txtpname = ""
txtserialno = ""
txtquantity = ""
txtpprice = ""
txtsellingprice = ""
cbosupplier = ""
'txtRLevel = ""
MsgBox "Record Saved Successfully"
End Sub

Private Sub Command1_Click()
frmSupplier.Show vbModal
End Sub
Private Sub Form_Load()
txtdateenterered = Format(Date, "dd,mm,yyyy")
 Set rst = New Recordset
    Dim cn As Connection
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    'provider = cn
     cn.Open Provider, "atm", "atm"
    Set rst = New Recordset
    sql = "Select companyname from ag_Supplier1"
    rst.Open sql, cn, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
    cbosupplier.AddItem rst.Fields(0)
    rst.MoveNext
    Wend
    lbldracc = "2001"
    lblcracc = "201"
    
End Sub

Private Sub lblcracc_Change()
    Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lblcracc & "'")
    If Not rst.EOF Then
    txtcracc = rst.Fields("glaccname")
    End If

End Sub

Private Sub lblcracc_Click()
    Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lblcracc & "'")
    If Not rst.EOF Then
    txtcracc = rst.Fields("glaccname")
    End If

End Sub

Private Sub lbldracc_Change()
    Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lbldracc & "'")
    If Not rst.EOF Then
    txtdracc = rst.Fields("glaccname")
    End If
End Sub

Private Sub lbldracc_Click()
    Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lbldracc & "'")
    If Not rst.EOF Then
    txtdracc = rst.Fields("glaccname")
    End If
End Sub

Private Sub Picture1_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lbldracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub Picture2_Click()
frmSearch.Show vbModal
Dim Y As String
Y = sel

If Y <> "" Then

Provider = "MAZIWA"

Set cn = New ADODB.Connection
 cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierID,pprice,sprice from ag_products where p_code='" & Y & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
If Not IsNull(rs.Fields(4)) Then cbosupplier = (rs.Fields(4))
If Not IsNull(rs.Fields(5)) Then txtpprice = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))
'If Not IsNull(rs.Fields(2)) Then txtserialno = (rs.Fields(2))
If Not IsNull(rs.Fields(3)) Then txtbalance = (rs.Fields(3))

If txtbalance <= 0 Then
MsgBox "Your stock is below zero please reorder", vbInformation
End If
'// check with serial no if it exist


End If
End If
End Sub



Private Sub Picture3_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lblcracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub txtdateenterered_Click()
fra1.Visible = True
End Sub

Private Sub txtdateenterered_KeyPress(KeyAscii As Integer)
fra1.Visible = True
End Sub

Private Sub txtdateenterered_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
fra1.Visible = True
End Sub

Private Sub txtpassword_LostFocus()
'// verufy where you have admin right to change the date
''fra1.Visible = True
'Dim rsp As Recordset
'Set cn = CreateObject("adodb.connection")
'Provider = cn
'Set cn = New ADODB.Connection
' cn.Open Provider, "atm", "atm"
'Set rsp = CreateObject("adodb.recordset")
'sql = "select *  from useraccounts where UserLoginID='" & User & "' and usergroup='administrator'"
'rsp.Open sql, cn
'Dim pass As String
'
'If Not rsp.EOF Then
'pass = modsecurity.Encript_String(txtpassword)
'If pass = rsp.Fields("password") Then
'fra1.Visible = False
'Else
'MsgBox "You are not allowed to change the date . Consult administrator only", vbInformation
'Exit Sub
'txtdateenterered = Date
'End If
'Else
'MsgBox "You are not allowed to change the date . Consult administrator only", vbInformation
'Exit Sub
'txtdateenterered = Date
'fra1.Visible = True
'End If
'
'
'End Sub
'
'Private Sub txtpassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'txtpassword_LostFocus
End Sub

Private Sub txtpcode11_Change()
'//TWNG001
Provider = "MAZIWA"
Set cn = New ADODB.Connection
 cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid,pprice, sprice from ag_products where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
 txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
If Not IsNull(rs.Fields(3)) Then txtbalance = (rs.Fields(3))
If Not IsNull(rs.Fields(4)) Then cbosupplier = (rs.Fields(4))
If Not IsNull(rs.Fields(5)) Then txtpprice = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))
If txtbalance <= 0 Then
MsgBox "Warning:Your stock is below zero please reorder", vbInformation
Else

End If
End If


'// check with serial no if it exist
End Sub

Private Sub txtpcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtpcode11_Change
Else
Exit Sub
End If
End Sub

Private Sub txtquantity_Validate(Cancel As Boolean)
If Not IsNumeric(txtquantity) Then
MsgBox "Enter values please", vbCritical
txtquantity = ""
txtquantity.SetFocus
Exit Sub
End If
End Sub
