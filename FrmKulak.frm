VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmKulak 
   Caption         =   "Re-Stock"
   ClientHeight    =   8460
   ClientLeft      =   2190
   ClientTop       =   1620
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   14835
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   19
      Top             =   7320
      Width           =   6855
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2640
         TabIndex        =   22
         Top             =   240
         Width           =   1635
      End
      Begin VB.CommandButton cmdKeluar 
         Caption         =   "&Keluar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4560
         TabIndex        =   21
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton cmdBaru 
         Caption         =   "&Baru"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   360
         TabIndex        =   20
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detail Pembalian"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   14655
      Begin VB.ComboBox cmbSatuan 
         Height          =   315
         Left            =   9360
         TabIndex        =   10
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtKode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   9
         Top             =   360
         Width           =   2550
      End
      Begin VB.TextBox txtHarga 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   8
         Text            =   "0"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtJumlah 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "0"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtnota 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   6
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton cmdMasuk 
         Caption         =   "Tambah"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10800
         TabIndex        =   5
         Top             =   1560
         Width           =   1245
      End
      Begin VB.TextBox txtNama 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   4
         Top             =   840
         Width           =   4335
      End
      Begin VB.CommandButton cmdCari 
         BackColor       =   &H8000000A&
         Caption         =   "&Cari"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         MaskColor       =   &H00404040&
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   960
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Format          =   118685697
         CurrentDate     =   44122
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5160
         TabIndex        =   18
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5160
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   8400
         TabIndex        =   16
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Beli (Rp)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5160
         TabIndex        =   15
         Top             =   1320
         Width           =   1560
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Barang"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5160
         TabIndex        =   14
         Top             =   1800
         Width           =   1470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Masuk"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1470
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor  Masuk"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1410
      End
      Begin VB.Line Line1 
         X1              =   5040
         X2              =   5040
         Y1              =   120
         Y2              =   1680
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      TabIndex        =   0
      Top             =   7320
      Width           =   7695
      Begin VB.Label Label10 
         Caption         =   "Grand Total :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label8 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   0
      TabIndex        =   23
      Top             =   3480
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NO"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kode Barang"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama Barang"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Harga Beli "
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Satuan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Sub Total"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   0
      Picture         =   "FrmKulak.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu ini digunakan untuk melakukan Transaksi Barang Masuk"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1200
      TabIndex        =   25
      Top             =   600
      Width           =   5460
   End
   Begin VB.Label LblPenjualanObat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transaksi Barang Masuk Toko Rafacell"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1200
      TabIndex        =   24
      Top             =   120
      Width           =   6345
   End
End
Attribute VB_Name = "FrmKulak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Koneksi As New ADODB.Connection
Dim rsdatabarang As New ADODB.Recordset
Dim rsbarangmasuk As New ADODB.Recordset
Dim rspemasok As New ADODB.Recordset
Dim rscustomer As New ADODB.Recordset
Dim rspenjualan As New ADODB.Recordset
Dim rspembelian As New ADODB.Recordset

Private Sub oto1()
Call Konek_DB
Set rsnooto = New ADODB.Recordset
rsnooto.Open ("SELECT * FROM Pembelian Where No_Masuk In(Select max(No_Masuk) FROM Pembelian)Order By NO asc"), Koneksi
rsnooto.Requery
    Dim Urutan As String * 11
    Dim Hitung As Long
    With rsnooto
    If rsnooto.EOF Then
        Urutan = "B." + "000001"
        txtnota.Text = Urutan
    Else
        Hitung = Right(rsnooto!No_Masuk, 6) + 1
        Urutan = "B." + Right("00000" & Hitung, 6)
    End If
    txtnota.Text = Urutan
    End With
End Sub

Sub Konek_DB()
Set Koneksi = New ADODB.Connection
Set rspenjualan = New ADODB.Recordset
Set rsdatabarang = New ADODB.Recordset
Koneksi.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBrafacell.mdb"
End Sub

Private Sub cmdCari_click()
Unload Me
FrmCariBeli.Show
End Sub

Private Sub cmdKeluar_Click()
Unload Me
FrmMain.Show
End Sub

Private Sub Form_Load()
Call Konek_DB
Call oto1

 ' Menambah daftar pilihan pada ComboBox
  cmbSatuan.AddItem ("PCS")
  cmbSatuan.AddItem ("RTG")
  cmbSatuan.AddItem ("DUS")
  
DTPicker1.Value = Format(Now, "dd mm yyyy")
DTPicker1.Format = dtpCustom
DTPicker1.CustomFormat = "dd-MMMM-yyyy"
End Sub
Sub KondisiAwal()
    txtKode.Text = ""
    txtNama.Text = ""
    txtHarga.Text = "0"
    txtJumlah.Text = "0"
    cmbSatuan.Text = ""
End Sub
Private Sub txtKode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If KeyAscii = 13 Then
    If rsdatabarang.State = adStateOpen Then rsdatabarang.Close

    rsdatabarang.Open "SELECT * FROM MasterBarang where Kode_Barang = '" & txtKode.Text & "'", Koneksi, adOpenKeyset
    
    If Not rsdatabarang.EOF Then
        txtNama.Text = rsdatabarang!Nama_Barang
        cmbSatuan.Text = rsdatabarang!Satuan
        txtHarga.Text = rsdatabarang!HARGA_BELI
    Else
        MsgBox "Kode yang anda ketikkan salah", vbInformation, "TOKO RAFACELL"
    End If
    txtJumlah.SetFocus
End If
End Sub

Private Sub txtJumlah_Change()
On Error Resume Next
JumlahBeli = Val(txtJumlah.Text)
HargaBeli = Val(txtHarga.Text)
End Sub

Private Sub txtJumlah_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim LV As ListItem
    
    Set LV = ListView1.ListItems.Add(, , 0)
    With LV
        .SubItems(1) = txtKode.Text
        .SubItems(2) = txtNama.Text
        .SubItems(3) = txtHarga.Text
        .SubItems(4) = txtJumlah.Text
        .SubItems(5) = cmbSatuan.Text
        .SubItems(6) = Val(Text5.Text) * Val(Text6.Text)
    End With
    Call NoUrut
    Call Grandtotal
End If
End Sub

Private Sub Grandtotal()
total = 0
For i = 1 To ListView1.ListItems.Count
    total = total + Val(ListView1.ListItems.Item(i).SubItems(6))
Next i
    Label8.Caption = "Rp." & Format(total, "###,###,###")
End Sub

Private Sub cmdMasuk_Click()
Dim LV As ListItem
    Set LV = ListView1.ListItems.Add(, , 0)
    With LV
        .SubItems(1) = txtKode.Text
        .SubItems(2) = txtNama.Text
        .SubItems(3) = txtHarga.Text
        .SubItems(4) = txtJumlah.Text
        .SubItems(5) = cmbSatuan.Text
        .SubItems(6) = Val(txtJumlah.Text) * Val(txtHarga.Text)
    End With
    Call NoUrut
    Call Grandtotal
    Call KondisiAwal
End Sub

Private Sub cmdBatal_Click()
If ListView1.ListItems.Count > 0 Then
    ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
End If
Call KondisiAwal
Call Grandtotal
Call NoUrut
End Sub

Private Sub cmdBaru_Click()
If ListView1.ListItems.Count = 0 Then
    MsgBox "Belum ada transaksi untuk disimpan", vbCritical
Else
simpanbeli = "insert into Pembelian values('" & txtnota.Text & "'," & _
    "'" & Format(DTPicker1.Value, "YYYY/MM/DD") & "'," & _
    "'" & Label8.Caption & "')"
Koneksi.Execute simpanbeli
    
    
Dim rssimpandebeli As ADODB.Recordset
Set rssimpandebeli = New ADODB.Recordset
rssimpandebeli.Open "select * from Pembelian_Detail", Koneksi, adOpenStatic, adLockOptimistic
With ListView1
    For i = 1 To .ListItems.Count
    rssimpandebeli.AddNew
    rssimpandebeli.Fields("No_Masuk") = txtnota.Text
    rssimpandebeli.Fields("Tgl_Masuk") = Format(DTPicker1.Value, "YYYY/MM/DD")
    rssimpandebeli.Fields("Kode_Barang") = .ListItems(i).SubItems(1)
    rssimpandebeli.Fields("Nama_Barang") = .ListItems(i).SubItems(2)
    rssimpandebeli.Fields("Harga_Beli") = .ListItems(i).SubItems(3)
    rssimpandebeli.Fields("Jumlah_Barang") = .ListItems(i).SubItems(4)
    rssimpandebeli.Fields("Satuan") = .ListItems(i).SubItems(5)
    rssimpandebeli.Fields("Subtotal") = .ListItems(i).SubItems(6)
    rssimpandebeli.Update

    Dim Ubahstok As String
    Ubahstok = "update MasterBarang set Stok = Stok + " & Val(.ListItems(i).SubItems(4)) & ", Harga_beli= " & Val(.ListItems(i).SubItems(3)) & "" & _
    " where Kode_Barang = '" & .ListItems(i).SubItems(1) & "'"
    Koneksi.Execute Ubahstok
    Next i
End With
MsgBox "Transaksi Berhasil disimpan", vbInformation
Call kosongform
ListView1.ListItems.Clear
Call oto1
Call Grandtotal
End If
End Sub

Private Sub kosongform()
txtKode.Text = ""
txtNama.Text = ""
cmbSatuan.Text = ""
txtHarga.Text = ""
txtJumlah.Text = ""
txtKode.SetFocus
End Sub

Private Sub NoUrut()
Dim i As Integer

If ListView1.ListItems.Count > 0 Then
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Text = CStr(i)
    Next
End If
End Sub

