VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCari 
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7815
      Begin VB.TextBox txtCari 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Kategori Pencarian ] "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   4455
      Begin VB.OptionButton rbKategori 
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
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton rbKategori 
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
         Height          =   210
         Index           =   2
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Width           =   1785
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   4560
      TabIndex        =   0
      Top             =   960
      Width           =   3255
      Begin VB.CommandButton cmdNormal 
         Caption         =   "&Normal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "&Tutup"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   1305
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NO"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kode Barang"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama Barang"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Harga Jual"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Stok"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Satuan"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FrmCari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Koneksi As New ADODB.Connection
Dim rsdatabarang As New ADODB.Recordset
Dim rsbarangmasuk As New ADODB.Recordset
Dim rspemasok As New ADODB.Recordset
Dim rscustomer As New ADODB.Recordset
Dim rscaribarang As New ADODB.Recordset
Dim rspembelian As New ADODB.Recordset
Sub Konek_DB()
Set Koneksi = New ADODB.Connection
Set rspembelian = New ADODB.Recordset
Set rscaribarang = New ADODB.Recordset
Koneksi.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBrafacell.mdb"
End Sub

Private Sub cmdTutup_Click()
Unload Me
FrmTransaksi.Show
End Sub

Private Sub Form_Load()
Call Konek_DB
Call TampilData
Call NoUrut
End Sub
Private Sub NoUrut()
Dim i As Integer

If ListView1.ListItems.Count > 0 Then
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Text = CStr(i)
    Next
End If
End Sub
Private Sub TampilData()
If rscaribarang.State = adStateOpen Then rs.Close
rscaribarang.Open ("SELECT * FROM MasterBarang "), Koneksi, adOpenKeyset
rscaribarang.Requery
Dim lvw As ListItem
Do Until rscaribarang.EOF
Set lvw = ListView1.ListItems.Add(, , 0)
With lvw
    .SubItems(1) = rscaribarang!Kode_Barang
    .SubItems(2) = rscaribarang!Nama_Barang
    .SubItems(3) = rscaribarang!Harga_Jual
    .SubItems(4) = rscaribarang!STOK
    .SubItems(5) = rscaribarang!Satuan
End With
rscaribarang.MoveNext
Loop
End Sub

Private Sub txtCari_Change()
Call Konek_DB
  If rbKategori(0).Value = True Then
     Set rscaribarang = New ADODB.Recordset
     rscaribarang.Open "Select * from MasterBarang where Kode_Barang like '%" & txtCari & "%'", Koneksi, adOpenStatic

     With rscaribarang
     If Not .EOF Then
        i = 1
        ListView1.ListItems.Clear
        While Not .EOF
           Set View = ListView1.ListItems.Add
           View.Text = i
           View.SubItems(1) = !Kode_Barang
           View.SubItems(2) = !Nama_Barang
           View.SubItems(3) = !Harga_Jual
           View.SubItems(4) = !STOK
           View.SubItems(5) = !Satuan
           i = i + 1
           .MoveNext
       Wend
     End If
     End With
     rscaribarang.Close
  Else
     Set rscaribarang = New ADODB.Recordset
     rscaribarang.Open "Select * from MasterBarang where Nama_Barang like '%" & txtCari & "%'", Koneksi, adOpenStatic

     With rscaribarang
     If Not .EOF Then
        i = 1
        ListView1.ListItems.Clear
        While Not .EOF
           Set View = ListView1.ListItems.Add
           View.Text = i
           View.SubItems(1) = !Kode_Barang
           View.SubItems(2) = !Nama_Barang
           View.SubItems(3) = !Harga_Jual
           View.SubItems(4) = !STOK
           View.SubItems(5) = !Satuan
           i = i + 1
           .MoveNext
       Wend
     End If
     End With
     rscaribarang.Close
   End If
End Sub

Private Sub ListView1_DblClick()
    If rsdatabarang.State = adStateOpen Then rsdatabarang.Close
    rsdatabarang.Open "select * from MasterBarang where Kode_Barang = '" & ListView1.SelectedItem.SubItems(1) & "'", Koneksi, adOpenKeyset
    If Not rsdatabarang.EOF Then
    FrmTransaksi.txtKode.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
    FrmTransaksi.txtNama.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(2)
    FrmTransaksi.txtHarga.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(3)
    FrmTransaksi.cmbSatuan.Text = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(5)
    FrmTransaksi.Show
    Unload Me
    End If
End Sub

Private Sub cmdNormal_Click()
    Call Form_Load
    txtCari.Text = ""
    txtCari.SetFocus
End Sub
