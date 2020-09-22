VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Create DAO Database From Scratch"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Check if DB exists"
      Height          =   525
      Left            =   2550
      TabIndex        =   11
      Top             =   3990
      Width           =   2625
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   525
      Left            =   810
      TabIndex        =   10
      Top             =   3990
      Width           =   1245
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   810
      TabIndex        =   9
      Top             =   3300
      Width           =   4695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   810
      TabIndex        =   7
      Top             =   2640
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   810
      TabIndex        =   5
      Top             =   1980
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   330
      TabIndex        =   3
      Top             =   1320
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   330
      TabIndex        =   1
      Top             =   420
      Width           =   5175
   End
   Begin VB.Label Label7 
      Caption         =   "The database that you create will be created in the same folder that the program is located."
      Height          =   1005
      Left            =   570
      TabIndex        =   13
      Top             =   4560
      Width           =   1725
   End
   Begin VB.Label Label6 
      Caption         =   $"Form1.frx":0000
      Height          =   1155
      Left            =   2550
      TabIndex        =   12
      Top             =   4590
      Width           =   3225
   End
   Begin VB.Label Label5 
      Caption         =   "Field Three:"
      Height          =   525
      Left            =   810
      TabIndex        =   8
      Top             =   3060
      Width           =   1905
   End
   Begin VB.Label Label4 
      Caption         =   "Field Two:"
      Height          =   525
      Left            =   810
      TabIndex        =   6
      Top             =   2400
      Width           =   1905
   End
   Begin VB.Label Label3 
      Caption         =   "Field One:"
      Height          =   525
      Left            =   810
      TabIndex        =   4
      Top             =   1740
      Width           =   1905
   End
   Begin VB.Label Label2 
      Caption         =   "Table Name:"
      Height          =   525
      Left            =   330
      TabIndex        =   2
      Top             =   1080
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "Database Name:"
      Height          =   525
      Left            =   330
      TabIndex        =   0
      Top             =   180
      Width           =   1905
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************
'*****************************************************
'You have to Reference the DAO Object Library and
'the Microsoft ActiveX Data Objects Library
'in order for this program to work.
'*****************************************************
'This program just shows you how to create a database
'using code instead of using the Visual Data Manager.
'*****************************************************
'You only need to Reference the Microsoft ActiveX
'Data Objects if you are going to use the "Check if
'Database Exists" function of the program.
'Otherwise just Reference the DAO Objects.
'*****************************************************

Dim db As Database
Dim td As TableDef, TempTd As TableDef
Dim fields(3) As Field, indexfield As Field
'If you want there to be more fields then change the number in the parenthesis above.
Dim dbindex As Index

Private Sub Command1_Click()
Set db = DBEngine.Workspaces(0).CreateDatabase(App.Path + "\" + Text1.Text, dbLangGeneral) 'creates the database
Set td = db.CreateTableDef(Text2.Text) 'creates the table for the fields
    Set fields(0) = td.CreateField(Text3.Text, dbText) 'defines the fields
    Set fields(1) = td.CreateField(Text4.Text, dbText)
    Set fields(2) = td.CreateField(Text5.Text, dbText)
    'If you add more fields, also create more of these lines putting their respective numbers in the parenthesis.
    
    td.fields.Append fields(0)
    td.fields.Append fields(1)
    td.fields.Append fields(2)
    'Here also, for more fields
    
    Set dbindex = td.CreateIndex(Text3.Text & "index") 'creates an index for the fields
    Set indexfield = dbindex.CreateField(Text3.Text)
    dbindex.fields.Append indexfield
    td.Indexes.Append dbindex
    db.TableDefs.Append td
End Sub

Private Sub Command2_Click()
On Error GoTo HandleIT
Dim strConnect As String
Dim strProvider As String
Dim strDataSource As String
Dim strDataBaseName As String
strProvider = "Provider= Microsoft.Jet.OLEDB.3.51;"
strDataSource = App.Path 'This tells the program to look in the App's Path for the DB
strDataBaseName = "\" + Text1 + ";" 'Whatever the database name may be.
strDataSource = "Data Source=" & strDataSource & strDataBaseName
strConnect = strProvider & strDataSource
Set connConnection = New ADODB.Connection
connConnection.Open strConnect
MsgBox "This Database does exist.", vbOKOnly, "Attention!"
Exit Sub

HandleIT:
MsgBox "Database does not exist.  Please create it first.", vbOKOnly, "Attention!"
'If it doesn't exist then it will tell you!
End Sub

Private Sub Form_Load()

End Sub
