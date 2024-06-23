VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------
'AccessDB�t�@�C���ւ̐ڑ��ƃN�G�����s��S���N���X
'------------------------------


Option Explicit

Private conn As ADODB.Connection

' �f�[�^�x�[�X�ɐڑ�
Public Sub OpenDatabase(dbPath As String)
    Set conn = New ADODB.Connection
    conn.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    conn.Open
End Sub

' �N�G�������s���Č��ʂ�Ԃ�
Public Function ExecuteQuery(sql As String) As ADODB.Recordset
    Set ExecuteQuery = conn.Execute(sql)
End Function

' �f�[�^�x�[�X�ڑ������
Public Sub CloseDatabase()
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Sub