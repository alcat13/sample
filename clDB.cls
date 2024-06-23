VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------------------
'必要参照設定：
'Microsoft ActiveX Data Objects x.x Library
'
'コメント：
'AccessDBファイルへの接続とクエリ実行を担うクラス
'------------------------------

Option Explicit

Private conn As ADODB.Connection

' データベースに接続
Public Sub OpenDatabase(dbPath As String)
    Set conn = New ADODB.Connection
    conn.connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    conn.Open
End Sub

' クエリを実行して結果を返す
Public Function ExecuteQuery(sql As String) As ADODB.Recordset
    Set ExecuteQuery = conn.Execute(sql)
End Function

' データベース接続を閉じる
Public Sub CloseDatabase()
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Sub
