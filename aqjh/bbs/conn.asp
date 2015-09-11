<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Response.Buffer=True
IsSqlDataBase=0      '定义数据库类别，0为Access数据库，1为SQL数据库
If IsSqlDataBase=0 Then
'''''''''''''''''''''''''''''' Access数据库 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
datapath    ="database/"      '数据库目录的相对路径
datafile    ="aqjh_bbs.asp"      '数据库的文件名
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(""&datapath&""&datafile&"")
'Connstr="DBQ="&server.mappath(""&datapath&""&datafile&"")&";DRIVER={Microsoft Access Driver (*.mdb)};"
SqlNowString="Now()"
SqlChar="'"
ver="5.15"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Else
'''''''''''''''''''''''''''''' SQL数据库 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SqlLocalName   ="(local)"     '连接IP  [ 本地用 (local) 外地用IP ]
SqlUsername    ="sa"          '用户名
SqlPassword    ="1"           '用户密码
SqlDatabaseName="bbsxp"       '数据库名
ConnStr = "Provider=Sqloledb;User ID=" & SqlUsername & "; Password=" & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source=" & SqlLocalName & ";"
SqlNowString="GetDate()"
ver="5.15 SQL"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
END IF
On Error Resume Next
Set conn=Server.CreateObject("ADODB.Connection")
conn.open ConnStr
If Err Then
err.Clear
Set Conn = Nothing
Response.Write "数据库连接出错，请检查连接字串。"
Response.End
End If
On Error GoTo 0
%>