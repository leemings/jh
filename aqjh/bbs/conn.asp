<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Response.Buffer=True
IsSqlDataBase=0      '�������ݿ����0ΪAccess���ݿ⣬1ΪSQL���ݿ�
If IsSqlDataBase=0 Then
'''''''''''''''''''''''''''''' Access���ݿ� '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
datapath    ="database/"      '���ݿ�Ŀ¼�����·��
datafile    ="aqjh_bbs.asp"      '���ݿ���ļ���
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(""&datapath&""&datafile&"")
'Connstr="DBQ="&server.mappath(""&datapath&""&datafile&"")&";DRIVER={Microsoft Access Driver (*.mdb)};"
SqlNowString="Now()"
SqlChar="'"
ver="5.15"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Else
'''''''''''''''''''''''''''''' SQL���ݿ� ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SqlLocalName   ="(local)"     '����IP  [ ������ (local) �����IP ]
SqlUsername    ="sa"          '�û���
SqlPassword    ="1"           '�û�����
SqlDatabaseName="bbsxp"       '���ݿ���
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
Response.Write "���ݿ����ӳ������������ִ���"
Response.End
End If
On Error GoTo 0
%>