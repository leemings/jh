<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../error.asp?id=486"
nickname=Session("aqjh_name")
grade=Int(Session("aqjh_grade"))
if grade<>10 then Response.Redirect "manerr.asp?id=100"
Const Jet_Conn_Partial = "Provider=Microsoft.Jet.OLEDB.4.0;Data source="
Dim strDatabase, strFolder, strFileName
xp="aqjh.asp"
dataxp="../aqjh_data/"&xp
strFolder = Server.MapPath(dataxp)
strFolder = Replace(strFolder,xp,"")
strDatabase =xp
Private Sub dbCompact()
Dim SourceConn
Dim DestConn
Dim oJetEngine
SourceConn = Jet_Conn_Partial & strFolder & strDatabase
DestConn = Jet_Conn_Partial & strFolder & date() &"ѹ������"& strDatabase
mymdb=strFolder & strDatabase
tomdb=strFolder & date() &"ѹ������"& strDatabase
If DbExists(DestConn ) Then 
	Response.Write ("�Բ������ݿ�:"& tomdb &"�Ѿ�����,ʹ��Ftpɾ�����ԣ�!") 
	Response.End
end if
Set oJetEngine = Server.CreateObject("JRO.JetEngine")
With oJetEngine
	.CompactDatabase SourceConn, DestConn
End With
Response.Write ("���ݿ�:<font color=blue>" & mymdb & "</font>ѹ�����!<br>ѹ���ļ��ļ���<font color=blue>"&tomdb&"</font><br><br>���½ftp��ɾ��ԭʼ���ݿ⣺<font color=blue>"&mymdb&"</font><br>�ٽ�ѹ�����ݿ⣺<font color=blue>"&tomdb&"</font><br>�����ɣ�<font color=blue>"&mymdb&"</font>ʹ�ã���")
End Sub
Public function DbExists(byVal dbPath) 
'�������ݿ��ļ��Ƿ���� 
On Error resume Next 
Dim c 
Set c = Server.CreateObject("ADODB.Connection") 
c.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath 
If Err.number<>0 Then 
Err.Clear 
DbExists = false 
else 
DbExists = True 
End If 
set c = nothing 
End function 

%>
<%
' Compact database and tell the user the database is optimized
Select Case Request.form("cmd")
	Case "ѹ��"
		call dbCompact 
End Select
%>
<LINK href=css/css.css type=text/css rel=stylesheet>
<title><%=Application("aqjh_chatroomname")%>ά������</title>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><b><font color=blue>�û�����ѹ�����޸�ϵͳ</b></p>
<form method="POST" action="">
<p align="center">
<input type="submit" value="ѹ��" name="cmd">
</p>
</form>

<p align="center"><font color="#000000">��ѡ����ѹ�������ݿ⣬��ѹ����ʼ���У�</font><br>
  ��ѹ�����ݿ����<font color="#0000FF"><b>��Fsoϵͳ </b></font>��Ҫʹ��Ftp�ϴ��������ɣ�<br>
  <font color="#FF0000"><b>ǿ��Ҫ�󱸷�ԭʼ���ݿ��ļ�������һ�к���Ը���<br>
  </b></font><br>ִ�д˲�����Ҫ��IE��ȫ���óɵͲſ���ִ�У�</p>