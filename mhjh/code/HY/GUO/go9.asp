
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
if Session("usepro")<>"4" then
Session("usepro")=""
response.redirect "index.asp"
end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT ���� FROM �û� where ����='" & username & "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "�㲻�ǽ�������"
conn.close
response.end
else
at=rs("����")
Randomize
at=int(at*rnd)
Randomize
at1=int(100000000*rnd)
if at>at1 then
Session("usepro")="5"
msg="����һ���භ������"&at&"������Ӯ����ƽ֮��"&at1&"��������<a href='go10.asp'>�������������𿪱��ص��Űɡ�</a>"
else
Session("usepro")=""
msg="����һ���භ������"&at&"������������ƽ֮��"&at1&"������"
Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
end if
%>
<html>

<head>
<title>������</title><LINK href="../style.css" rel=stylesheet></head>
<body background="../images/back2.gif" oncontextmenu=self.event.returnValue=false >
<div align="center"><%=msg%></div>
</html>
<%
end if
%>

