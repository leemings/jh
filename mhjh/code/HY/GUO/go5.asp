<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
if Session("usepro")<>"2" then
Session("usepro")=""
response.redirect "index.asp"
end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT ���� FROM �û� where ����='" & username & "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "�㲻�ǽ�������"
conn.close
response.end
else
at=rs("����")
randomize timer
r=int(rnd*10)
s=int(rnd*50)
at=at*s
randomize timer
r=int(rnd*10)
s=int(rnd*50000)
at1=100*s
if at>at1 then
Session("usepro")="3"
msg="����һ���භ������"&at&"�㹥��Ӯ�˴�����"&at1&"�㹥����<a href='go6.asp'>�����ﴳ��һ�ذɡ�</a>"
else
Session("usepro")=""
msg="����һ���භ������"&at&"�㹥�����д�����"&at1&"�㹥����<a href='javascript:self.close()'>�������ȥ�ɰɡ�</a>"
end if
%>
<html>

<head>
<title>������</title>
<LINK href="../../style.css" rel=stylesheet>
</head><body bgcolor="#000000" oncontextmenu=self.event.returnValue=false text="#FFFFFF" >
<div align="center"><b><%=msg%></b></div></body>
</html>
<%
end if
conn.close
%>
