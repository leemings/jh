<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
cmid=Request.QueryString("id")
if not isnumeric(cmid) then Response.Redirect "../error.asp?id=024"
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from ��Ʒ where id="&cmid&" and ����>0 and ������='"&username&"'"
rst.open sqlstr,conn
if  rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=028"
cmname=rst("����")
rst.Close
sqlstr="select * from �̵� where ����='"&cmname&"'"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=033"
cmprice=clng(rst("�۸�")/2)
rst.Close
set rst=nothing 	
if cmprice<1 then cmprice=1
conn.BeginTrans
conn.Execute "update ��Ʒ set ����=����-1 where id="&cmid&" and ������='"&username&"'"
conn.CommitTrans
conn.Execute "update �û� set ����=����+"&cmprice&" where ����='"&username&"'"
conn.Close
set conn=nothing
%>
<head><title>����</title><link rel="stylesheet" href="../system/style.css"><script language=javascript>setTimeout("location.replace('pawnshop.asp');",3000);</script></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color=0000FF size=6 face='��Բ'><%=Application("Ba_jxqy_systemname")%>����</font></p>
<p><font  size=4>��Ļ������������ˣ�������<%=cmprice%>�������� �պ���</font></p>
<p align=right><a href="javascript:location.replace('pawnshop.asp');" onmouseover="window.status='����';return true;" onmouseout="window.status='';return true;">����</a></p>
</body>