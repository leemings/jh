<!--#include file="data2.asp"-->
<%
if session("yx8_mhjh_username")="" then
response.redirect"warning.htm"
else
sheepname=request.form("sheepname")
if sheepname="" then
response.redirect"Warning.htm"
else
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&session("yx8_mhjh_username")&"'",conn,1,1
if rs.bof then
rs.close
conn.close
response.redirect"Warning.htm"
else
rs.close
rs.open"select * from rules ",conn,1,1
if rs.bof then
rs.close
conn.close
response.redirect"Warning.htm"
else
cleandel=rs("cleandel")
cleanplusclean=rs("cleanplusclean")
cleanpluslife=rs("cleanpluslife")
rs.close
conn.close 
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")

rs.open"select ���� from �û� where ����='"&session("yx8_mhjh_username")&"'",conn,1,1
tempsplosh=rs("����")-cleandel
if tempsplosh<0 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"��û���㹻��Ǯ����ȥ׬��Ǯ�ɡ�",0,"FLUSH"
history.back
</script>

<%
else%><!--#include file="data2.asp"-->

<%rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&session("yx8_mhjh_username")&"'",conn,1,1
feeddate=rs("feeddate")
workload=rs("workload")
clean=rs("clean")+cleanplusclean
if clean>100 then
clean=100
end if
life=rs("life")+cleanpluslife
if life>100 then
life=100
end if
rs.close
if datediff("d",feeddate,date)<>0 then 
tempdate=date
conn.execute"update sheep set life='"&life&"',clean='"&clean&"',feeddate='"&date&"',feedsheepday=feedsheepday+1,workload='1' where sheepname='"&sheepname&"' and user='"&session("yx8_mhjh_username")&"'"
conn.close
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rs.open"select ���� from �û� where ����='"&session("yx8_mhjh_username")&"'",conn,1,1
conn.execute"update �û� set ����='"&tempsplosh&"' where ����='"&session("yx8_mhjh_username")&"'"

%>
<script language="vbscript">
msgbox"��ϴ��ϣ�",0,"FLUSH"
history.back
</script>

<%
else
if workload>=3 then
conn.close
%>
<script language="vbscript">
msgbox"���Ѿ�ά��С�����Σ����������ɡ�:-)",0,"FLUSH"
history.back
</script>
<%
else
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
conn.execute"update �û� set ����='"&tempsplosh&"' where ����='"&session("yx8_mhjh_username")&"'"
conn.close
%>
<!--#include file="data1.asp"-->

<%
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&session("yx8_mhjh_username")&"'",conn,1,1
conn.execute"update sheep set life='"&life&"',clean='"&clean&"',workload=workload+1 where sheepname='"&sheepname&"' and user='"&session("yx8_mhjh_username")&"'"
rs.close
conn.close
%>
<script language="vbscript">
msgbox"��ϴ��ϣ�",0,"FLUSH"
history.back
</script>
<%
end if
end if
end if
end if
end if
end if
end if
%>