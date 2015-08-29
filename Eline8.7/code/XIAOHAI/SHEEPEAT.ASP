<!--#include file="data2.asp"-->
<%
if session("sjjh_name")="" then
response.redirect"warning.htm"
else
sheepname=request.form("sheepname")
if sheepname="" then
response.redirect"Warning.htm"
else
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&session("sjjh_name")&"'",conn,1,1
if rs.bof then
rs.close
cn.close
response.redirect"Warning.htm"
else
rs.close
rs.open"select * from rules ",conn,1,1
if rs.bof then
rs.close
conn.close
response.redirect"Warning.htm"
else
eatdel=rs("eatdel")
eatplushungry=rs("eatplushungry")
eatplusmilk=rs("eatplusmilk")
rs.close
set conn=server.createobject("adodb.connection") 
conn.open Application("sjjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
rs.open"select 银两 from 用户 where 姓名='"&session("sjjh_name")&"'",conn,1,1
tempsplosh=rs("银两")-eatdel
if tempsplosh<0 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"您没有足够的钱！先去赚点钱吧。",0,"FLUSH"
history.back
</script>
<%
else
rs.close%>
<!--#include file="data2.asp"-->
<%
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&session("sjjh_name")&"'",conn,1,1
feeddate=rs("feeddate")
workload=rs("workload")
hungry=rs("hungry")+eatplushungry
if hungry>100 then
hungry=100
end if
milk=rs("milk")+eatplusmilk
rs.close
if datediff("d",feeddate,date)<>0 then 
tempdate=date
conn.execute"update sheep set milk='"&milk&"',hungry='"&hungry&"',feeddate='"&date&"',feedsheepday=feedsheepday+1,workload='1' where sheepname='"&sheepname&"' and user='"&session("sjjh_name")&"'"

conn.close
set conn=server.createobject("adodb.connection") 
conn.open Application("sjjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
conn.execute"update 用户 set 银两='"&tempsplosh&"' where 姓名='"&session("sjjh_name")&"'"
%>
<script language="vbscript">
msgbox"喂养完毕！",0,"FLUSH"
history.back
</script>
<%rs.close
conn.close%>
<!--#include file="data1.asp"-->

<%
else
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&session("sjjh_name")&"'",conn,1,1

if workload>=3 then
conn.close
%>
<script language="vbscript">
msgbox"您已经维护小孩三次！明天再来吧。:-)",0,"FLUSH"
history.back
</script>
<%
else
conn.execute"update sheep set milk='"&milk&"',hungry='"&hungry&"',workload=workload+1 where sheepname='"&sheepname&"' and user='"&session("sjjh_name")&"'"
rs.close
conn.close
set conn=server.createobject("adodb.connection") 
conn.open Application("sjjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
conn.execute"update 用户 set 银两='"&tempsplosh&"' where 姓名='"&session("sjjh_name")&"'"
conn.close
%>
<script language="vbscript">
msgbox"喂养完毕！",0,"FLUSH"
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