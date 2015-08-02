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
sqlstr="select * from 物品 where id="&cmid&" and 数量>0 and 所有者='"&username&"'"
rst.open sqlstr,conn
if  rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=028"
cmname=rst("名称")
rst.Close
sqlstr="select * from 商店 where 名称='"&cmname&"'"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=033"
cmprice=clng(rst("价格")/2)
rst.Close
set rst=nothing 	
if cmprice<1 then cmprice=1
conn.BeginTrans
conn.Execute "update 物品 set 数量=数量-1 where id="&cmid&" and 所有者='"&username&"'"
conn.CommitTrans
conn.Execute "update 用户 set 银两=银两+"&cmprice&" where 姓名='"&username&"'"
conn.Close
set conn=nothing
%>
<head><title>当铺</title><link rel="stylesheet" href="../system/style.css"><script language=javascript>setTimeout("location.replace('pawnshop.asp');",3000);</script></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color=0000FF size=6 face='幼圆'><%=Application("Ba_jxqy_systemname")%>当铺</font></p>
<p><font  size=4>你的货物我们收下了，这里是<%=cmprice%>两银子您 收好了</font></p>
<p align=right><a href="javascript:location.replace('pawnshop.asp');" onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>