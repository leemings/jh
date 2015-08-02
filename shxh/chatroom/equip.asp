<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
commodityname=Request.QueryString("commodityname")
if instr(commodityname,"'")<>0 or instr(commodityname," ")<>0 then Response.End
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "chaterror.asp?id=000"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select tc.ID,tc.属性,tc.攻击,tc.防御,tc.特效,tu.特技 from 物品 tc inner join 用户 tu on tu.姓名=tc.所有者 where tc.名称='"&commodityname&"' and tc.所有者='"&username&"' and tc.数量>0"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then
	msg="你没有此样物品可供装备"
else
	cmid=rst("id")
	cmtype=rst("属性")
	cmattack=rst("攻击")
	cmdefence=rst("防御")
	cmespecial=rst("特效")
	especial=rst("特技")
	rst.Close
	rst.Open "select id from 物品 where 属性='"&cmtype&"' and 装备=true and 所有者='"&username&"'",conn
	if rst.EOF or rst.BOF then
		if cmespecial="无" or isnull(cmespecial) or cmespecial="" then
			especial=especial
		else 
			especial=especial&cmespecial&";"
		end if	
		conn.Execute "update 用户 set 攻击=攻击+"&cmattack&",防御=防御+"&cmdefence&",特技='"&especial&"' where 姓名='"&username&"'"
		conn.Execute "update 物品 set 数量=数量-1,装备=True where id="&cmid
		msg="你成功的装备了"&commodityname&",攻击因些而增加"&cmattack&",防御因此而增加"&cmdefence
	else
		msg="错误，你已经装备有"&cmtype&",无法再行装备"&commodityname
	end if
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing	
%>
<head><title>装备物品</title><link rel="stylesheet" href="style1.css"><script language=javascript>setTimeout("location.replace('seeequip.asp');",3000);</script></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false >
<p align=center><font color=0000FF>装备物品</font></p>
<font color=FF0000><%=msg%></font>
<br><br><br>3秒钟后自动返回<br><a href="javascript:location.replace('seeequip.asp');" onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></body>