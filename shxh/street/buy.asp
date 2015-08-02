<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
commodityid=Request.QueryString("id")
if isnumeric(commodityid) then
	set conn=server.CreateObject("adodb.connection")
	conn.Open Application("Ba_jxqy_connstr")
	set rst=server.CreateObject("adodb.recordset")
	sqlstr="select * from 商店 where id="&commodityid
	rst.Open sqlstr,conn
	if rst.EOF or rst.BOF then
		msg="错误，不出售此物品！"
	else
		commodityname=rst("名称")
		commoditytype=rst("属性")
		commodityhealth=rst("体力")
		magic=rst("内力")
		attack=rst("攻击")
		defence=rst("防御")
		especial=rst("特效")
		commodityprice=rst("价格")
		rst.Close
		rst.Open "select 银两 from 用户 where 姓名='"&username&"' and 银两>="&commodityprice,conn
		if rst.EOF or rst.BOF then
			msg="错误，你的钱好象没有带够哟！"
		else
			rst.Close
			sqlstr="select * from 物品 where 名称='"&commodityname&"' and 所有者='"&username&"'"
			rst.Open sqlstr,conn
			if rst.EOF or rst.BOF then
				sqlstr="insert into 物品(名称,属性,体力,内力,攻击,防御,特效,价格,数量,所有者,寄售,装备) values('"&commodityname&"','"&commoditytype&"',"&commodityhealth&","&magic&","&attack&","&defence&",'"&especial&"',"&commodityprice/2&",1,'"&username&"',False,False)"
			else
				sqlstr="update 物品 set 数量=数量+1 where 名称='"&commodityname&"' and 所有者='"&username&"'"
			end if
			conn.BeginTrans
			conn.Execute sqlstr
			conn.Execute "update 用户 set 银两=银两-"&commodityprice&" where 姓名='"&username&"'"
			conn.CommitTrans
			msg="承惠纹银"&commodityprice&"两，这是"&commodityname&"，您收好了！"
		end if	
	end if
	rst.Close
	set rst=nothing
	conn.Close
	set conn=nothing	
else
	msg="错误，不出售此物品！"
end if
%>
<head><title>商店</title><LINK href="../chatroom/css.css" rel=stylesheet><script language=javascript>setTimeout('history.back();',3000);</script></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false topMargin=200>
<div align=center>3秒钟后自动返回<br>
  <font color=FF0000><img src="../images/error.gif" width="32" height="32"> <%=msg%></font></div>
<p align=right><a href="#" onclick=javascript:history.back(); onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>