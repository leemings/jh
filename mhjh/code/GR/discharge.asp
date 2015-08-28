<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
cmtype=Request.QueryString("cmtype")
if instr(cmtype,"'")<>0 or instr(cmtype," ")<>0 then Response.End
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "chaterror.asp?id=000"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select ID,名称,攻击,防御,特效 from 物品 where 属性='"&cmtype&"' and 所有者='"&username&"' and 装备=True"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then
msg="错误，你没有装备此样物品，所以无法卸载！"
else
cmid=rst("id")
cmname=rst("名称")
cmattack=rst("攻击")
cmdefence=rst("防御")
cmespecial=rst("特效")
rst.Close
rst.Open "select * from 用户 where 姓名='"&username&"'  and 攻击>"&cmattack&" and 防御>"&cmdefence,conn
if rst.EOF or rst.BOF then
msg="错误，能力不足，不能卸下"&cmname
else
if cmespecial="无" then
especial=rst("特技")
else
especial=replace(rst("特技"),";"&cmespecial&";",";",1,1)
end if
conn.BeginTrans
conn.Execute "update 用户 set 攻击=攻击-"&cmattack&",防御=防御-"&cmdefence&",特技='"&especial&"' where 姓名='"&username&"'"
conn.Execute "update 物品 set 数量=数量+1,装备=False where id="&cmid
if conn.Errors.Count>0 then
conn.RollbackTrans
Response.Redirect "../error.asp?id=104&errormsg="&conn.Errors(0).Description
else
conn.CommitTrans
msg="你卸下了"&cmname&",攻击因些而减少"&cmattack&",防御因此而减少"&cmdefence
end if
end if
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><title>商店</title><link rel="stylesheet" href="../chatroom/css.css"><script language=javascript>setTimeout("location.replace('seeequip.asp');",3000);</script></head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false>
<p align=center><font color=0000FF >卸下装备</font></p>
<p align="center">
<font color=FF0000><%=msg%></font>
<br><br><br>3秒钟后自动返回<br><a href="javascript:location.replace('seeequip.asp');" onmouseover="window.status='返回';return true;" onmouseout="window.status='';return true;">返回</a></body>
