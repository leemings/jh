<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("yx8_mhjh_username")
if un="" then Response.Redirect "../../error.asp?id=016"
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
        response.write "禁止站外提交！"
        response.end
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hs")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 形态,体力,银两 from 用户 where 姓名='"&un&"'",conn
tl=rst("体力")
yl=rst("银两")
zt=rst("形态")
rst.Close
if zt<>"鬼魂" then 
set rst=nothing
conn.Close
set conn=nothing
Response.write "<script language='javascript'>alert ('你不是鬼魂，来这里只有死路一条');location.href='fh.asp';</script>"
Response.end
end if
if tl<500000 then
set rst=nothing
conn.Close
set conn=nothing
Response.write "<script language='javascript'>alert ('你的体力少于五十万');location.href='fh.asp';</script>"
Response.end
end if
if yl<500000 then 
set rst=nothing
conn.Close
set conn=nothing
Response.write "<script language='javascript'>alert ('你的银两少于五十万');location.href='fh.asp';</script>"
Response.end
end if
%><!--#include file="data.asp"--><%
sql="select 等级,类型,时间 from mj where 所有者='"&un&"' and 名称='鬼斧神功'"
set rs=connb.execute(sql)
if rs.EOF or rs.BOF then
rs.Close
connb.Execute "insert into mj(所有者,名称,等级,攻击,防御,特效,速度,时间,类型) values('"&un&"','鬼斧神功',1,0,0,'无',0,now(),'鬼')"
conn.execute "update 用户 set 体力=体力-100000,银两=银两-100000 where 姓名='"&un&"'"
msg="你修习“鬼斧神功”第1重成功。耗用银两十万,体力十万"
else
dj=rs("等级")
  sj=DateDiff("s",rs("时间"),now())
  rs.Close
if sj<0.1 then
	ss=70-sj
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]秒再操作！');location.href='fh.asp';}</script>"
	Response.End
end if
rst.Open "select id from 用户 where 姓名='"&un&"' and 精力>"&dj*10&" and 体力>="&dj*1000,conn
if rst.EOF or rst.BOF then
rst.Close
msg="你修习鬼斧神功第"&dj+1&"重时心有旁骛失败，无法专心入定，因而失败，需精力"&dj*10&",体力"&dj*1000
elseif dj=48 then
randomize timer
dd=int(rnd*9)+1
if dd>5 then
conn.Execute "update 用户 set 体力=体力-"&dj*1000&",精力=精力-"&dj*10&",形态='人',门派='无',身份='无' where 姓名='"&un&"'"
sql="delete * from mj where 所有者='"&un&"' and 名称='鬼斧神功' and 类型='鬼'"
Set Rs=connb.Execute(sql)
Response.write "<script language='javascript'>alert ('你修习鬼斧神功第七七四十九重成功，鬼斧神功化为紫气，使你转世为人，你可以去找你的仇人报仇了！');location.href='fh.asp';</script>"
Response.end
elseif dd>3 or dd<=5 then
conn.Execute "update 用户 set 体力=体力-"&dj*1000&",精力=精力-"&dj*10&",形态='魔',门派='无',身份='无' where 姓名='"&un&"'"
sql="delete * from mj where 所有者='"&un&"' and 名称='鬼斧神功' and 类型='鬼'"
Set Rs=connb.Execute(sql)
Response.write "<script language='javascript'>alert ('你修习鬼斧神功第七七四十九重失败,误入“魔道”,成为“魔”，魔有一些事情是做不了的啊，要想做人，就去出家，用佛法修炼自己，再重新做人吧！');location.href='fh.asp';</script>"
Response.end
else
conn.Execute "update 用户 set 体力=体力-"&dj*1000&",精力=精力-"&dj*10&",状态='死亡' where 姓名='"&un&"'"
sql="update mj set 等级=0,时间=now() where 所有者='"&un&"' and 名称='鬼斧神功' and 类型='鬼'"
Set Rs=connb.Execute(sql)
Response.write "<script language='javascript'>alert ('你修习鬼斧神功第七七四十九重失败,被阎罗再次打入十八层抵御，辛苦修炼的鬼斧神功化为虚有！');location.href='fh.asp';</script>"
Response.end
session.Abandon
end if
else
conn.Execute "update 用户 set 体力=体力-"&dj*1000&",精力=精力-"&dj*10&" where 姓名='"&un&"'"
sql="update mj set 等级=等级+1,时间=now() where 所有者='"&un&"' and 名称='鬼斧神功' and 类型='鬼'"
Set Rs=connb.Execute(sql)
msg="你修习“鬼斧神功”第"&dj+1&"重成功。耗用精力"&dj*10&",体力"&dj*1000
end if
end if
set rst=nothing
conn.Close
set conn=nothing
set rs=nothing
connb.Close
set connb=nothing
msg=replace(msg,"\","\\")
msg=replace(msg,"/","\/")
msg=replace(msg,chr(34),"\"&chr(34))
msg=replace(msg,chr(39),"\"&chr(39))
%>
<html>
<head>
<title>凤凰涅盘</title>
<link rel=stylesheet href='../../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='../../chatroom/bg1.gif'>
<div align=center>
<font color=0000ff size=4>凤凰涅盘</font>
<hr>
</div>
<p align="center"><%=msg%><br><br><input type=button value='返回' onclick="javascript:location.href='fh.asp'"></p>
</body>
</html>


