<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("yx8_mhjh_username")
if un="" then Response.Redirect "../../error.asp?id=016"
mg=Request.QueryString("mg")
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") or instr(elevalue,"?") or instr(elevalue,"&") or instr(elevalue,"<")  or instr(elevalue,">") or instr(elevalue,"or")  then Response.Redirect "../../error.asp?id=056"
next
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
if not(mg="我佛慈悲" or mg="佛光普照" or mg="万佛朝宗") then Response.Redirect "../../error.asp?id=024"
co=Session("yx8_mhjh_usercorp")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 门派,内力,银两 from 用户 where 姓名='"&un&"'",conn
mp=rst("门派")
nl=rst("内力")
yl=rst("银两")
if mp<>"出家" then 
set rst=nothing
conn.Close
set conn=nothing
Response.write "<script language='javascript'>alert ('你还不是出家人');location.href='hs2.asp';</script>"
Response.end
end if
if nl<10000 then
set rst=nothing
conn.Close
set conn=nothing
Response.write "<script language='javascript'>alert ('你的内力少于一万');location.href='hs2.asp';</script>"
Response.end
end if
if yl<10000 then 
set rst=nothing
conn.Close
set conn=nothing
Response.write "<script language='javascript'>alert ('你的银两少于一万');location.href='hs2.asp';</script>"
Response.end
end if
%><!--#include file="data.asp"--><%
sql="select 等级,攻击,防御,类型,时间 from mj where 所有者='"&un&"' and 名称='"&mg&"'"
set rs=connb.execute(sql)
if rs.EOF or rs.BOF then
rs.Close
connb.Execute "insert into mj(所有者,名称,等级,攻击,防御,特效,速度,时间,类型) values('"&un&"','"&mg&"',1,1000,1000,'无',1,now(),'佛')"
conn.execute "update 用户 set 内力=内力-1000,银两=银两-1000 where 姓名='"&un&"'"
msg="你修习"&mg&"第1级成功。耗用银两1000,内力1000"
else
dj=rs("等级")
gj=rs("攻击")
fy=rs("防御")
  sj=DateDiff("s",rs("时间"),now())
if sj<140 then
	ss=140-sj
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]秒再操作！');location.href='hs2.asp';}</script>"
	Response.End
end if
rst.Close
rs.Close
if gj>2000000000 or fy>2000000000 then 
Response.write "<script language='javascript'>alert ('你的这招禅功的威力已经大于20亿了，无法继续修炼');location.href='hs2.asp';</script>"
Response.end
end if
if mg="我佛慈悲" then
gs=10000
elseif mg="佛光普照" then
gs=20000
else
gs=30000
end if
rst.Open "select id from 用户 where 姓名='"&un&"' and 内力>="&dj*gs,conn
if rst.EOF or rst.BOF then
rst.Close
msg="你修习"&mg&"第"&dj+1&"级时心有旁骛失败，无法专心入定，因而失败，需内力"&dj*gs
else
conn.Execute "update 用户 set 内力=内力-"&dj*gs&",银两=银两-1000 where 姓名='"&un&"'"

sql="update mj set 等级=等级+1,攻击=攻击*1.5,防御=防御*1.5,时间=now() where 所有者='"&un&"' and 名称='"&mg&"' and 类型='佛'"
Set Rs=connb.Execute(sql)
msg="你修习"&mg&"第"&dj+1&"级成功。耗用银两1000,内力"&dj*gs
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
<title>修习禅功</title>
<link rel=stylesheet href='../../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='../../chatroom/bg1.gif'>
<div align=center>
<font color=0000ff size=4>修习禅功</font>
<hr>
</div>
<p align="center"><%=msg%><br><br><input type=button value='返回' onclick="javascript:location.href='hs2.asp'"></p>
</body>
</html>

