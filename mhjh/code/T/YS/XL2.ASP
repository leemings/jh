<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("yx8_mhjh_username")
mg=Request.QueryString("mg")
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") or instr(elevalue,"?") or instr(elevalue,"&") or instr(elevalue,"<")  or instr(elevalue,">") or instr(elevalue,"or")  then Response.Redirect "../../error.asp?id=056"
next
if un="" then Response.Redirect "../../error.asp?id=016"
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
        response.write "禁止站外提交！"
        response.end
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"ys")=0 then 
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
if zt="鬼魂" then 
set rst=nothing
conn.Close
set conn=nothing
Response.write "<script language='javascript'>alert ('你是鬼魂，来这里只有死路一条');parent.history.go(-1);</script>"
Response.end
end if
if tl<500000 then
set rst=nothing
conn.Close
set conn=nothing
Response.write "<script language='javascript'>alert ('你的体力少于五十万');parent.history.go(-1);</script>"
Response.end
end if
if yl<500000 then 
set rst=nothing
conn.Close
set conn=nothing
Response.write "<script language='javascript'>alert ('你的银两少于五十万');parent.history.go(-1);</script>"
Response.end
end if
%><!--#include file="data.asp"--><%
sql="select 等级,类型,时间,攻击,防御,状态 from mj where 所有者='"&un&"' and 名称='"&mg&"' and 类型='轩'"
set rs=connb.execute(sql)
if rs.EOF or rs.BOF then
Response.write "<script language='javascript'>alert ('你没有这种武器');parent.history.go(-1);</script>"
set rst=nothing
conn.Close
set conn=nothing
set rs=nothing
connb.Close
set connb=nothing
Response.end
rs.Close
else
dj=rs("等级")
gj=rs("攻击")
fy=rs("防御")
zztt=rs("状态")
  sj=DateDiff("s",rs("时间"),now())
  rs.Close
if sj<0.1 then
	ss=70-sj
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]秒再操作！');location.href='xl.asp';}</script>"
set rst=nothing
conn.Close
set conn=nothing
set rs=nothing
connb.Close
set connb=nothing
	Response.End
end if
if dj>100 or gj>5000000 or fy>5000000 then
	Response.Write "<script language=JavaScript>{alert('你的要塞武器已经练到威力最高程度，无法继续');location.href='xl.asp';}</script>"
set rst=nothing
conn.Close
set conn=nothing
set rs=nothing
connb.Close
set connb=nothing
	Response.End
end if
rst.Open "select id from 用户 where 姓名='"&un&"' and 精力>"&dj*30&" and 体力>="&dj*1000,conn
if rst.EOF or rst.BOF then
rst.Close
msg="你修习"&mg&"第"&dj+1&"级时心有旁骛失败，无法专心入定，因而失败，需精力"&dj*100&"，精力"&dj*30&",体力"&dj*1000
else
if zztt=False then lx="防御"
if zztt=True then lx="攻击"
conn.Execute "update 用户 set 体力=体力-"&dj*1000&",精力=精力-"&dj*30&" where 姓名='"&un&"'"
sql="update mj set 等级=等级+1,"&lx&"="&lx&"*1.2,时间=now(),速度=速度+1 where 所有者='"&un&"' and 名称='"&mg&"' and 类型='轩'"
Set Rs=connb.Execute(sql)
msg="你修习“"&mg&"”第"&dj+1&"级成功。耗用精力"&dj*30&",体力"&dj*1000
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
<title>修炼要塞武器</title>
<link rel=stylesheet href='../../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='../../chatroom/bg1.gif'>
<div align=center>
<font color=0000ff size=4>修炼要塞武器</font>
<hr>
</div>
<p align="center"><%=msg%><br><br><input type=button value='返回' onclick="javascript:location.href='xl.asp'"></p>
</body>
</html>



