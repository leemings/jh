<!--#include file="data.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")                                  
conn.Open Application("yx8_mhjh_connstr")                                  
set rs=server.CreateObject("adodb.recordset")                                  
rs.Open "select 性别 from 用户 where 姓名='"&username&"'",conn
sex=rs("性别") 
if sex="男" then
sex="她"
else
sex="他"
end if
rs.Close                                  
set rs=nothing  
opt=Request.QueryString("opt")
if opt="" then opt="洞房"
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=clng(acpage)
if acpage<1 then acpage=1
msg="<head><title>洞房</title><link rel=stylesheet href='../../style.css'></head><body background='../../chatroom/bg1.gif' oncontextmenu='self.event.returnValue=false;'><table width=40% border='1' cellspacing='0' cellpadding='1' align=center bgcolor=FFB442><tr align=center>"
if opt="洞房" then
msg=msg&"<td bgcolor=ffffff><a href='marriage.asp?opt=洞房' onmouseover="&chr(34)&"window.status='向你的心上人坦露心声吧，主会祝福你！';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title='向你的心上人坦露心声吧，主会祝福你！'><font color=#000000>洞房</font></a></td><td><a href='marriage.asp?opt=分房' onmouseover="&chr(34)&"window.status='主会原谅你的，我可怜的孩子！';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title='主会原谅你的，我可怜的孩子！'><font color=#000000>分房</font></a></td></tr></table><p align=center></p><table width=100% border='1' cellspacing='0' cellpadding='1'><tr bgcolor=ffffff align=center><td>性需求者</td><td>性爱对象</td><td>性爱蜜语</td><td>操作</td></td></tr>"
sqlstr="select  * from 洞房 where 操作=False order by 时间 desc"
 Set Rst=connt.Execute(sqlstr)
else
msg=msg&"<td><a href='marriage.asp?opt=洞房' onmouseover="&chr(34)&"window.status='向你的心上人坦露心声吧，主会祝福你！';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title='向你的心上人坦露心声吧，主会祝福你！'><font color=#000000>洞房</font></a></td><td bgcolor=ffffff><a href='marriage.asp?opt=分房' onmouseover="&chr(34)&"window.status='主会原谅你的，我可怜的孩子！';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title='主会原谅你的，我可怜的孩子！'><font color=#000000>分房</font></a></td></tr></table><p align=center><b></b></p><table width=100% border='1' cellspacing='0' cellpadding='1'><tr bgcolor=ffffff align=center><td>申请人</td><td>分房对象</td><td>分房理由</td><td>操作</td></tr>"
sqlstr="select  * from 洞房 where 操作=True order by 时间 desc"
 Set Rst=connt.Execute(sqlstr)
end if
set rst=server.CreateObject ("adodb.recordset")
rst.open sqlstr,connt,1,2
if not (rst.EOF or rst.BOF) then
rst.Pagesize=10
if acpage>rst.PageCount then acpage=rst.PageCount
rst.AbsolutePage=acpage
for i=1 to rst.PageSize
if rst.EOF then exit for
msg=msg&"<tr><td>"&rst("申请人")&"</td><td>"&rst("对象")&"</td><td>"&rst("说明")&"</td><td>"
if rst("对象")=username then
msg=msg&"<a href='modmar.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='"&sex&"说的好象很对，我还是同意了吧';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title='"&sex&"说的好象很对，我还是同意了吧'>答应</a>"
elseif rst("申请人")=username then
msg=msg&"<a href='delmar.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='想想还是算了吧';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title='想想还是算了吧'>删除</a>"
else
msg=msg&"操作"
end if
msg=msg&"</td></tr>"
rst.movenext
next
end if
msg=msg+"</table><form action='marriage.asp?opt="&opt&"' method=post id=form1 name=form1><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=ffffff><tr align=center>"
if acpage>1 then
msg=msg&"<td><a href='marriage.asp?opt="&opt&"&acpage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='marriage.asp?opt="&opt&"&acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
else
msg=msg&"<td>第一页</td><td>前一页</td>"
end if
if acpage<rst.PageCount then
msg=msg&"<td><a href='marriage.asp?opt="&opt&"&acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='marriage.asp?opt="&opt&"&acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
else
msg=msg&"<td>后一页</td><td>最后一页</td>"
end if
msg=msg&"<td>第<input type=text name=acpage size=4 value='"&acpage&"'>页共"&rst.PageCount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"
rst.Close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
rst.Open "select 洞房花烛 from 用户 where 姓名='"&username&"'",conn
mate=rst("洞房花烛")
if opt="洞房" and mate="未" then
msg=msg&"<form action='addmar.asp' method=post><table align=center><tr><td>性爱对象：<input type=text name=quest size=10 maxlength=7>性爱蜜语<input type=text name=content size=20 maxlength=50><input type=submit name=opt value='洞房'></td></tr></table></form>"
elseif opt="分房" and mate<>"未" then
msg=msg&"<form action='addmar.asp' method=post><table align=center><tr><td>分房：<input type=text name=quest size=10 maxlength=7 readonly value="&mate&">我的原因<input type=text name=content size=20 maxlength=50><input name=opt type=submit value='分房'></td></tr></table></form>"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
connt.Close
set connt=nothing
msg=msg&"<p align=center><a href='#' onclick='javascript:window.close();'>关闭</a></p></body>"
Response.Write msg
%>

