<!--#include file="../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
mycorp=session("yx8_mhjh_usercorp")
mygrade=session("yx8_mhjh_usergrade")
if mygrade="" then Response.Redirect "../error.asp?id=016"
nowtime=now()
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 最后登录时间,姓名,状态 from 用户 where 状态 in('逮捕','入狱')"
rst.Open sqlstr,conn
do while not (rst.EOF or rst.BOF )
if rst("状态")="逮捕" or rst("状态")="入狱" then
if mycorp="官府" and mygrade>=arrestright then
msg=msg&"<tr><td>"&rst("姓名")&"</td><td>"&rst("最后登录时间")&"</td><td><a href='assoil.asp?un="&rst("姓名")&"' onmouseover="&chr(34)&"window.status='释放"&rst("姓名")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">释放</a>|<a href='zhanshou2.asp?name="&rst("姓名")&"'>斩首</a></td></tr>"
else
msg=msg&"<tr><td>"&rst("姓名")&"</td><td>"&rst("最后登录时间")&"</td><td><a href='mainprise.asp?un="&rst("姓名")&"'  onmouseover="&chr(34)&"window.status='保释"&rst("姓名")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">保释</a>|<a href='jieyu.asp?name="&rst("姓名")&"'>劫狱</a></td></tr>"
end if
else
if mycorp="官府" and mygrade>=gaolright then
msg=msg&"<tr><td>"&rst("姓名")&"</td><td>"&rst("最后登录时间")&"</td><td><a href='assoil.asp?un="&rst("姓名")&"'  onmouseover="&chr(34)&"window.status='释放"&rst("姓名")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">释放</a>|<a href='zhanshou2.asp?name="&rst("姓名")&"'>斩首</a></td></tr>"
else
msg=msg&"<tr><td>"&rst("姓名")&"</td><td>"&rst("最后登录时间")&"</td><td>不可保释</td></tr>"
end if
end if
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><title>官府牢房</title><LINK href="../style.css" rel=stylesheet></head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false>
<p align=center><b><font size="4" >官府牢房</font></b></p>
<div align=center>
<table width='95%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF>
<tr align=center bgcolor=#ffffff>
<td>姓名</td>
<td>出狱时间</td>
<td>操作</font></td>
</tr>
<%=msg%>
</table>
</div>

</body>



