<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
mycorp=session("Ba_jxqy_usercorp")
mygrade=session("Ba_jxqy_usergrade")
if mygrade="" then Response.Redirect "../error.asp?id=016"
nowtime=now()
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 最后登录时间,姓名,状态 from 用户 where 状态 in('逮捕','入狱')"
rst.Open sqlstr,conn
do while not (rst.EOF or rst.BOF )
	if rst("状态")="逮捕" then
		if mycorp="官府" and mygrade>=Application("Ba_jxqy_arrestright") then
			msg=msg&"<tr><td>"&rst("姓名")&"</td><td>"&rst("最后登录时间")&"</td><td><a href='assoil.asp?un="&rst("姓名")&"' onmouseover="&chr(34)&"window.status='释放"&rst("姓名")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">释放</a></td></tr>"
		else
			msg=msg&"<tr><td>"&rst("姓名")&"</td><td>"&rst("最后登录时间")&"</td><td><a href='mainprise.asp?un="&rst("姓名")&"'  onmouseover="&chr(34)&"window.status='保释"&rst("姓名")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">保释</a></td></tr>"	
		end if
	else
		if mycorp="官府" and mygrade>=Application("Ba_jxqy_gaolright") then
			msg=msg&"<tr><td>"&rst("姓名")&"</td><td>"&rst("最后登录时间")&"</td><td><a href='assoil.asp?un="&rst("姓名")&"'  onmouseover="&chr(34)&"window.status='释放"&rst("姓名")&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">释放</a></td></tr>"
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
<head><title><%=Application("Ba_jxqy_systemname")%></title><LINK href="../chatroom/css.css" rel=stylesheet></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color="#CC0000" size="4" face="幼圆"><b><%=Application("Ba_jxqy_systemname")%>监狱</b></font></p>
<div align=center>
  <table width='95%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF>
    <tr align=center bgcolor=#FFFF00> 
      <td>姓名</td>
      <td>出狱时间</td>
      <td>操作</td>
    </tr>
<%=msg%>
</table>
</div>
<p align=center><input type="button" value=" 关 闭 " onClick="javascript:top.window.close();" name="button"></p>
</body>