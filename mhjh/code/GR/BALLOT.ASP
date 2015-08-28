<%
Response.Expires=-1
nowtime=now()
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
acpage=Request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=clng(acpage)
if acpage<1 then acpage=1
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select id,title,explain,deadline,condition from ballotsystem where active=true order by deadline desc",conn,1,3
if rst.EOF or rst.BOF then
msg="<tr><td>目前没有活动的投票箱</td></tr>"
else
rst.PageSize=1
if acpage>rst.PageCount then acpage=rst.PageCount
rst.AbsolutePage=acpage
pid=rst("id")
title=rst("title")
explain=replace(rst("explain"),chr(13)&chr(10),"<br>")
deadline=rst("deadline")
condition=Trim(rst("condition"))
if condition="true" then
condition="所有人"
elseif condition="false" then
condition="禁止任何人投票"
else
condition=replace(condition,">=","不小于")
condition=replace(condition,"<=","不大于")
condition=replace(condition,">","高于")
condition=replace(condition,"<","低于")
condition=replace(condition,"=","为")
condition=replace(condition," and ","并且")
condition=replace(condition," or ","或")
end if
msgadd="</table></form><form action='ballot.asp' method=post id=form2 name=form2><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=005b00><tr align=center>"
if acpage>1 then
msgadd=msgadd&"<td><a href='ballot.asp?acpage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='ballot.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
else
msgadd=msgadd&"<td>第一页</td><td>前一页</td>"
end if
if acpage<rst.PageCount then
msgadd=msgadd&"<td><a href='ballot.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='ballot.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
else
msgadd=msgadd&"<td>后一页</td><td>最后一页</td>"
end if
msgadd=msgadd&"<td>第<input type=text name=acpage size=4 value='"&acpage&"'>页共"&rst.PageCount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr>"
rst.Close
rst.Open "select id,parti from ballot where pid="&pid,conn
if rst.EOF or rst.BOF then
msg="<tr><td>无效的投票箱</td></tr>"&msgadd
else
msg="<form name=form1 action=ballotnow.asp method=post><tr><td align=center bgcolor=005b00>"&title&"</td></tr><tr><td>"&explain&"</td></tr><tr><td>投票截止时间："&deadline&"</td></tr><tr><td>参加投票条件："&condition&"</td><input type=hidden name=pid value="&pid&"></tr>"
do while not rst.EOF
id=rst("id")
parti=rst("parti")
msg=msg&"<tr><td><input type=radio name=id value="&id&">"&parti&"</td></tr>"
rst.MoveNext
loop
msg=msg&"<tr align=center><td> <input type=submit value=' 投 票 '  id=submit1 name=submit1> <input type=button onclick=javascript:location.href='showballot.asp?pid="&pid&"'; value='查看投票结果' > </td></tr></td></tr>"&msgadd
end if
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel='stylesheet' href='../style.css'>
<script language=javascript>function check(){for(var i=0;i<document.form1.id.length;i++){if(document.form1.id[i].checked){return true;}}alert('请投票！');return false;}</script>
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=0  background='../chatroom/bg1.gif'>
<p align=center><font face="隶书" size="3"><b>社区投票</b></font></p>
<hr><form action='ballotnow.asp' method=post onsubmit='javascript:return(check());' name=form1>
<table border=0 width=80% align=center>
<%=msg%>
</table></form></body>
