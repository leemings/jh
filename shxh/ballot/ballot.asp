<%
Response.Expires=-1
nowtime=now()
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
acpage=Request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=clng(acpage)
if acpage<1 then acpage=1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select id,title,explain,deadline,condition from ballotsystem where active=true order by deadline desc",conn,1,3
if rst.EOF or rst.BOF then
	msg="<tr><td>Ŀǰû�л��ͶƱ��</td></tr><tr><td align=center><input type=button value=' �� �� ' onclick=javascript:location.href='../myhome/myhome.asp'></td></tr>"
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
		condition="������"
	elseif condition="false" then
		condition="��ֹ�κ���ͶƱ"
	else		
		condition=replace(condition,">=","��С��")
		condition=replace(condition,"<=","������")
		condition=replace(condition,">","����")
		condition=replace(condition,"<","����")
		condition=replace(condition,"=","Ϊ")
		condition=replace(condition," and ","����")
		condition=replace(condition," or ","��")
	end if
	msgadd="</table></form><form action='ballot.asp' method=post id=form2 name=form2><table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=ffff00><tr align=center>"
	if acpage>1 then
		msgadd=msgadd&"<td><a href='ballot.asp?acpage=1' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='ballot.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='ǰһҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">ǰһҳ</a></td>"
	else
		msgadd=msgadd&"<td>��һҳ</td><td>ǰһҳ</td>"
	end if
	if acpage<rst.PageCount then
		msgadd=msgadd&"<td><a href='ballot.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='ballot.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='���һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">���һҳ</a></td>"
	else
		msgadd=msgadd&"<td>��һҳ</td><td>���һҳ</td>"
	end if
	msgadd=msgadd&"<td>��<input type=text name=acpage size=4 value='"&acpage&"'>ҳ��"&rst.PageCount&"ҳ<input type=submit value='GO' id=submit1 name=submit1></td></tr>"
	rst.Close
	rst.Open "select id,parti from ballot where pid="&pid,conn
	if rst.EOF or rst.BOF then
		msg="<tr><td>��Ч��ͶƱ��</td></tr><tr><td align=center><input type=button value=' �� �� ' onclick=javascript:location.href='../myhome/myhome.asp' ></td></tr>"&msgadd
	else
		msg="<form name=form1 action=ballotnow.asp method=post><tr><td align=center bgcolor=ffff00>"&title&"</td></tr><tr><td>"&explain&"</td></tr><tr><td>ͶƱ��ֹʱ�䣺"&deadline&"</td></tr><tr><td>�μ�ͶƱ������"&condition&"</td><input type=hidden name=pid value="&pid&"></tr>"
		do while not rst.EOF
			id=rst("id")
			parti=rst("parti")
			msg=msg&"<tr><td><input type=radio name=id value="&id&">"&parti&"</td></tr>"
			rst.MoveNext
		loop
		msg=msg&"<tr align=center><td> <input type=submit value=' Ͷ Ʊ '  id=submit1 name=submit1> <input type=button onclick=javascript:location.href='showballot.asp?pid="&pid&"'; value='�鿴ͶƱ���' > <input type=button value=' �� �� ' onclick=javascript:location.href='../myhome/myhome.asp' ></td></tr></td></tr>"&msgadd
	end if
end if
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel='stylesheet' href='../chatroom/css.css'>
<script language=javascript>function check(){for(var i=0;i<document.form1.id.length;i++){if(document.form1.id[i].checked){return true;}}alert('��ͶƱ��');return false;}</script>
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=0 background='<%=bgimage%>' bgcolor='<%=bgcolor%>'>
<p align=center><font color=0000ff size=6 face='����'>ͶƱϵͳ</font></p>
<hr><form action='ballotnow.asp' method=post onsubmit='javascript:return(check());' name=form1>
<table border=0 width=80% align=center>
<%=msg%>
</table></form></body>