<!-- #include file="setup.asp" -->
<%
if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")

id=int(Request("id"))
incept=HTMLEncode(Request("incept"))
username=HTMLEncode(Request("username"))
Honor=HTMLEncode(Request("Honor"))

sql="select * from [user] where username='"&Request.Cookies("username")&"'"
Set Rs=Conn.Execute(sql)
faction=rs("faction")
experience=rs("experience")
money=rs("money")
rs.close
top

if Request.form("menu")="factionadd" then
if faction<>"" then error2("���Ѿ����� "&faction&" �ˣ������ټ����������ɣ�")
factionname=Conn.Execute("Select factionname From faction where id="&id&"")(0)
if conn.execute("Select count(id) from [user] where faction='"&factionname&"'")(0)>99 then error2("�ð����Ѿ�����100����Ա���޷��ټ����»�Ա")
conn.execute("delete from [message] where id="&int(Request("messageid"))&" and incept='"&Request.Cookies("username")&"'")
conn.execute("update [user] set faction='"&factionname&"' where username='"&Request.Cookies("username")&"'")
error2("������ɳɹ�")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="factionout" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>Ч�������<li>�����·���ˢ�º�����")
if faction=empty then error("<li>��Ŀǰû�м����κΰ��ɣ�")
If not conn.Execute("Select id From [faction] where buildman='"&Request.Cookies("username")&"'").eof Then error("<li>Ҫ�˳����Ƚ�ɢ����")
conn.execute("update [user] set faction='',honor='' where username='"&Request.Cookies("username")&"'")
message=message&"<li>�˳����ɳɹ�<li><a href=faction.asp>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="invite" then
if username="" then error("<li>����д�����˵�����")
if Request.Cookies("username")=username then error("<li>�����Լ������Լ�")
if Conn.Execute("Select buildman From [faction] where id="&id&"")(0)<>Request.Cookies("username") then error("<li>ֻ�а�������Ȩ��ִ�иò���")
if conn.execute("Select count(id) from [user] where faction='"&faction&"'")(0)>99 then error2("�����Ѿ�����100����Ա���޷��ټ����»�Ա")
if Conn.Execute("Select faction From [user] where username='"&username&"'")(0)<>"" then error("<li>�Է��Ѿ���������������")
messageid=conn.execute("select Max(ID)+1 From Message")(0)
conn.Execute("insert into [message] (author,incept,content) values ('"&Request.Cookies("username")&"','"&username&"','<form name=factionAdd"&messageid&" method=post action=faction.asp?id="&id&"&messageid="&messageid&"><input type=hidden name=menu value=factionadd></form><font color=0000FF>��ϵͳ��Ϣ����"&Request.Cookies("username")&" ���������� "&faction&" ����<br><br><center><a href=javascript:factionAdd"&messageid&".submit()>ͬ��</a>������<a href=message.asp?menu=del&id="&messageid&">�ܾ�</a></font>')")
conn.execute("update [user] set newmessage=newmessage+1 where username='"&username&"'")
message=message&"<li>�����Ѿ��ɹ�����<li><a href=faction.asp>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="factiondel" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>Ч�������<li>�����·���ˢ�º�����")
if Conn.Execute("Select buildman From [faction] where id="&id&"")(0)<>Request.Cookies("username")then error("<li>ֻ�а�������Ȩ��ִ�иò���")
conn.execute("update [user] set faction='',honor='' where faction='"&faction&"'")
conn.execute("delete from [faction] where id="&id&"")
message=message&"<li>��ɢ���ɳɹ�<li><a href=faction.asp>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="factionUserOut" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>Ч�������<li>�����·���ˢ�º�����")
if Request.Cookies("username")=username then error("<li>�����Լ������Լ�")
if Conn.Execute("Select buildman From [faction] where id="&id&"")(0)<>Request.Cookies("username")then error("<li>ֻ�а�������Ȩ��ִ�иò���")
conn.execute("update [user] set faction='',honor='' where username='"&username&"' and faction='"&faction&"'")
message=message&"<li>�Ѿ��� "&username&" �Ӱ����п�����<li><a href=faction.asp>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="factionUserHonor" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>Ч�������<li>�����·���ˢ�º�����")
if Len(Honor)>7 then error("<li>ͷ�γ��Ȳ��ܳ���7���ַ�")
if Conn.Execute("Select buildman From [faction] where id="&id&"")(0)<>Request.Cookies("username")then error("<li>ֻ�а�������Ȩ��ִ�иò���")
conn.execute("update [user] set honor='"&Honor&"' where username='"&username&"' and faction='"&faction&"'")
message=message&"<li> "&username&" �Ѿ���� "&Honor&" ��ͷ��<li><a href=faction.asp>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="addok" then
factionname=HTMLEncode(Request.Form("factionname"))
allname=HTMLEncode(Request.Form("allname"))
tenet=HTMLEncode(Request.Form("tenet"))
if faction<>empty then message=message&"<li>���Ѿ��������������ɣ�"
if experience< 10000 then message=message&"<li>���ľ���ֵС�� 10000 ��"
if money< 10000 then message=message&"<li>���Ľ������ 10000 ��"
if factionname="" then message=message&"<li>���ɼ��û����д"
if Len(factionname)>7 then message=message&"<li>���ɼ�����7���ַ�"
if allname="" then message=message&"<li>����ȫ��û����д"
If not conn.Execute("Select id From [faction] where factionname='"&factionname&"' or buildman='"&Request.Cookies("username")&"'").eof Then  message=message&"<li>�������Ѵ���ͬ������<li>���Ѿ�����������"
if message<>"" then error(""&message&"")
conn.Execute("insert into faction (factionname,allname,tenet,buildman) values ('"&factionname&"','"&allname&"','"&tenet&"','"&Request.Cookies("username")&"')")
conn.execute("update [user] set faction='"&factionname&"',[money]=[money]-10000 where username='"&Request.Cookies("username")&"'")
message=message&"<li>�������ɳɹ�<li><a href=faction.asp>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="xiuok" then
allname=HTMLEncode(Request.Form("allname"))
tenet=HTMLEncode(Request.Form("tenet"))
if allname="" then message=message&"<li>����ȫ��û����д"
if message<>"" then error(""&message&"")
if Conn.Execute("Select buildman From [faction] where id="&id&"")(0)<>Request.Cookies("username")then error("<li>ֻ�а�������Ȩ��ִ�иò���")
conn.execute("update [faction] set allname='"&allname&"',tenet='"&tenet&"' where id="&id&"")
message=message&"<li>�޸İ��ɳɹ�<li><a href=faction.asp>������������</a><li><a href=Default.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="look" then
sql="select * from faction where id="&id&""
Set Rs=Conn.Execute(sql)
%>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� 
<a href="faction.asp">��������</a></td>
</tr>
</table><br>
<table width="82%" border="0" align="center" cellspacing="1" cellpadding="2"  class=a2 height="150">
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>���ɼ��:</b></font></div>
</td>
<td width="82%"><%=rs("factionname")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>����ȫ��:</b></font></div>
</td>
<td width="82%"><%=rs("allname")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>���ɹ���:</b></font></div>
</td>
<td width="82%"><%=rs("tenet")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>����ʱ��:</b></font></div>
</td>
<td width="82%"><%=rs("addtime")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>��������:</b></font></div>
</td>
<td width="82%"><%=rs("buildman")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>���л�Ա:</b></font></div>
</td>
<td width="82%">
<%
sql="select username from [user] where faction='"&rs("factionname")&"'"
Set Rs=Conn.Execute(sql)
Do While Not RS.EOF
i=i+1
list=list&""&rs("username")&" "
RS.MoveNext
loop
%><%=i%>��</td>
</tr>

<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>��Ա����:</b></font></div>
</td>
<td width="82%">
<%=list%>
</td>
</tr>

</table>
<br><center><INPUT onclick=history.back(-1) type=button value=" << �� �� ">
<%

htmlend
end if



%>
<center>

<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� 
<a href="faction.asp">��������</a></td>
</tr>
</table><br>
<%
if Request("menu")="add" then

%>
<form method=post name=form action=faction.asp?menu=addok>

<table cellspacing=1 cellpadding=2 width=442 border=0 align="center" class=a2>
<tr class=a1>
<td width=526 colspan="2" align="center" height="25">
��������</td>
</tr>
<tr class=a3>
<td width=187 align="right">
<b><font color="0033CC">���ɼ�ƣ�</font></b></td>
<td width=339>
<input maxlength=7 name=factionname size="10"> ���7���ַ�</td>
</tr>
<tr class=a3>
<td width=187>
<div align="right"><b><font color="0033CC">����ȫ�ƣ� </font></b></div>
</td>
<td width=339>
<input size=30 name=allname>
</td>
</tr>
<tr class=a3>
<td width=187 height=15>
<div align="right"><b><font color="0033CC">���ɹ��棺 </font></b></div>
</td>
<td width=339 height=15>
<input size=40 name=tenet>
</td>
</tr>
<tr class=a3>
<td width=526 colspan=2 height=8>
<div align=center>
<input type=submit value=" �� �� " name=Submit23>
<input type=reset value=" �� �� " name=Submit24>
</div>
</td>
</tr>
<tr class=a3>
<td width=526 colspan=2 height=7>
<ol>
�������ɵ�ע�����
<li>���ľ���ֵ���� 10000 ����
<li>��Ҫ�۳������� 10000 �����Ϊ���ɻ��� </li>
<li>�������ֻ������ 100 ����Ա</td>
</tr>
</table>


</form>



<%
elseif Request("menu")="xiu" then
sql="select * from faction where id="&id&""
Set Rs=Conn.Execute(SQL)
%>

<form method=post action=faction.asp?menu=xiuok&id=<%=rs("id")%>>
<table cellpadding="2" cellspacing="1" width="70%" border="0" class=a2>

<tr>
<td colspan="2" height="25" align="center" class=a1>���������趨</td>
</tr>


<tr class=a3>
<td>�������ɼ�ƣ� </td>
<td>
<%=rs("factionname")%></td>
</tr>
<tr class=a3>
<td>��������ȫ�ƣ� </td>
<td><input size="30" name="allname" value="<%=rs("allname")%>"> </td>
</tr>
<tr class=a3>
<td>�������ɹ��棺 </td>
<td><input size="60" name="tenet" value="<%=rs("tenet")%>"> </td>
</tr>
<tr class=a3>
<td colSpan="2">
<div align="center">
<input type="submit" value=" �� �� " name="Submit">
<input type="reset" value=" �� �� " name="Submit2">
</div>
</td>
</tr>
</table>
</form>
<%
else
%>


<p>&nbsp;<a href="?menu=add"><img src="images/plus/niu05.gif" width="26" height="26" border="0"><img src="images/plus/cj.gif" width="95" height="26" border="0"></a>��������<a href="?menu=factionout&sessionid=<%=session.sessionid%>"><img src="images/plus/niu05.gif" width="26" height="26" border="0"><img src="images/plus/tc.gif" width="95" height="26" border="0"></a><br>
</p>

<SCRIPT>valigntop()</SCRIPT>



<table width="97%" border="0" align="center" cellspacing="1" cellpadding="5"  class=a2>



<%


if faction<>"" then
sql="select * from faction where factionname='"&faction&"'"
Set Rs=Conn.Execute(sql)
if rs.eof then conn.execute("update [user] set faction='',honor='' where username='"&Request.Cookies("username")&"'")
sql="select username from [user] where faction='"&rs("factionname")&"'"
Set Rs1=Conn.Execute(sql)
Do While Not RS1.EOF
i=i+1
list=list&"<input type=radio value='"&rs1("username")&"' name=username id="&i&"><label for="&i&">"&rs1("username")&"</label> "
RS1.MoveNext
loop
Set Rs1 = Nothing

%>
<SCRIPT>
function factionUser(username,act){
var usernames
if (username.length > 1){
for(iIndex=0;iIndex<username.length;iIndex++){if(username[iIndex].checked==true){usernames = username[iIndex].value;}}
}
else{if(username.checked==true){usernames=username.value;}}
if(usernames==undefined){alert('��ѡ�����');return;}
if(act=="add"){document.location='friend.asp?menu=add&username='+usernames+'';}
if(act=="post"){open('friend.asp?menu=post&incept='+usernames+'','','width=320,height=170')}
if(act=="look"){open('Profile.asp?username='+usernames+'')}
if(act=="out"){document.location='faction.asp?menu=factionUserOut&id=<%=rs("id")%>&sessionid=<%=session.sessionid%>&username='+usernames+'';}
if(act=="Honor"){var id=prompt("������û�Ա��ͷ�Σ�","");if(id){document.location='faction.asp?menu=factionUserHonor&id=<%=rs("id")%>&sessionid=<%=session.sessionid%>&username='+usernames+'&Honor='+id+'';}}
}

function invite(){
var id=prompt("������������뱾���ɵĻ�Ա���ƣ�","");if(id){document.location='faction.asp?menu=invite&id=<%=rs("id")%>&sessionid=<%=session.sessionid%>&username='+id+'';}
}

</SCRIPT>

<tr bgcolor="FFFFFF">
<td width="10%">
<div align="center"><font color="000066"><b>���ɼ��</b></font><font color="#000066"><b>:</b></font></div>
</td>
<td width="28%"><%=rs("factionname")%></td>
<td width="10%" align="center"><font color="000066"><b>����ȫ��</b></font><font color="#000066"><b>:</b></font></td>
<td width="27%"><%=rs("allname")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="10%">
<div align="center"><font color="000066"><b>������</b></font><font color="#000066"><b>:</b></font></div>
</td>
<td width="28%"><%=rs("buildman")%></td>
<td width="10%" align="center"><font color="000066"><b>����ʱ��</b></font><font color="#000066"><b>:</b></font></td>
<td width="27%"><%=rs("addtime")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="10%">
<div align="center"><font color="000066"><b>���ɹ���</b></font><font color="#000066"><b>:</b></font></div>
</td>
<td width="82%" colspan="3"><%=rs("tenet")%></td>
</tr>
<form method=post name=faction>
<tr bgcolor="FFFFFF">
<td width="10%">
<div align="center"><font color="000066"><b>��Ա����</b></font><font color="#000066"><b>:</b></font><br>
	<font size="1">�� <font color=red><%=i%></font> ��</font></div>
</td>
<td width="65%" colspan="3">

<%=list%>
</td>
</tr>

<tr>
<td width="38%" colspan="2" align="center" class=a1>
��Ա����</td>
<td width="37%" colspan="2" align="center" class=a1>
��������</td>
</tr>
<tr bgcolor="FFFFFF">
<td width="38%" colspan="2" align="center">
<a onclick="factionUser(document.faction.username,'add')" href=#>��Ӻ���</a>����<a onclick="factionUser(document.faction.username,'post')" href=#>������Ϣ</a>����<a onclick="factionUser(document.faction.username,'look')" href=#>�鿴����</a></td>
<td width="37%" colspan="2" align="center">


<a onclick="invite()" href=#>��ӻ�Ա</a>����<a onclick="factionUser(document.faction.username,'out')" href=#>������Ա</a>����<a onclick="factionUser(document.faction.username,'Honor')" href=#>����ͷ��</a>����<a href="?menu=xiu&id=<%=rs("id")%>">�޸�����</a>����<a onclick=checkclick('��ȷ��Ҫ��ɢ�ð��ɣ�') href="?menu=factiondel&sessionid=<%=session.sessionid%>&id=<%=rs("id")%>">��ɢ�˰�</a></td>
</tr>
</form>
<%
RS.Close


else
%>
<tr bgcolor="FFFFFF"><td align="center">û�д������߼����κΰ���</td></tr>
<%
end if
%>
</table>




<SCRIPT>valignbottom()</SCRIPT>

<%
end if


htmlend
%>