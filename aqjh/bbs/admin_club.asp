<!-- #include file="setup.asp" -->
<%
if adminpassword<>session("pass") then response.redirect "admin.asp?menu=login"
log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")
id=int(Request("id"))

TimeLimit=HTMLEncode(Request("TimeLimit"))
username=HTMLEncode(Request("username"))
membercode=HTMLEncode(Request("membercode"))


response.write "<center>"

select case Request("menu")


case "message"
message

case "broadcast"
broadcast

case "sendmail"
sendmail

case "sendmailok"
sendmailok



case "messageok"
if TimeLimit="" then error2("��û��ѡ�����ڣ�")
conn.execute("delete from [message] where time<"&SqlNowString&"-"&TimeLimit&"")
error2("�Ѿ���"&TimeLimit&"����ǰ�Ķ�ѶϢɾ���ˣ�")


case "delmessageuser"
if username="" then error2("��û�������û�����")
conn.execute("delete from [message] where author='"&username&"' or incept='"&username&"'")
error2("�Ѿ���"&username&"�Ķ�ѶϢȫ��ɾ���ˣ�")

case "delmessagekey"
key=HTMLEncode(Request("key"))
if key="" then error2("��û������ؼ��ʣ�")
conn.execute("delete from [message] where content like '%"&key&"%'")
error2("�Ѿ��������а��� "&key&" �Ķ�ѶϢɾ���ˣ�")


case "faction"
faction

case "editfaction"
editfaction

case "editfactionok"
editfactionok

case "delfaction"
delfaction


case "link"
link

case "linkok"
linkok

case "editlink"
editlink


case "editlinkok"
rs.Open "select * from [link] where id="&id&"",conn,1,3
rs("name")=Request("name")
rs("url")=Request("url")
rs("logo")=Request("logo")
rs("intro")=Request("intro")
rs.update
rs.close
%>�� �� �� �� ��<p><a href=javascript:history.back()>< �� �� ></a><%


case "dellink"
conn.execute("delete from [link] where id="&id&"")
link



end select


sub sendmail
%>

<form method="post" action="?menu=sendmailok">
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>Ⱥ���ʼ�</td>
  </tr>
  <tr height=25>
    <td class=a3 align=left>&nbsp;&nbsp; ���⣺<input size="40" name="title"></td>
    <td class=a3 align=middle>���ն���
<select name=membercode>
<option value="">���л�Ա</option>
<option value="1">��ͨ��Ա</option>
<option value="2">�����Ա</option>
<option value="4">�� �� Ա</option>
<option value="5">��������</option>
</select>&nbsp;&nbsp;&nbsp; </td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
 <textarea name="content" rows="5" cols="70"></textarea>
</td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
 <input type="submit" value=" �� �� " name="submit">
<input type="reset" value=" �� �� " name="Submit2"><br></td></tr></table></form>

<%
end sub


sub sendmailok

if Request("title")="" then error2("����д�ʼ�����")
if Request("content")="" then error2("����д�ʼ�����")


if membercode<>"" then
sql="select usermail from [user] where membercode="&membercode&""
else
sql="select usermail from [user]"
end if

Set Rs=Conn.Execute(sql)
do while not rs.eof

mailaddress=""&rs("usermail")&""
mailtopic=Request("title")
body=""&Request("content")&""&vbCrlf&""&vbCrlf&"���ʼ�ͨ�� BBSxp Ⱥ��ϵͳ���͡�����������Yuzi������(http://www.yuzi.net)"
%><!-- #include file="inc/mail.asp" --><%

rs.movenext
loop
rs.close

response.write "�ʼ����ͳɹ���"

end sub


sub message
%>




���ݿ⹲�� <%=conn.execute("Select count(id)from message")(0)%> ����ѶϢ
<br><br>

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25 class=a1>
		<td align="center">����ɾ������Ϣ</td>
	</tr>
	<tr class=a3>
		<td align="center"><form method="post" action="?menu=delmessageuser">����ɾ�� <input size="13" name="username" onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')" onfocus="javascript:focusEdit(this)" onblur="javascript:blurEdit(this)" value="�û���" helptext="�û���"> �Ķ�ѶϢ
<input type="submit" value="ȷ��" id='btnadd' disabled>
		</td></form>
	</tr>
	
	<tr class=a3>
		<td align="center"><form method="post" action="?menu=delmessagekey">����ɾ�����ݺ��� <input size="20" name="key" onkeyup="ValidateTextboxAdd(this, 'nrkey')" onpropertychange="ValidateTextboxAdd(this, 'nrkey')" onfocus="javascript:focusEdit(this)" onblur="javascript:blurEdit(this)" value="�ؼ���" helptext="�ؼ���"> �Ķ�ѶϢ
<input type="submit" value="ȷ��" id='nrkey' disabled>
		</td></form>
	</tr>
	
		<tr class=a3>
		<td align="center"><form method="post" action="?menu=messageok">ɾ�� <INPUT size=2 name=TimeLimit value="30">
			����ǰ�Ķ�ѶϢ
<input type="submit" value="ȷ��">

		</td></form>
	</tr>
	
</table>
</form>
<form method="post" action="?menu=broadcast">
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle width="50%">ϵͳ�㲥</td>
    <td class=a1 align=middle width="50%">���ն���
<select name=membercode>
<option value="">���߻�Ա</option>
<option value="1">��ͨ��Ա</option>
<option value="2">�����Ա</option>
<option value="4">�� �� Ա</option>
<option value="5">��������</option>
</select>
</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
	<textarea name="content" rows="5" cols="70"></textarea>
</td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
	<input type="submit" value=" �� �� " name="submit5">
<input type="reset" value=" �� �� " name="Submit6"></td></tr></table></form>
<%
end sub

sub broadcast
content=HTMLEncode(Request.Form("content"))

if content="" then error2("����д�㲥����!")

if membercode<>"" then
sql="select username from [user] where membercode="&membercode&""
else
sql="select username from online where username<>''"
end if

Set Rs=Conn.Execute(sql)
do while not rs.eof
Count=Count+1
conn.Execute("insert into message (author,incept,content) values ('"&Request.Cookies("username")&"','"&rs("username")&"','<font color=0000FF>��ϵͳ�㲥����"&content&"</font>')")
conn.execute("update [user] set newmessage=newmessage+1 where username='"&rs("username")&"'")
rs.movenext
loop
rs.close

%>
�����ɹ�
<br><br>
�����͸� <%=Count%> λ�����û�<br><br>
<a href=javascript:history.back()>�� ��</a>
<%
end sub




sub link
%>
<FORM action=?menu=linkok method=post>
<table cellspacing="1" cellpadding="5" width="97%" border="0" class="a2" align=center><tr>
	<td height="7" class="a1" colspan="2">����<b> </b>�������ӹ���</td></tr><tr>
	<td height="6" class="a3">��̳���ƣ�<INPUT size=40 name=name></td>
	<td height="6" class="a3">��ַURL��<INPUT size=40 name=url value="http://"></td></tr><tr>
	<td height="6" class="a3">��̳��飺<INPUT size=40 name=intro></td>
	<td height="6" class="a3">ͼ��URL��<INPUT size=40 name=logo value="http://"></td></tr><tr>
	<td height="6" class="a4" colspan="2" align="center"><INPUT type=submit value=" �� �� " name=Submit>
<input type="reset" value=" �� �� " name="Submit2">

</td></tr></table>
</FORM>


<table cellspacing="1" cellpadding="5" width="97%" border="0" class="a2" align="center"><tr><td height="25" colspan="2" class="a1">����<b> </b>��������</td></tr><tr>
<td align="center" bgcolor=FFFFFF width="5%"><img src="images/shareforum.gif"></td>
<td class="a4"><%
rs.Open "link",Conn
do while not rs.eof
if rs("logo")="" or rs("logo")="http://" then
link1=link1+"<a onmouseover="&Chr(34)&"showmenu(event,'<div class=menuitems><a href=?menu=editlink&id="&rs("id")&">�༭</a></div><div class=menuitems><a href=?menu=dellink&id="&rs("id")&">ɾ��</a></div>')"&Chr(34)&" title='"&rs("intro")&"' href="&rs("url")&" target=_blank>"&rs("name")&"</a>����"
else
link2=link2+"<a onmouseover="&Chr(34)&"showmenu(event,'<div class=menuitems><a href=?menu=editlink&id="&rs("id")&">�༭</a></div><div class=menuitems><a href=?menu=dellink&id="&rs("id")&">ɾ��</a></div>')"&Chr(34)&" title='"&rs("name")&""&chr(10)&""&rs("intro")&"' href="&rs("url")&" target=_blank><img src="&rs("logo")&" border=0 width=88 height=31></a>����"
end if
rs.movenext
loop
rs.close
%>
<%=link1%>
<br><br>
<%=link2%>
</td></tr></table>


<%


end sub



sub editlink

sql="Select * From link where id="&id&""
Set Rs=Conn.Execute(sql)
%>
<FORM action=?menu=editlinkok method=post>
<input type=hidden name=id value=<%=id%>>
<table cellspacing="1" cellpadding="5" width="97%" border="0" class="a2" align=center><tr>
	<td height="7" class="a1" colspan="2">����<b> </b>�������ӹ���</td></tr><tr>
	<td height="6" class="a3">��̳���ƣ�<INPUT size=40 name=name value="<%=rs("name")%>"></td>
	<td height="6" class="a3">��ַURL��<INPUT size=40 name=url value="<%=rs("url")%>"></td></tr><tr>
	<td height="6" class="a3">��̳��飺<INPUT size=40 name=intro value="<%=rs("intro")%>"></td>
	<td height="6" class="a3">ͼ��URL��<INPUT size=40 name=logo value="<%=rs("logo")%>"></td></tr><tr>
	<td height="6" class="a4" colspan="2" align="center"><INPUT type=submit value=" �� �� " name=Submit>
<input type="reset" value=" �� �� " name="Submit2">
</td></tr></table>
</FORM><p><a href=javascript:history.back()>< �� �� ></a>
<%
end sub



sub linkok

if Request("url")="http://" or Request("url")="" then error2("��̳URLû����д")

rs.Open "link",conn,1,3
rs.addnew
rs("name")=Request("name")
rs("url")=Request("url")
rs("logo")=Request("logo")
rs("intro")=Request("intro")
rs.update
rs.close

link
end sub


sub faction
%>
<table border="0" cellpadding="5" cellspacing="1" class=a2 width=97%>
<tr>
<td width="15%" align="center" height="25" class=a1>�ɱ�</td>
<td width="40%" align="center" height="25" class=a1>����</td>
<td width="15%" align="center" height="25" class=a1>��ʼ��</td>
<td width="20%" align="center" height="25" class=a1>����</td>
</tr>
<%
sql="select * from faction order by addtime Desc"
rs.Open sql,Conn,1
pagesetup=20 '�趨ÿҳ����ʾ����
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '��ҳ��
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '��ת��ָ��ҳ��
i=0
Do While Not RS.EOF and i<pagesetup
i=i+1%>
<tr>
<td width="10%" align="center" height="25" bgcolor="FFFFFF"> <a target="_blank" href=faction.asp?menu=look&id=<%=rs("id")%>><%=rs("factionname")%></a>
<td width="50%" align="center" height="25" bgcolor="FFFFFF"><%=rs("tenet")%>
<td width="10%" align="center" height="25" bgcolor="FFFFFF"><a target="_blank" href="Profile.asp?username=<%=rs("buildman")%>"><%=rs("buildman")%></a>
<td width="20%" align="center" height="25" bgcolor="FFFFFF"><a href="?menu=editfaction&id=<%=rs("id")%>">�޸�����</a> <a onclick=checkclick('��ȷ��Ҫɾ���ð��ɣ�') href="?menu=delfaction&id=<%=rs("id")%>">ɾ������</a></td>
</tr>
<%
RS.MoveNext
loop
RS.Close
%>
</table>
[<b>
<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"menu=faction")
</script>
</b>]

<%
end sub

sub editfaction
sql="select * from faction where id="&id&""
Set Rs=Conn.Execute(SQL)
%>
<form method=post action=?menu=editfactionok&id=<%=rs("id")%>>
<table cellpadding="2" cellspacing="1" width="70%" border="0" class=a2>
<tr>
<td colspan="2" height="25" align="center" class=a1>�����趨</td>
</tr>
<tr class=a3>
<td>�������ɼ�ƣ� </td>
<td>
<input size="20" maxlength=7 name="factionname" value="<%=rs("factionname")%>"> 
���7���ַ�</td>
</tr>
<tr class=a3>
<td>�����������ƣ� </td>
<td><input size="30" name="allname" value="<%=rs("allname")%>"> </td>
</tr>
<tr class=a3>
<td>�����������ƣ� </td>
<td><input size="30" name="buildman" value="<%=rs("buildman")%>"> </td>
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
</form><p><a href=javascript:history.back()>< �� �� ></a>
<%
end sub


sub editfactionok
factionname=HTMLEncode(Request("factionname"))
allname=HTMLEncode(Request("allname"))
tenet=HTMLEncode(Request("tenet"))
buildman=HTMLEncode(Request("buildman"))

if factionname="" then error2("���ɼ��û����д")
if allname="" then error2("��������û����д")
if buildman="" then error2("��������û����д")

sql="select * from faction where id="&id&""
rs.Open sql,Conn,1,3
oldfactionname=rs("factionname")
rs("factionname")=factionname
rs("allname")=allname
rs("buildman")=buildman
rs("tenet")=tenet
rs.update
rs.close
conn.execute("update [user] set faction='"&factionname&"' where faction='"&oldfactionname&"'")
faction
end sub

sub delfaction
sql="select * from faction where id="&id&""
Set Rs1=Conn.Execute(sql)
conn.execute("update [user] set faction='' where faction='"&rs1("factionname")&"'")
set rs1=nothing
conn.execute("delete from [faction] where id="&id&"")
faction
end sub


htmlend

%>