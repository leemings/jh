<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
if adminpassword<>session("pass") then response.redirect "admin.asp?menu=login"
log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")

username=HTMLEncode(Request("username"))


response.write "<center>"

select case Request("menu")


case "agreement"
agreement

case "agreementok"
rs.Open "clubconfig",Conn,1,3
rs("agreement")=Request("agreement")
rs.update
rs.close
%>
���³ɹ�<br><br><a href=javascript:history.back()>�� ��</a>
<%
case "Activation"
Activation

case "user2"
user2
case "user"
user
case "showall"
showall

case "showallok"
showallok

case "userdeltopic"
userdeltopic

case "userdel"
userdel

case "userok"
userok

case "fix"

conn.execute("update [user] set userpass='"&md5("123")&"'  where username='"&username&"'")
error2("�Ѿ��� "&username&" �����뻹ԭ�� 123 ")

case "sendmoney"

membercode=int(Request("membercode"))
sql="select username from [user] where membercode="&membercode&""
Set Rs=Conn.Execute(sql)
do while not rs.eof
Count=Count+1

content=HTMLEncode(Request.Form("content"))

conn.Execute("insert into message (author,incept,content) values ('"&Request.Cookies("username")&"','"&rs("username")&"','��ϵͳ�㲥����"&content&"')")

conn.execute("update [user] set newmessage=newmessage+1,[money]=[money]+"&int(Request("money"))&" where username='"&rs("username")&"'")
rs.movenext
loop
rs.close
error2("������ɣ�")


case "activationok"
for each ho in request.form("id")
conn.execute("update [user] set membercode=1 where id="&ho&"")
next
error2("�Ѿ���������ѡ�û���")

end select





sub agreement
%><form method="post" action="?menu=agreementok">
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle>ע���û�Э������</td>
  </tr>
    <tr>
    <td class=a3 align=middle>
<textarea name="agreement" rows="18" style="width:90%"><%=Conn.Execute("Select agreement From [clubconfig]")(0)%></textarea></td>
  </tr>
        
  </table>
<input type="submit" value=" �� �� ">
</form>
<%
end sub


sub showall
%>
�û����ϣ�<b><font color=red><%=conn.execute("Select count(id)from [user]")(0)%></font></b> ��
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 colspan=2 align="center">�û�����</td>
  </tr>


<tr class=a4>
<td><form method="post" action="?menu=user2">
�������Ա������: <input size="13" name="username">
<input type="submit" value="ȷ��">

</td></form>
<td>
<form method="post" action="?">

<input type=hidden name=menu value=showallok>
�������� <select onchange="javascript:submit()" size="1" name="userSearch">
<option value="">��ѡ���ѯ����</option>


<option value="landtime">24Сʱ�ڵ�¼���û�</option>
<option value="regtime">24Сʱ��ע����û�</option>
<option value=" and membercode=1">������ͨ��Ա</option>
<option value=" and membercode=2">���й����Ա</option>
<option value=" and membercode=4">���й���Ա</option>
<option value=" and membercode=5">������������</option>
<option value=" and experience<2">����ֵ����2���û�</option>
<option value=" and degree<2">��¼��������2���û�</option>
</select>



</td>
</tr>

  <tr height=25>
    <td class=a1 align=center colspan=2>�߼���ѯ</td>
  </tr>




  <tr height=25 class=a3>
    <td>�����ʾ��¼��</td>
    <td><input size="45" value="50" name="searchMax"></td>
  </tr>


  <tr height=25 class=a3>
    <td>�û�������</td>
    <td><input size="45" name="username"></td>
  </tr>




  <tr height=25 class=a3>
    <td>������Ϣ����</td>
    <td><input size="45" name="UserInfo"></td>
  </tr>


  <tr height=25 class=a3>
    <td>��ϵ��Ϣ����</td>
    <td><input size="45" name="UserIM"></td>
  </tr>


  <tr height=25 class=a3>
    <td>Email����</td>
    <td><input size="45" name="usermail"></td>
  </tr>


  <tr height=25 class=a3>
    <td>��ҳ����</td>
    <td><input size="45" name="userhome"></td>
  </tr>
  <tr height=25 class=a3>
    <td>ǩ������</td>
    <td><input size="45" name="sign"></td>
  </tr>


  <tr height=25 class=a3>
    <td>ע������</td>
    <td><input size="45" name="regtime"></td>
  </tr>


  <tr height=25 class=a3>
    <td colspan="2" align="center">
	<input type="submit" value="   ��  ��   " name="submit0"></td>
  </tr>

</form>
  </table>



</div>



<br>
<%
end sub







sub showallok
if Request.form<>empty then session("temp")=Request.form
%>
<script>
function CheckAll(form){for (var i=0;i<form.elements.length;i++){var e = form.elements[i];if (e.name != 'chkall')e.checked = form.chkall.checked;}}
</script>
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
<TR align=middle>
<TD class=a1>�û���</TD>
<TD class=a1 height="25">Email</TD>

<TD class=a1>��ѶϢ</TD>

<TD class=a1 height="25">Ȩ��</TD>

<TD class=a1>ע��ʱ��</TD>
<TD class=a1>����¼ʱ��</TD>
<TD class=a1>ɾ��</TD>
</TR>
<form method="post" action="?menu=userdel">
<%

item=Request("userSearch")

if item="landtime" then item=" and landtime >"&SqlNowString&"-1"

if item="regtime" then item=" and regtime >"&SqlNowString&"-1"

if username<>"" then item=""&item&" and username like '%"&username&"%'"

if Request("usermail")<>"" then item=""&item&" and usermail like '%"&Request("usermail")&"%'"

if Request("userhome")<>"" then item=""&item&" and userhome like '%"&Request("userhome")&"%'"

if Request("UserInfo")<>"" then item=""&item&" and UserInfo like '%"&Request("UserInfo")&"%'"

if Request("UserIM")<>"" then item=""&item&" and UserIM like '%"&Request("UserIM")&"%'"

if Request("sign")<>"" then item=""&item&" and sign like '%"&Request("sign")&"%'"

if Request("regtime")<>"" then item=""&item&" and regtime like '%"&Request("regtime")&"%'"

if item="" or Request("searchMax")="" then error2("��������Ҫ����������")

item="where"&item&""
item=replace(item,"where and","where")

sql="select top "&Request("searchMax")&" * from [user] "&item&""
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
i=i+1


%>
<TR align=middle>
<TD class=a4><a target="_blank" href=Profile.asp?username=<%=rs("username")%>><%=rs("username")%></a>��</TD>
<TD class=a3><%=rs("usermail")%>��</TD>
<TD class=a4><a style=cursor:hand onclick="javascript:open('friend.asp?menu=post&incept=<%=rs("username")%>','','width=320,height=170')"><img border="0" src="images/message1.gif"></a></TD>
<TD class=a3>
<a href="admin_user.asp?menu=user2&username=<%=rs("username")%>">�༭</a></TD>
<TD class=a4><%=rs("regtime")%>��</TD>
<TD class=a3><%=rs("landtime")%>��</TD>
<TD class=a3><input type="checkbox" value="<%=rs("username")%>" name="username"></TD></TR>



<%
RS.MoveNext
loop
RS.Close
%>

	<TR align=middle class=a3>
<TD colspan="7" align="right">&nbsp;<input onclick="checkclick('��ȷ��Ҫɾ������ѡ�û���ȫ������?');"  type="submit" value="   ȷ ��   "> <input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON">&nbsp;</TD></form>
	</TR>
</TABLE><br>

<b>[<script>ShowPage(<%=TotalPage%>,<%=PageCount%>,"<%=session("temp")%>")</script>]</b>

<%
end sub

sub user


%>

�û����ϣ�<b><font color=red><%=conn.execute("Select count(id)from [user]")(0)%></font></b> ��
<br>
<form method="post" action="?menu=sendmoney">
<table cellspacing="1" cellpadding="2" width="64%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>���Ź���</td>
  </tr>
  <tr height=25>
    <td class=a3 align=middle colspan=2>������
<select size="1" name="membercode">
<option value="1">��ͨ�û�</option>
<option value="2">�����Ա</option>
<option value="4" selected>�� �� Ա</option>
<option value="5">��������</option>
</select>&nbsp;���ű��¹��� <input size="8" name="money" value="1000">
���
<input type="submit" value="ȷ��"> </td>
  </tr>
  <tr height=25 class=a3>
    <td align=middle width="13%">��<br>
		Ѷ<br>
		��<br>
		��</td>
    <td align=middle width="85%"><textarea name="content" rows="5" cols="55">���¹����Ѿ����͵������ʺ��У���ע����գ�</textarea></td>
  </tr>
   </table>

</form>
<%
end sub

sub user2

sql="select * from [user] where username='"&HTMLEncode(username)&"'"
Set Rs=Conn.Execute(sql)
if rs.eof then error2(""&username&" ���û����ϲ�����")

UserInfo=split(rs("UserInfo"),"\")
realname=UserInfo(0)
country=UserInfo(1)
province=UserInfo(2)
city=UserInfo(3)
postcode=UserInfo(4)
blood=UserInfo(5)
belief=UserInfo(6)
occupation=UserInfo(7)
marital=UserInfo(8)
education=UserInfo(9)
college=UserInfo(10)
address=UserInfo(11)
phone=UserInfo(12)
character=UserInfo(13)
personal=UserInfo(14)





%>
<form method="post" name=form action="?menu=userok">
<input type=hidden name=username value="<%=rs("username")%>">
<div align="center">

<table cellSpacing="1" cellPadding="3" border="0" width="60%" class=a2>
<tr class=a1 id=TableTitleLink>
<td height="25" width="600" colspan="4" align="center">
<font color="000000">
<a target="_blank" href="Profile.asp?username=<%=rs("username")%>">
�鿴��<%=rs("username")%>������ϸ����</a></font></td>
</tr>
<tr class=a3>
<td colspan="2">&nbsp;�û����룺<a onclick="checkclick('�˲�������Ѹ��û�������ĳɣ�123');" href="?menu=fix&username=<%=rs("username")%>">��ԭ����</a></td>
<td width="300" colspan="2" height="25">&nbsp;�û�Ȩ�ޣ�<select size="1" name="membercode">
<option value=0 <%if rs("membercode")=0 then%>selected<%end if%>>��δ����</option>
<option value=1 <%if rs("membercode")=1 then%>selected<%end if%>>��ͨ��Ա</option>
<option value=2 <%if rs("membercode")=2 then%>selected<%end if%>>�����Ա</option>
<option value=4 <%if rs("membercode")=4 then%>selected<%end if%>>�� �� Ա</option>
<option value=5 <%if rs("membercode")=5 then%>selected<%end if%>>��������</option>
</select>
</td>
</tr>
<tr class=a4>
<td colspan="2">&nbsp;�û�ͷ�Σ�<input size="10" name="honor" value="<%=rs("honor")%>"></td>
<td width="300" colspan="2">&nbsp;�������ɣ�<input size="10" name="faction" value="<%=rs("faction")%>"></td>
</tr>
<tr class=a3>
<td colspan="2">&nbsp;�� �� ֵ��<input size="10" name="userlife" value="<%=rs("userlife")%>"></td>
<td width="300" colspan="2">&nbsp;�������£�<input size="10" name="posttopic" value="<%=rs("posttopic")%>"></td>
</tr>
<tr class=a4>
<td colspan="2">&nbsp;������ң�<input size="10" name="money" value="<%=rs("money")%>"></td>
<td width="300" colspan="2">&nbsp;�ظ����£�<input size="10" name="postrevert" value="<%=rs("postrevert")%>"></td>
</tr>
<tr class=a3>
<td colspan="2">&nbsp;������<input size="10" name="savemoney" value="<%=rs("savemoney")%>"></td>
<td width="300" colspan="2">&nbsp;��ɾ���ӣ�<input size="10" name="deltopic" value="<%=rs("deltopic")%>"></td>
</tr>
<tr class=a3>
<td colspan="2">&nbsp;ע�����ڣ�<input size="10" name="regtime" value="<%=rs("regtime")%>"></td>
<td width="300" colspan="2">&nbsp;�� �� ֵ��<input size="10" name="experience" value="<%=rs("experience")%>"></td>
</tr>

<tr class=a3>
<td colspan="4">&nbsp;�û�ͷ��<input size="50" name="userface" value="<%=rs("userface")%>"></td>
</tr>

<tr class=a3>
<td colspan="4">&nbsp;�û���Ƭ��<input size="50" name="userphoto" value="<%=rs("userphoto")%>"></td>
</tr>

<tr class=a1 id=TableTitleLink>
<td colspan="4" align="center" height="25">
��������</td>
</tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;��ʵ������<input type="text" name="realname" size="10" value="<%=realname%>"><td width="50%" height="3" colspan="2">
&nbsp;�ԡ�����<input type="text" name="sex" size="10" value="<%=rs("sex")%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;�������ڣ�<input type="text" name="birthday" size="10" value="<%=rs("birthday")%>"><td width="50%" height="3" colspan="2">
&nbsp;�������ң�<b><input type="text" name="country" size="10" value="<%=country%>"></b></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;ʡ�����ݣ�<input type="text" name="province" size="10" value="<%=province%>"><td width="50%" height="3" colspan="2">
&nbsp;�ǡ����У�<input type="text" name="city" size="10" value="<%=city%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;������ţ�<input type="text" name="postcode" size="10" value="<%=postcode%>"><td width="50%" height="3" colspan="2">
&nbsp;Ѫ�����ͣ�<input maxlength=4 size=10 name=blood value="<%=blood%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;�š�������<input type="text" name="belief" size="10" value="<%=belief%>"><td width="50%" colspan="2">
&nbsp;ְ����ҵ��<input type="text" name="occupation" size="10" value="<%=occupation%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;����״����<input maxlength=4 size=10 name=marital value="<%=marital%>"><td width="50%" colspan="2">
&nbsp;���ѧ����<input type="text" name="education" size="10" value="<%=education%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;��ҵԺУ��<input type="text" name="college" size="20" value="<%=college%>"><td width="50%" colspan="2">
&nbsp;�绰���룺<input type="text" name="phone" size="10" value="<%=phone%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;��ϵ��ַ��<input type="text" name="address" size="20" value="<%=address%>"><td width="50%" colspan="2">
&nbsp;�ֻ����룺<input type="text" name="UserMobile" size="10" value="<%=rs("UserMobile")%>"></tr>

<tr class=a4>
<td width="600" colspan="4">&nbsp;�û�ǩ����<textarea name="sign" rows="4" cols="50"><%=rs("sign")%></textarea></td>
</tr>

<tr class=a1 id=TableTitleLink>
<td width="50%" align="center" height="13">
<a onclick="checkclick('��ȷ��Ҫɾ�����û����з����������?');" href="?menu=userdeltopic&username=<%=rs("username")%>">
ɾ�����û�����������</a>


<td width="201" colspan="2" align="center" height="13">
<input type="submit" value=" �� �� " name="Submit"></td>
<td width="50%" align="center" height="13">

<a onclick="checkclick('��ȷ��Ҫɾ�����û�����������?');" href="?menu=userdel&username=<%=rs("username")%>">
ɾ�����û�����������</a></td>

</td>
</tr>


</table>
</form><A href="javascript:history.back()">�� ��</A>
<Script>
document.form.sign.value = unybbcode(document.form.sign.value);
function unybbcode(temp) {
temp = temp.replace(/<br>/ig,"\n");
return (temp);
}
</Script>
<%
end sub

sub userdeltopic
conn.execute("delete from [forum] where username='"&username&"'")
conn.execute("delete from [reforum"&ReForumName&"] where username='"&username&"'")
%>
�Ѿ��� <%=username%> ���з����������ȫ��ɾ��<br><br><a href=javascript:history.back()>�� ��</a>
<%
end sub

sub userdel

if username=Request.Cookies("username") then error2("��������")

for each ho in Request("username")
ho=HTMLEncode(ho)
conn.execute("delete from [user] where username='"&ho&"'")
conn.execute("delete from [reforum"&ReForumName&"] where username='"&ho&"'")
next

%>
�Ѿ��ɹ�ɾ�� <%=username%> ����������<br><br><a href=javascript:history.back()>�� ��</a>
<%
end sub


sub userok

if Request("userlife")>100 then error2("�������ܴ���100��")

sql="select * from [user] where username='"&username&"'"
rs.Open sql,Conn,1,3
rs("userface")=Request("userface")
rs("userphoto")=Request("userphoto")
rs("membercode")=Request("membercode")
rs("honor")=Request("honor")
rs("faction")=Request("faction")
rs("posttopic")=Request("posttopic")
rs("postrevert")=Request("postrevert")
rs("experience")=Request("experience")
rs("userlife")=Request("userlife")
rs("money")=Request("money")
rs("savemoney")=Request("savemoney")
rs("regtime")=Request("regtime")
rs("deltopic")=Request("deltopic")
rs("birthday")=Request("birthday")
rs("sign")=HTMLEncode(Request.Form("sign"))
rs("sex")=HTMLEncode(Request.Form("sex"))

rs("UserInfo")=""&HTMLEncode(Request("realname"))&"\"&HTMLEncode(Request("country"))&"\"&HTMLEncode(Request("province"))&"\"&HTMLEncode(Request("city"))&"\"&HTMLEncode(Request("postcode"))&"\"&HTMLEncode(Request("blood"))&"\"&HTMLEncode(Request("belief"))&"\"&HTMLEncode(Request("occupation"))&"\"&HTMLEncode(Request("marital"))&"\"&HTMLEncode(Request("education"))&"\"&HTMLEncode(Request("college"))&"\"&HTMLEncode(Request("address"))&"\"&HTMLEncode(Request("phone"))&"\"&HTMLEncode(Request("character"))&"\"&HTMLEncode(Request("personal"))&""
rs("UserMobile")=""&HTMLEncode(Request("UserMobile"))&""


rs.update
rs.close
%> ���³ɹ�</b></font></td>
</tr></table><br><br><a href=javascript:history.back()>�� ��</a>
<%
end sub

sub Activation
%>
<script>
function CheckAll(form){for (var i=0;i<form.elements.length;i++){var e = form.elements[i];if (e.name != 'chkall')e.checked = form.chkall.checked;}}
</script>
  <form method="POST" action=?menu=activationok>

<TABLE cellSpacing=1 cellPadding=1 width=60% align=center border=0 class=a2>
<TR height=25 class=a1>
		<td align="center">����</td>
		<td align="center">�û���</td>
		<td align="center">Email</td>
		<td align="center">ע��ʱ��</td></tr>
<%



sql="select * from [user] where membercode=0 order by id Desc"
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
i=i+1

%>
<TR height=25>

      <TD align=middle width="6%" class=a3>
<INPUT type=checkbox value=<%=rs("id")%> name=id></TD>
      <TD width="22%" align=left class=a4>&nbsp;<a target="_blank" href=Profile.asp?username=<%=rs("username")%>><%=rs("username")%></a></TD>
      <TD width="45%" class=a3>&nbsp;<%=rs("usermail")%></TD>

      <TD align=center width=30% class=a4>&nbsp;<%=rs("regtime")%></TD>
    </TR>

<%

RS.MoveNext
loop
RS.Close


%>


<TR height=25>

      <TD width="6%" class=a3 align="center"><input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON"></TD>

      <TD width="22%" class=a3 align="center">		
<input type=submit  onclick="checkclick('��ȷ��Ҫ������ѡ���û�?');" value=" �� �� "></TD>

      <TD width="71%" class=a3 colspan="2">
		
		
&nbsp;<b>����
<font color="990000"><%=TotalPage%></font> ҳ [
<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"menu=Activation")
</script>]</b>
		
</TD>
    </TR>
    


</table><%
end sub

htmlend

%>