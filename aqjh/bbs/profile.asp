<!-- #include file="setup.asp" -->
<%
top
username=HTMLEncode(Request("username"))


sql="select * from [user] where username='"&username&"'"
Set Rs=Conn.Execute(sql)

if rs.eof then error("<li>"&username&"���û����ϲ�����")

select case rs("sex")
case "male"
sex="��"
case "female"
sex="Ů"
end select



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

UserIM=split(rs("UserIM"),"\")
qq=UserIM(0)
icq=UserIM(1)
uc=UserIM(2)
aim=UserIM(3)
msn=UserIM(4)
Yahoo=UserIM(5)


%>
<title><%=rs("username")%>�û����� - Powered By BBSxp</title>
<script src="inc/birth.js"></script>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� �鿴�û���<%=rs("username")%>������</td>
</tr>
</table><br>
<table width=97% border="0" cellspacing="0" cellpadding="5" align="center">
<tr>
<td><img src="<%=rs("userface")%>" width="32" height="32">��<b><%=rs("username")%></b></td>
<td align="right" valign="bottom"><b>:::��ز���:::</b>����
<img src="images/finds.gif"> 
<a href=ShowBBS.asp?menu=5&username=<%=rs("username")%>> ��������</a>��

<img src="images/sig.gif">
<a href="blog.asp?username=<%=rs("username")%>">������־</a>��

<img src="images/friend1.gif"> <a href=friend.asp?menu=add&username=<%=rs("username")%>> ��Ϊ����</a>��
<img src="images/message1.gif"> <a href=# onclick="javascript:open('friend.asp?menu=post&incept=<%=rs("username")%>','','width=320,height=170')"> ����ѶϢ</a> ��
</tr>
</table>

<SCRIPT>valigntop()</SCRIPT>
<table width=97% align=center cellspacing=0 cellpadding=0 border=0>
<tr>
<td class=a2 height="676">
<table width=100% cellspacing=1 cellpadding=6 border=0>
<tr bgcolor="FFFFFF" class=a1>
<td colspan="2" height="25" valign="middle" align="left"><b>&nbsp;:::������Ϣ:::</b></td>
</tr>
<tr>
<td class=a3 height="5" align="left" valign="middle" width="50%"><b>���ǡ����ƣ�</b><%=rs("username")%></td>
<td class=a3 height="5" align="left" valign="middle" width="50%"><b>������ԭ����</b><%=rs("posttopic")%></td>
</tr>
<tr>
<td class=a4 align="left" valign="middle" width="50%"><b>
���ȼ����ƣ�</b><Script>document.write(level(<%=rs("experience")%>,<%=rs("membercode")%>,'','')+levelname);</Script></td>
<td class=a4 align="left" valign="middle" width="50%"><b>�����������</b><%=rs("postrevert")%></td>
</tr>
<tr>
<td align="left" width="50%" class=a3><b>���š����ɣ�</b><%
if rs("faction")=empty then
response.write "��"
else
response.write rs("faction")
end if
%></td>
<td align="left" width="50%" class=a3>

<b>����¼������</b><%=rs("goodtopic")%></td>
</tr>
<tr class=a4>
<td align="left" width="50%"><b>���䡡��ż��</b><%
if rs("consort")=empty then
response.write "��"
else
response.write "<a href=Profile.asp?username="&rs("consort")&">"&rs("consort")&"</a>"
end if
%></td>
<td align="left" width="50%">

<b>����ɾԭ����</b><%=rs("deltopic")%></td>
</tr>
<tr class=a3>
<td align="left" width="50%"><b>
��ע�����ڣ�</b><%=rs("regtime")%></td>
<td align="left" width="50%"><b>��������ң�</b><%=rs("money")%></td>
</tr>
<tr class=a4>
<td align="left" width="50%"><b>
���ϴε�¼��</b><%=rs("landtime")%></td>
<td align="left" width="50%"><b>
���� �� ֵ��</b><%=rs("userlife")%></td>
</tr>
<tr class=a3>
<td align="left" width="50%"><b>
����¼������</b><%=rs("degree")%></td>
<td align="left" width="50%"><b>���� �� ֵ��</b><%=rs("experience")%></td>
</tr>

<tr bgcolor="FFFFFF" class=a1>
<td height="25" align="left" valign="middle" colspan="2"><b>&nbsp;:::������Ϣ:::</b></td>
</tr>
<tr>
<td valign="top" class=a3 width="50%"> ��<b>��ʵ������</b><%=realname%>
</td>
<td height="100%" align="left" valign="bottom" class=a3 rowspan="17" width="50%">
<div style="OVERFLOW:auto; HEIGHT:500px">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">

<tr><td valign=top align=center height="100%"><div align=left><b> ���û���Ƭ��</b></div>
<script>
if("<%=rs("userphoto")%>"!=""){
document.write("<a target=_blank href=<%=rs("userphoto")%>><img src=<%=rs("userphoto")%> border=0 onload='javascript:if(this.width>300)this.width=300'></a>")
}
</script>

</td></tr>

<tr>
<td height="100">

<table width="100%" border="0" cellspacing="0" cellpadding="1">
<tr>
<td class=a2>
<table width="100%" border="0" cellspacing="0" cellpadding="10">
<tr>
<td class=a4 height="100" valign="top"><b>
&nbsp;�ԡ�����</b><br>
��<table border="0">
<tr>
<td width="100%"><%=character%></td>
</tr>
</table>
</td>
</tr>
</table>
</td>
</tr>
</table>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="1">
<tr>
<td class=a2 height="100">
<table width="100%" border="0" cellspacing="0" cellpadding="10">
<tr>
<td class=a4 height="100" valign="top"><b>&nbsp;�򡡽飺</b><br>
<br>
<%=personal%> </td>
</tr>
</table>
</td>
</tr>
</table>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="1">
<tr>
<td class=a2>
<table width="100%" border="0" cellspacing="0" cellpadding="10">
<tr>
<td class=a4 height="100" valign="top"><b>&nbsp;ǩ������</b><br>
<br><script>document.write(ybbcode("<%=rs("sign")%>"))</script>
</td>
</tr>
</table>
</td>
</tr>
</table>


</td>
</tr>


</table></div>
</td>
</tr>
<tr>
<td valign="top" class=a4 width="50%"><b>���ԡ�����</b><%=sex%> </td>
</tr>
<tr>
<td valign="top" class=a3 width="50%"><b> ���������ң�</b><%=country%>
</td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">��<b>ʡ�����ݣ�</b><%=province%> </td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">��<b>�ǡ����У�</b><%=city%></td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">��<b>�������룺</b><%=postcode%></td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">��<b>������Ф��</b><script>document.write(getpet("<%=rs("birthday")%>"));</script></td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">��<b>Ѫ�����ͣ�</b><%=blood%> </td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">��<b>�ǡ�������</b><script>document.write(astro("<%=rs("birthday")%>"));</script></td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">��<b>�š�������</b><%=belief%>
</td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">��<b>ְ����ҵ��</b><%=occupation%>
</td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">��<b>����״����</b><%=marital%>
</td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">��<b>���ѧ����</b><%=education%>
</td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">��<b>��ҵԺУ��</b><%=college%>
</td>
</tr>
<tr class=a1>
<td valign="top" bgcolor="FFFFFF" width="50%"><b>&nbsp;:::��ϵ����:::</b></td>
</tr>
<tr>
<td valign="top" class=a4 width="50%">��<b>�����ʼ���</b><a href="mailto:<%=rs("usermail")%>"><%=rs("usermail")%></a></td>
</tr>
<tr>
<td valign="top" class=a3 width="50%">��<b>������ҳ��</b><a target="_blank" href="<%=rs("userhome")%>"><%=rs("userhome")%></a></td>
</tr>
<tr bgcolor="FFFFFF" class=a1>
<td colspan="2" height="25" align="left"><b>&nbsp;:::��ʱͨѶ:::</b></td>
</tr>
<tr bgcolor="FFFFFF" class=a3>
<td height="4" valign="middle" align="left">��<b>QQ���룺</b><%if qq<>"" then%><a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin=<%=qq%>&Menu=yes&Site=<%=clubname%>><%=qq%></a><%end if%>
</td>
<td height="4" valign="middle" align="left">��<b>ICQ��</b><%=icq%></td>
</tr>
<tr bgcolor="FFFFFF" class=a4>
<td height="4" valign="middle" align="left">��<b>UC���룺</b><%=uc%></td>
<td height="4" valign="middle" align="left">��<b>AIM��</b><%=aim%></td>
</tr>
<tr bgcolor="FFFFFF" class=a3>
<td height="4" valign="middle" align="left">��<b>MSN IM��</b><%=msn%>��</td>
<td height="4" valign="middle" align="left">��<b>Yahoo IM��</b><%=Yahoo%></td>
</tr>





</table>

</td>
</tr>
</table>


<%rs.close%>

<center>
<SCRIPT>valignbottom()</SCRIPT>

<br>
<INPUT onclick=history.back(-1) type=button value=" << �� �� ">
<br><%htmlend%>