<!-- #include file="setup.asp" --><%
id=int(Request("id"))

content=HTMLEncode(Request.Form("content"))

top

if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")


consort=Conn.Execute("Select consort From [user] where username='"&Request.Cookies("username")&"'")(0)



select case Request("menu")
case "add"
aim=HTMLEncode(Request("aim"))
if content=empty then error("<li>表白内容不能为空")
if aim=Request.Cookies("username") then error("<li>不能自己追求自己！")

If conn.Execute("Select id From [user] where username='"&aim&"'" ).eof Then error("<li>系统不存在"&aim&"的资料")

If Not conn.Execute("Select id From [consort] where username='"&Request.Cookies("username")&"' and aim='"&aim&"'" ).eof Then error("<li>"&aim&"已经在您的追求列表当中")

sql="insert into consort (username,aim,unburden) values ('"&Request.Cookies("username")&"','"&aim&"','"&content&"')"
conn.Execute(SQL)


sql="insert into message(author,incept,content) values ('"&Request.Cookies("username")&"','"&aim&"','<font color=0000FF>【社区配偶】："&content&"</font>')"
conn.Execute(SQL)


conn.execute("update [user] set newmessage=newmessage+1 where username='"&aim&"'")

case "accept"

if consort<>empty then error("<li>您当前已有配偶")

aim=Conn.Execute("Select aim From [consort] where id="&id&"")(0)
if aim<>Request.Cookies("username") then error("<li>非法操作")

consort=Conn.Execute("Select username From [consort] where id="&id&"")(0)
if Conn.Execute("Select consort From [user] where username='"&consort&"'")(0)<>empty then error("<li>"&consort&"已经有配偶了")

conn.execute("update [user] set consort='"&aim&"' where username='"&consort&"'")
conn.execute("update [user] set consort='"&consort&"' where username='"&aim&"'")
conn.execute("delete from [consort] where id="&id&"")
succeed("<li>您已经接受了"&consort&"的追求<li><a href=consort.asp>返回社区配偶</a><meta http-equiv=refresh content=3;url=consort.asp>")



case "del"
conn.execute("delete from [consort] where id="&id&" and (username='"&Request.Cookies("username")&"' or aim='"&Request.Cookies("username")&"') ")

case "part"
conn.execute("update [user] set consort='' where username='"&consort&"'")
conn.execute("update [user] set consort='' where username='"&Request.Cookies("username")&"'")
succeed("<li>分手成功<li><a href=consort.asp>返回社区配偶</a><meta http-equiv=refresh content=3;url=consort.asp>")

end select

%>
<table border="0" width="97%" align="center" cellspacing="1" cellpadding="4" class="a2">
	<tr class="a3">
		<td height="25">&nbsp;<img src="images/Forum_nav.gif">&nbsp; <%ClubTree%> 
		→ <a href="consort.asp">社区配偶</a></td>
	</tr>
</table>
<br>
<center>
<script>valigntop()</script>
<table cellspacing="1" cellpadding="6" width="97%" border="0" class="a2">
	<tr>
		<td width="100%" height="10" align="middle" class="a1" colspan="4">
		<font face="宋体">我的追求者</font></td>
	</tr>
	<tr class="a3">
		<td width="10%" height="5" align="middle">用户名</td>
		<td height="5" align="middle">爱的表白</td>
		<td width="20%" height="5" align="middle"><font face="宋体">追求</font>时间</td>
		<td width="15%" height="5" align="middle">操作</td>
	</tr>
	<%
sql="select * from consort where aim='"&Request.Cookies("username")&"' order by id Desc"
Set Rs=Conn.Execute(sql)
Do While Not RS.EOF 
%>
	<tr class="a4">
		<td height="5" align="middle">
		<a href="Profile.asp?username=<%=rs("username")%>"><%=rs("username")%></a></td>
		<td height="5" align="middle"><%=rs("unburden")%></td>
		<td height="5" align="middle"><%=rs("addtime")%></td>
		<td height="5" align="middle"><%if consort=rs("username") then%>
		<a href="?menu=part">分 手</a> <%else%>
		<a href="?menu=accept&id=<%=rs("id")%>">接受</a>
		<a href="?menu=del&id=<%=rs("id")%>">拒绝</a> <%end if%></td>
	</tr>
	<%
RS.MoveNext
loop
RS.Close
%>
</table>
<script>valignbottom()</script>
<br>
<script>valigntop()</script>
<table cellspacing="1" cellpadding="6" width="97%" border="0" class="a2">
	<tr class="a1">
		<td width="100%" height="10" align="middle" colspan="4">我追求的人</td>
	</tr>
	<tr class="a3">
		<td width="10%" height="5" align="middle">用户名</td>
		<td height="5" align="middle">爱的表白</td>
		<td width="20%" height="5" align="middle"><font face="宋体">追求</font>时间</td>
		<td width="15%" height="5" align="middle">操作</td>
	</tr>
	<%
sql="select * from consort where username='"&Request.Cookies("username")&"' order by id Desc"
Set Rs=Conn.Execute(sql)
Do While Not RS.EOF 
%>
	<tr class="a4">
		<td height="5" align="middle">
		<a href="Profile.asp?username=<%=rs("aim")%>"><%=rs("aim")%></a></td>
		<td height="5" align="middle"><%=rs("unburden")%></td>
		<td height="5" align="middle"><%=rs("addtime")%></td>
		<td height="5" align="middle"><%if consort=rs("aim") then%>
		<a href="?menu=part">分 手</a> <%else%>
		<a href="?menu=del&id=<%=rs("id")%>">取消追求</a> <%end if%> </td>
	</tr>
	<%
RS.MoveNext
loop
RS.Close

%>
</table>
<script>valignbottom()</script>

<br>
<script>valigntop()</script>
<%if consort<>empty then

sql="select * from [user] where username='"&consort&"'"
Set Rs=Conn.Execute(sql)

if rs.eof then conn.execute("update [user] set consort='' where username='"&Request.Cookies("username")&"'")

select case rs("sex")
case "male"
sex="男"
case "female"
sex="女"
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
userphoto=rs("userphoto")
rs.close
%>
<table cellspacing="1" cellpadding="6" width="97%" border="0" class="a2">
	<tr class="a1" id="TableTitleLink">
		<td align="middle" class="a1" colspan="7">我的配偶</td>
	</tr>
	<tr>
		<td width="20%" align="left" class="a4" rowspan="3"><script>
if("<%=userphoto%>"!=""){
document.write("<img src=<%=userphoto%> border=0 onload='javascript:if(this.width>200)this.width=200'>")
}
</script>
		
		</td>
		<td width="80" align="left" class="a4">昵称：</td>
		<td align="left" class="a4" width="15%"><a href="Profile.asp?username=<%=consort%>">
		<%=consort%></a></td>
		<td width="80" align="left" class="a4">姓名：</td>
		<td align="left" class="a4" width="15%"><%=realname%></td>
		<td width="80" align="left" class="a4">性别：</td>
		<td align="left" class="a4" width="15%"><%=sex%></td>
	</tr>
	<tr>
		<td width="80" align="left" class="a4" valign="top">国家：</td>
		<td align="left" class="a4" width="15%"><%=country%></td>
		<td width="80" align="left" class="a4">省份：</td>
		<td align="left" class="a4" width="15%"><%=province%></td>
		<td width="80" align="left" class="a4">城市：</td>
		<td align="left" class="a4" width="15%"><%=city%></td>
	</tr>
	<tr>
		<td width="80" align="left" class="a4" valign="top">个人说明：</td>
		<td align="left" class="a4" colspan="5"><%=personal%></td>
	</tr>
	<tr>
		<td align="right" class="a4" valign="top" colspan="7">
		<a onclick="checkclick('您确定要与当前配偶分手？')" href="?menu=part">与当前配偶分手</a></td>
	</tr>
</table>
<%else%>
<table cellspacing="1" cellpadding="6" width="97%" border="0" class="a2">
	<form action="consort.asp" method="post">
		<input type="hidden" value="add" name="menu">
		<tr>
			<td width="77%" height="2" align="middle" class="a1" colspan="2">添加我想追求的人</td>
		</tr>
		<tr>
			<td width="12%" height="2" align="left" class="a4">对方用户名：</td>
			<td width="64%" height="2" align="left" class="a4">
			<input maxlength="30" size="15" name="aim"></td>
		</tr>
		<tr>
			<td width="12%" height="1" align="left" class="a4" valign="top">爱的表白：</td>
			<td width="64%" height="1" align="left" class="a4">
			<textarea name="content" rows="5" style="width:95%"></textarea></td>
		</tr>
		<tr>
			<td width="77%" height="1" align="center" class="a4" colspan="2">
			<input tabindex="4" type="submit" value=" 确 定 ">
			<input onclick="checkclick('该项操作要清除全部的内容，您确定要清除吗?');" type="reset" value=" 重 写 "></td>
		</tr>
	</form>
</table>
<%end if%>
<script>valignbottom()</script>
</center><%htmlend%>