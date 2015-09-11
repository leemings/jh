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
更新成功<br><br><a href=javascript:history.back()>返 回</a>
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
error2("已经将 "&username&" 的密码还原成 123 ")

case "sendmoney"

membercode=int(Request("membercode"))
sql="select username from [user] where membercode="&membercode&""
Set Rs=Conn.Execute(sql)
do while not rs.eof
Count=Count+1

content=HTMLEncode(Request.Form("content"))

conn.Execute("insert into message (author,incept,content) values ('"&Request.Cookies("username")&"','"&rs("username")&"','【系统广播】："&content&"')")

conn.execute("update [user] set newmessage=newmessage+1,[money]=[money]+"&int(Request("money"))&" where username='"&rs("username")&"'")
rs.movenext
loop
rs.close
error2("发放完成！")


case "activationok"
for each ho in request.form("id")
conn.execute("update [user] set membercode=1 where id="&ho&"")
next
error2("已经将激活所选用户！")

end select





sub agreement
%><form method="post" action="?menu=agreementok">
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle>注册用户协议设置</td>
  </tr>
    <tr>
    <td class=a3 align=middle>
<textarea name="agreement" rows="18" style="width:90%"><%=Conn.Execute("Select agreement From [clubconfig]")(0)%></textarea></td>
  </tr>
        
  </table>
<input type="submit" value=" 更 新 ">
</form>
<%
end sub


sub showall
%>
用户资料：<b><font color=red><%=conn.execute("Select count(id)from [user]")(0)%></font></b> 条
<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 colspan=2 align="center">用户管理</td>
  </tr>


<tr class=a4>
<td><form method="post" action="?menu=user2">
请输入会员的名称: <input size="13" name="username">
<input type="submit" value="确定">

</td></form>
<td>
<form method="post" action="?">

<input type=hidden name=menu value=showallok>
快速搜索 <select onchange="javascript:submit()" size="1" name="userSearch">
<option value="">请选择查询条件</option>


<option value="landtime">24小时内登录的用户</option>
<option value="regtime">24小时内注册的用户</option>
<option value=" and membercode=1">所有普通会员</option>
<option value=" and membercode=2">所有贵宾会员</option>
<option value=" and membercode=4">所有管理员</option>
<option value=" and membercode=5">所有社区区长</option>
<option value=" and experience<2">经验值少于2的用户</option>
<option value=" and degree<2">登录次数少于2的用户</option>
</select>



</td>
</tr>

  <tr height=25>
    <td class=a1 align=center colspan=2>高级查询</td>
  </tr>




  <tr height=25 class=a3>
    <td>最多显示记录数</td>
    <td><input size="45" value="50" name="searchMax"></td>
  </tr>


  <tr height=25 class=a3>
    <td>用户名包含</td>
    <td><input size="45" name="username"></td>
  </tr>




  <tr height=25 class=a3>
    <td>基本信息包含</td>
    <td><input size="45" name="UserInfo"></td>
  </tr>


  <tr height=25 class=a3>
    <td>联系信息包含</td>
    <td><input size="45" name="UserIM"></td>
  </tr>


  <tr height=25 class=a3>
    <td>Email包含</td>
    <td><input size="45" name="usermail"></td>
  </tr>


  <tr height=25 class=a3>
    <td>主页包含</td>
    <td><input size="45" name="userhome"></td>
  </tr>
  <tr height=25 class=a3>
    <td>签名包含</td>
    <td><input size="45" name="sign"></td>
  </tr>


  <tr height=25 class=a3>
    <td>注册日期</td>
    <td><input size="45" name="regtime"></td>
  </tr>


  <tr height=25 class=a3>
    <td colspan="2" align="center">
	<input type="submit" value="   搜  索   " name="submit0"></td>
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
<TD class=a1>用户名</TD>
<TD class=a1 height="25">Email</TD>

<TD class=a1>短讯息</TD>

<TD class=a1 height="25">权限</TD>

<TD class=a1>注册时间</TD>
<TD class=a1>最后登录时间</TD>
<TD class=a1>删除</TD>
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

if item="" or Request("searchMax")="" then error2("请输入您要搜索的条件")

item="where"&item&""
item=replace(item,"where and","where")

sql="select top "&Request("searchMax")&" * from [user] "&item&""
rs.Open sql,Conn,1



pagesetup=20 '设定每页的显示数量
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '总页数
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '跳转到指定页数
i=0
Do While Not RS.EOF and i<pagesetup
i=i+1


%>
<TR align=middle>
<TD class=a4><a target="_blank" href=Profile.asp?username=<%=rs("username")%>><%=rs("username")%></a>　</TD>
<TD class=a3><%=rs("usermail")%>　</TD>
<TD class=a4><a style=cursor:hand onclick="javascript:open('friend.asp?menu=post&incept=<%=rs("username")%>','','width=320,height=170')"><img border="0" src="images/message1.gif"></a></TD>
<TD class=a3>
<a href="admin_user.asp?menu=user2&username=<%=rs("username")%>">编辑</a></TD>
<TD class=a4><%=rs("regtime")%>　</TD>
<TD class=a3><%=rs("landtime")%>　</TD>
<TD class=a3><input type="checkbox" value="<%=rs("username")%>" name="username"></TD></TR>



<%
RS.MoveNext
loop
RS.Close
%>

	<TR align=middle class=a3>
<TD colspan="7" align="right">&nbsp;<input onclick="checkclick('您确定要删除您所选用户的全部资料?');"  type="submit" value="   确 定   "> <input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON">&nbsp;</TD></form>
	</TR>
</TABLE><br>

<b>[<script>ShowPage(<%=TotalPage%>,<%=PageCount%>,"<%=session("temp")%>")</script>]</b>

<%
end sub

sub user


%>

用户资料：<b><font color=red><%=conn.execute("Select count(id)from [user]")(0)%></font></b> 条
<br>
<form method="post" action="?menu=sendmoney">
<table cellspacing="1" cellpadding="2" width="64%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>发放工资</td>
  </tr>
  <tr height=25>
    <td class=a3 align=middle colspan=2>给所有
<select size="1" name="membercode">
<option value="1">普通用户</option>
<option value="2">贵宾会员</option>
<option value="4" selected>管 理 员</option>
<option value="5">社区区长</option>
</select>&nbsp;发放本月工资 <input size="8" name="money" value="1000">
金币
<input type="submit" value="确定"> </td>
  </tr>
  <tr height=25 class=a3>
    <td align=middle width="13%">短<br>
		讯<br>
		内<br>
		容</td>
    <td align=middle width="85%"><textarea name="content" rows="5" cols="55">本月工资已经发送到您的帐号中，请注意查收！</textarea></td>
  </tr>
   </table>

</form>
<%
end sub

sub user2

sql="select * from [user] where username='"&HTMLEncode(username)&"'"
Set Rs=Conn.Execute(sql)
if rs.eof then error2(""&username&" 的用户资料不存在")

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
查看“<%=rs("username")%>”的详细资料</a></font></td>
</tr>
<tr class=a3>
<td colspan="2">&nbsp;用户密码：<a onclick="checkclick('此操作将会把该用户的密码改成：123');" href="?menu=fix&username=<%=rs("username")%>">还原密码</a></td>
<td width="300" colspan="2" height="25">&nbsp;用户权限：<select size="1" name="membercode">
<option value=0 <%if rs("membercode")=0 then%>selected<%end if%>>尚未激活</option>
<option value=1 <%if rs("membercode")=1 then%>selected<%end if%>>普通会员</option>
<option value=2 <%if rs("membercode")=2 then%>selected<%end if%>>贵宾会员</option>
<option value=4 <%if rs("membercode")=4 then%>selected<%end if%>>管 理 员</option>
<option value=5 <%if rs("membercode")=5 then%>selected<%end if%>>社区区长</option>
</select>
</td>
</tr>
<tr class=a4>
<td colspan="2">&nbsp;用户头衔：<input size="10" name="honor" value="<%=rs("honor")%>"></td>
<td width="300" colspan="2">&nbsp;所属帮派：<input size="10" name="faction" value="<%=rs("faction")%>"></td>
</tr>
<tr class=a3>
<td colspan="2">&nbsp;体 力 值：<input size="10" name="userlife" value="<%=rs("userlife")%>"></td>
<td width="300" colspan="2">&nbsp;发表文章：<input size="10" name="posttopic" value="<%=rs("posttopic")%>"></td>
</tr>
<tr class=a4>
<td colspan="2">&nbsp;社区金币：<input size="10" name="money" value="<%=rs("money")%>"></td>
<td width="300" colspan="2">&nbsp;回复文章：<input size="10" name="postrevert" value="<%=rs("postrevert")%>"></td>
</tr>
<tr class=a3>
<td colspan="2">&nbsp;社区存款：<input size="10" name="savemoney" value="<%=rs("savemoney")%>"></td>
<td width="300" colspan="2">&nbsp;被删帖子：<input size="10" name="deltopic" value="<%=rs("deltopic")%>"></td>
</tr>
<tr class=a3>
<td colspan="2">&nbsp;注册日期：<input size="10" name="regtime" value="<%=rs("regtime")%>"></td>
<td width="300" colspan="2">&nbsp;经 验 值：<input size="10" name="experience" value="<%=rs("experience")%>"></td>
</tr>

<tr class=a3>
<td colspan="4">&nbsp;用户头像：<input size="50" name="userface" value="<%=rs("userface")%>"></td>
</tr>

<tr class=a3>
<td colspan="4">&nbsp;用户照片：<input size="50" name="userphoto" value="<%=rs("userphoto")%>"></td>
</tr>

<tr class=a1 id=TableTitleLink>
<td colspan="4" align="center" height="25">
生活资料</td>
</tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;真实姓名：<input type="text" name="realname" size="10" value="<%=realname%>"><td width="50%" height="3" colspan="2">
&nbsp;性　　别：<input type="text" name="sex" size="10" value="<%=rs("sex")%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;出生日期：<input type="text" name="birthday" size="10" value="<%=rs("birthday")%>"><td width="50%" height="3" colspan="2">
&nbsp;国　　家：<b><input type="text" name="country" size="10" value="<%=country%>"></b></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;省　　份：<input type="text" name="province" size="10" value="<%=province%>"><td width="50%" height="3" colspan="2">
&nbsp;城　　市：<input type="text" name="city" size="10" value="<%=city%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;邮政编号：<input type="text" name="postcode" size="10" value="<%=postcode%>"><td width="50%" height="3" colspan="2">
&nbsp;血　　型：<input maxlength=4 size=10 name=blood value="<%=blood%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;信　　仰：<input type="text" name="belief" size="10" value="<%=belief%>"><td width="50%" colspan="2">
&nbsp;职　　业：<input type="text" name="occupation" size="10" value="<%=occupation%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;婚姻状况：<input maxlength=4 size=10 name=marital value="<%=marital%>"><td width="50%" colspan="2">
&nbsp;最高学历：<input type="text" name="education" size="10" value="<%=education%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;毕业院校：<input type="text" name="college" size="20" value="<%=college%>"><td width="50%" colspan="2">
&nbsp;电话号码：<input type="text" name="phone" size="10" value="<%=phone%>"></tr>
<tr class=a3 id=TableTitleLink>
<td width="50%" colspan="2">
&nbsp;联系地址：<input type="text" name="address" size="20" value="<%=address%>"><td width="50%" colspan="2">
&nbsp;手机号码：<input type="text" name="UserMobile" size="10" value="<%=rs("UserMobile")%>"></tr>

<tr class=a4>
<td width="600" colspan="4">&nbsp;用户签名：<textarea name="sign" rows="4" cols="50"><%=rs("sign")%></textarea></td>
</tr>

<tr class=a1 id=TableTitleLink>
<td width="50%" align="center" height="13">
<a onclick="checkclick('您确定要删除该用户所有发表过的帖子?');" href="?menu=userdeltopic&username=<%=rs("username")%>">
删除该用户的所有帖子</a>


<td width="201" colspan="2" align="center" height="13">
<input type="submit" value=" 更 新 " name="Submit"></td>
<td width="50%" align="center" height="13">

<a onclick="checkclick('您确定要删除该用户的所有资料?');" href="?menu=userdel&username=<%=rs("username")%>">
删除该用户的所有资料</a></td>

</td>
</tr>


</table>
</form><A href="javascript:history.back()">返 回</A>
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
已经将 <%=username%> 所有发表过的帖子全部删除<br><br><a href=javascript:history.back()>返 回</a>
<%
end sub

sub userdel

if username=Request.Cookies("username") then error2("操作错误！")

for each ho in Request("username")
ho=HTMLEncode(ho)
conn.execute("delete from [user] where username='"&ho&"'")
conn.execute("delete from [reforum"&ReForumName&"] where username='"&ho&"'")
next

%>
已经成功删除 <%=username%> 的所有资料<br><br><a href=javascript:history.back()>返 回</a>
<%
end sub


sub userok

if Request("userlife")>100 then error2("体力不能大于100！")

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
%> 更新成功</b></font></td>
</tr></table><br><br><a href=javascript:history.back()>返 回</a>
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
		<td align="center">激活</td>
		<td align="center">用户名</td>
		<td align="center">Email</td>
		<td align="center">注册时间</td></tr>
<%



sql="select * from [user] where membercode=0 order by id Desc"
rs.Open sql,Conn,1

pagesetup=20 '设定每页的显示数量
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '总页数
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '跳转到指定页数
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
<input type=submit  onclick="checkclick('您确定要激活所选的用户?');" value=" 激 活 "></TD>

      <TD width="71%" class=a3 colspan="2">
		
		
&nbsp;<b>共有
<font color="990000"><%=TotalPage%></font> 页 [
<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"menu=Activation")
</script>]</b>
		
</TD>
    </TR>
    


</table><%
end sub

htmlend

%>