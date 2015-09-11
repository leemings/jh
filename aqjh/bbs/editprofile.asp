<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" -->

<%
top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")


select case Request.form("menu")

case "editProfileok"
editProfileok
case "contactok"
contactok
case "optionok"
optionok
case "passok"
passok
end select


%>
<SCRIPT src="inc/birthday.js"></SCRIPT>

<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 控制面板</td>
</tr>
</table><br>

<table cellspacing=1 cellpadding=1 width=97% align=center border=0 class=a2>
  <TR id=TableTitleLink class=a1 height="25">
      <Td align="center"><b><a href="usercp.asp">控制面板首页</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp">基本资料修改</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp?menu=contact">联系资料修改</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp?menu=pass">用户密码修改</a></b></td>
      <TD align="center"><b><a href="EditProfile.asp?menu=option">编辑论坛选项</a></b></td>
      <TD align="center"><b><a href="message.asp">用户短信服务</a></b></td>
      <TD align="center"><b><a href="friend.asp">编辑好友列表</a></b></td>
      </TR></TABLE>
<br>
<%


select case Request("menu")
case ""
index
case "contact"
contactPage
case "option"
OptionPage
case "pass"
pass
end select




sub index
sql="select * from [user] where username='"&Request.Cookies("username")&"'"
Set Rs=Conn.Execute(sql)

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
<script>function New(para_URL){var URL=new String(para_URL);window.open(URL,'','resizable,scrollbars')}</script>


<SCRIPT>valigntop()</SCRIPT>

<table width=97% cellspacing=1 cellpadding=4 align=center border=0 class=a2>
  <form method="POST" name="form" action="EditProfile.asp">
<input type=hidden name=menu value="editProfileok">
<tr>
<td height="20" align="center" colspan="2" valign="bottom" class=a1>
<b>&nbsp;:::基本资料:::</b></td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="34%"> 　<b>真实姓名：</b> 
      <input type="text" name="realname" size="17" value="<%=realname%>">
</td>
    <td height="71" align="left" valign="top" class=a4 rowspan="12" width="64%"> 
      <table width="486" cellspacing="0" cellpadding="5">
<tr>
<td width="476"><b>头<font color="000000">　</font>像：</b>
<img src="<%=rs("userface")%>" name="tus" width=32 height=32>　<script>function showimage(){document.images.tus.src=""+document.form.userface.options[document.form.userface.selectedIndex].value+"";}</script><select name=userface size=1 onChange="showimage()">
<option value="<%=rs("userface")%>">默认</option>
<script>
for(i=1;i<=99;i++) {
document.write("<option value='images/face/"+i+".gif'>"+i+"</option>")
}
</script>
</select>　<a href="JavaScript:New('register.asp?menu=face')">查看所有的头像列表</a></td>
</tr>
<tr>
<td width="466"><b>生<font color="000000">　</font>日：</b> <input onfocus="show_cele_date(birthday,'','',birthday)" value="<%=rs("birthday")%>" name="birthday"></td>
</tr>
<tr>
<td width="476"><b><font color="000000">性　格</font>： &nbsp; </b>

<textarea name=character rows=5 cols=65><%=character%></textarea><font color="000000">　 </font>
</td>
</tr>
<tr>
<td width="476"><b>个人简介： &nbsp;
<textarea name=personal rows=5 cols=65><%=personal%></textarea>
</b></td>
</tr>
<tr>
<td height="10" width="476"><b>签名档： &nbsp;
<textarea name=sign rows=5 cols=65><%=rs("sign")%></textarea> </b> </td>
</tr>
</table>
</td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="34%"><b> 　性　　别：</b> 
      <select size=1 name=sex>
<option value="" selected>[请选择]</option>
<option value=male <%if rs("sex")="male" then%>selected<%end if%>>男</option>
<option value=female <%if rs("sex")="female" then%>selected<%end if%>>女</option>
</select>



</td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="34%">　<b>国　　家：</b> <b> 
      <input type="text" name="country" size="17" value="<%=country%>">
</b> </td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="34%">　<b>省　　份：</b> 
      <input type="text" name="province" size="17" value="<%=province%>">
</td>
</tr>
<tr>
    <td height="1" valign="top" class=a4 width="34%">　<b>城　　市： </b> 
      <input type="text" name="city" size="17" value="<%=city%>">
</td>
</tr>
<tr>
    <td height="1" valign="top" class=a3 width="34%">　<b>邮政编码： </b> 
      <input type="text" name="postcode" size="17" value="<%=postcode%>">
</td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="34%">　<b>血　　型：</b> 
      <input maxlength=4 size=4
name=blood value="<%=blood%>">
</td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="34%">　<b>信　　仰：</b> 
      <input type="text" name="belief" size="17" value="<%=belief%>"></td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="34%">　<b>职　　业：</b> 
      <input type="text" name="occupation" size="17" value="<%=occupation%>"></td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="34%">　<b>婚姻状况：</b> 
      <input maxlength=4 size=17 name=marital value="<%=marital%>"></td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="34%">　<b>最高学历：</b> 
      <input type="text" name="education" size="17" value="<%=education%>"></td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="34%">　<b>毕业院校：</b> 
      <input type="text" name="college" size="17" value="<%=college%>"></td>
</tr>
<tr bgcolor="FFFFFF" class=a1>
<td height="25" align="left" colspan="2"><b>&nbsp;:::私人联系资料:::</b></td>
</tr>

<tr bgcolor="FFFFFF" class=a3>
<td valign="middle" align="left">　<b>电话号码：</b><input type="text" name="phone" size="20" value="<%=phone%>"></td>
<td valign="middle" align="left"><b>手机号码：</b><input type="text" name="UserMobile" size="20" value="<%=rs("UserMobile")%>"></td>
</tr>

<tr bgcolor="FFFFFF" class=a4>
<td valign="middle" align="left" colspan="2">　<b>联系地址：</b><input type="text" name="address" size="60" value="<%=address%>"></td>
</tr>

<tr class=a3>
    <td height="1" align="center" width="100%" colspan="2">

<input type="submit" name="Submit1" value=" 确 定 "></td>
</tr>
</table>

<SCRIPT>valignbottom()</SCRIPT>

</form>

<Script>
document.form.sign.value = unybbcode(document.form.sign.value);
document.form.personal.value = unybbcode(document.form.personal.value);
document.form.character.value = unybbcode(document.form.character.value);
function unybbcode(temp) {
temp = temp.replace(/<br>/ig,"\n");
return (temp);
}
</Script><%
end sub


sub contactpage
sql="select * from [user] where username='"&Request.Cookies("username")&"'"
Set Rs=Conn.Execute(sql)



UserIM=split(rs("UserIM"),"\")
qq=UserIM(0)
icq=UserIM(1)
uc=UserIM(2)
aim=UserIM(3)
msn=UserIM(4)
Yahoo=UserIM(5)



%>


<SCRIPT>valigntop()</SCRIPT>

<table width=97% cellspacing=1 cellpadding=4 align=center border=0 class=a2>
  <form method="POST" name="form" action="EditProfile.asp">
<input type=hidden name=menu value="contactok">


<tr bgcolor="FFFFFF" class=a1>
<td colspan="2" height="25" align="center"><b>&nbsp;:::联系信息:::</b></td>
</tr>
<tr bgcolor="FFFFFF" class=a3>
<td height="4" valign="middle" align="left">　<b>QQ号码： 
<input type="text" name="qq" size="20" onkeyup=if(isNaN(this.value))this.value='' value="<%=qq%>"></b></td>
<td height="4" valign="middle" align="left"><b>ICQ：　　　<input type="text" name="icq" size="20" onkeyup=if(isNaN(this.value))this.value='' value="<%=icq%>"></b></td>
</tr>
<tr bgcolor="FFFFFF" class=a4>
<td height="4" valign="middle" align="left">　<b>UC号码： <input type="text" name="uc" size="20" onkeyup=if(isNaN(this.value))this.value='' value="<%=uc%>"></b></td>
<td height="4" valign="middle" align="left"><b>AIM：　　　<input type="text" name="aim" size="20" value="<%=aim%>"></b></td>
</tr>
<tr bgcolor="FFFFFF" class=a3>
<td height="2" valign="middle" align="left">　<b>MSN IM： <input type="text" name="msn" size="20" value="<%=msn%>"></b>　</td>
<td height="2" valign="middle" align="left"><b>Yahoo IM： <input type="text" name="Yahoo" size="20" value="<%=yahoo%>"></b></td>
</tr>

<tr bgcolor="FFFFFF" class=a4>
<td height="1" valign="middle" align="left">　<b>邮件地址：</b><input type="text" name="usermail" size="20" value="<%=rs("usermail")%>"></td>
<td height="1" valign="middle" align="left"><b>个人主页：　</b><input type="text" name="userhome" size="20" value="<%=rs("userhome")%>"></td>
</tr>

<tr class=a3>
    <td height="1" align="center" width="100%" colspan="2">

<input type="submit" name="Submit1" value=" 确 定 "></td>
</tr>
</table>

<SCRIPT>valignbottom()</SCRIPT>

</form>


<%
rs.close

end sub


sub contactok
username=Request.Cookies("username")
if instr(Request("usermail"),"@")=0 then error("<li>您的电子邮件地址填写错误")


sql="select * from [user] where username='"&HTMLEncode(username)&"'"
rs.Open sql,Conn,1,3
rs("usermail")=HTMLEncode(Request("usermail"))
rs("userhome")=HTMLEncode(Request("userhome"))
rs("UserIM")=""&HTMLEncode(Request("qq"))&"\"&HTMLEncode(Request("icq"))&"\"&HTMLEncode(Request("uc"))&"\"&HTMLEncode(Request("aim"))&"\"&HTMLEncode(Request("msn"))&"\"&HTMLEncode(Request("Yahoo"))&""
rs.update
rs.close
message=message&"<li>修改资料成功<li><a href=usercp.asp>控制面板首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=usercp.asp>")
end sub




sub editProfileok
username=Request.Cookies("username")
userface=HTMLEncode(Request("userface"))
sign=HTMLEncode(Request.Form("sign"))
temp=UCase(sign)

if instr(temp,"[/FLASH]")>0 or instr(temp,"[/RM]")>0 or instr(temp,"[/MP]")>0 then message=message&"<li>签名档中不能含有[FLASH] [RM] [MP]代码"

if Len(sign)>255 then message=message&"<li>签名档不能大于 255 个字节"

if instr(userface,";")>0 then message=message&"<li>头像URL中不能含有特殊符号"

if message<>"" then error(""&message&"")

sql="select * from [user] where username='"&HTMLEncode(username)&"'"
rs.Open sql,Conn,1,3

rs("birthday")=HTMLEncode(Request("birthday"))
rs("userface")=userface
rs("sex")=HTMLEncode(Request("sex"))
rs("sign")=sign
rs("landtime")=now()


rs("UserInfo")=""&HTMLEncode(Request("realname"))&"\"&HTMLEncode(Request("country"))&"\"&HTMLEncode(Request("province"))&"\"&HTMLEncode(Request("city"))&"\"&HTMLEncode(Request("postcode"))&"\"&HTMLEncode(Request("blood"))&"\"&HTMLEncode(Request("belief"))&"\"&HTMLEncode(Request("occupation"))&"\"&HTMLEncode(Request("marital"))&"\"&HTMLEncode(Request("education"))&"\"&HTMLEncode(Request("college"))&"\"&HTMLEncode(Request("address"))&"\"&HTMLEncode(Request("phone"))&"\"&HTMLEncode(Request("character"))&"\"&HTMLEncode(Request("personal"))&""
rs("UserMobile")=""&HTMLEncode(Request("UserMobile"))&""


rs.update
rs.close
message=message&"<li>修改资料成功<li><a href=usercp.asp>控制面板首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=usercp.asp>")
end sub





sub OptionPage
sql="select * from [user] where username='"&Request.Cookies("username")&"'"
Set Rs=Conn.Execute(sql)
%>

<SCRIPT>valigntop()</SCRIPT>
<table width=97% cellspacing=1 cellpadding=4 border=0 class=a2 align=center>
  <form method="POST" name="form"action="EditProfile.asp">
<input type=hidden name="menu" value="optionok">
<tr>
<td height="20" align="center" colspan="2" valign="bottom" class=a1>
<b>编辑论坛选项</b></td>
</tr>
<tr>
    <td height="2" align="right" width="45%" class=a4><b>YBB代码：</b></td>
    <td height="2" align="left" width="55%" class=a4> &nbsp; 
      
<input type=radio name="ybbcode" value="0" <%if Request.Cookies("ybbcode")="0" then%>checked<%end if%> id=ybbcode><label for=ybbcode>关闭</label>
<input type=radio name="ybbcode" value="1" <%if Request.Cookies("ybbcode")="1" then%>checked<%end if%> id=ybbcode_1><label for=ybbcode_1>打开</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[打开]</font>
      
</td>
</tr>
<tr>
    <td height="2" align="right" width="45%" class=a3><b>[IMG]代码：</b></td>
    <td height="2" align="left" width="55%" class=a3> &nbsp; 
<input type=radio name="ybbimg" value="0" <%if Request.Cookies("ybbimg")="0" then%>checked<%end if%> id=ybbimg><label for=ybbimg>关闭</label>
<input type=radio name="ybbimg" value="1" <%if Request.Cookies("ybbimg")="1" then%>checked<%end if%> id=ybbimg_1><label for=ybbimg_1>打开</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[打开]</font>
</td>
</tr>
<tr>
    <td class=a4 height="2" align="right" valign="middle" width="45%"><b>
    [FLASH]代码：</b></td>
    <td class=a4 height="2" align="left" valign="middle" width="55%"> &nbsp; 
<input type=radio name="ybbflash" value="0" <%if Request.Cookies("ybbflash")="0" then%>checked<%end if%> id=ybbflash><label for=ybbflash>关闭</label>
<input type=radio name="ybbflash" value="1" <%if Request.Cookies("ybbflash")="1" then%>checked<%end if%> id=ybbflash_1><label for=ybbflash_1>打开</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[打开]</font>
</td>
</tr>

<tr class=a3>
    <td height="2" align="right" width="45%"><b>表情代码：</b></td>
    <td height="2" align="left" width="55%"> &nbsp; 
<input type=radio name="ybbbrow" value="0" <%if Request.Cookies("ybbbrow")="0" then%>checked<%end if%> id=ybbbrow><label for=ybbbrow>关闭</label>
<input type=radio name="ybbbrow" value="1" <%if Request.Cookies("ybbbrow")="1" then%>checked<%end if%> id=ybbbrow_1><label for=ybbbrow_1>打开</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[打开]</font>
</td>
</tr>


<tr>
<td height="20" align="center" colspan="2" valign="bottom" class=a1>
<b>编辑帖子选项</b></td>
</tr>


<tr class=a4>
    <td height="2" align="right" width="45%"><b>在帖子中显示用户头像：</b></td>
    <td height="2" align="left" width="55%"> &nbsp; 
<input type=radio name="showface" value="0" <%if Request.Cookies("showface")="0" then%>checked<%end if%> id=showface><label for=showface>关闭</label>
<input type=radio name="showface" value="1" <%if Request.Cookies("showface")="1" then%>checked<%end if%> id=showface_1><label for=showface_1>打开</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[打开]</font></td>
</tr>


<tr class=a3>
    <td height="2" align="right" width="45%"><b>在帖子中显示用户签名：</b></td>
    <td height="2" align="left" width="55%"> &nbsp; 
<input type=radio name="sign" value="0" <%if Request.Cookies("sign")="0" then%>checked<%end if%> id=sign><label for=sign>关闭</label>
<input type=radio name="sign" value="1" <%if Request.Cookies("sign")="1" then%>checked<%end if%> id=sign_1><label for=sign_1>打开</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[打开]</font></td>
</tr>

<tr>
<td height="20" align="center" colspan="2" valign="bottom" class=a1>
<b>编辑个性选项</b></td>
</tr>

<tr class=a3>
    <td height="2" align="right" width="45%"><b>每页显示主题数：</b></td>
    <td height="2" align="left" width="55%"> &nbsp; 
<select name="pagesetup" size="1">
<option value="">默认</option>
<option value="10" <%if Request.Cookies("pagesetup")="10" then%>selected<%end if%>>10</option>
<option value="20" <%if Request.Cookies("pagesetup")="20" then%>selected<%end if%>>20</option>
<option value="30" <%if Request.Cookies("pagesetup")="30" then%>selected<%end if%>>30</option>
</select>
    </td>
</tr>
<tr class=a4>
    <td height="2" align="right" width="45%"><b>主题列表中突出关键词：</b><br>设定的关键词将用<font color="FF0000">红色</font>显示&nbsp;&nbsp; </td>
    <td height="2" align="left" width="55%"> &nbsp; 
<input size=20 name=key_word value="<%=Request.Cookies("key_word")%>"></td>
</tr>


<tr class=a3>
    <td height="2" align="right" width="45%"><b>过滤用户名单：</b><br><font color="FF0000">多人请用“|”分隔</font>&nbsp;&nbsp; </td>
    <td height="2" align="left" width="55%"> &nbsp; 
<input size=30 name=badlist value="<%=Request.Cookies("badlist")%>"></td>
</tr>



<tr class=a4>
    <td height="1" align="center" width="100%" colspan="2">
<input type="submit" name="Submit2" value=" 确 定 "></td>
</tr>


</table>
<SCRIPT>valignbottom()</SCRIPT>
</form><%
end sub



sub optionok


Response.Cookies("ybbcode")=Request("ybbcode")
Response.Cookies("ybbimg")=Request("ybbimg")
Response.Cookies("ybbflash")=Request("ybbflash")
Response.Cookies("ybbbrow")=Request("ybbbrow")
Response.Cookies("showface")=Request("showface")
Response.Cookies("sign")=Request("sign")
Response.Cookies("badlist")=Request("badlist")
Response.Cookies("key_word")=Request("key_word")
Response.Cookies("pagesetup")=Request("pagesetup")



Response.Cookies("ybbcode").Expires=date+365
Response.Cookies("ybbimg").Expires=date+365
Response.Cookies("ybbflash").Expires=date+365
Response.Cookies("ybbbrow").Expires=date+365
Response.Cookies("eremite").Expires=date+365
Response.Cookies("newmessage").Expires=date+365
Response.Cookies("showface").Expires=date+365
Response.Cookies("sign").Expires=date+365
Response.Cookies("badlist").Expires=date+365
Response.Cookies("key_word").Expires=date+365
Response.Cookies("pagesetup").Expires=date+365
message=message&"<li>更新设置成功<li><a href=usercp.asp>控制面板首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=usercp.asp>")
end sub



sub pass
sql="select * from [user] where username='"&Request.Cookies("username")&"'"
Set Rs=Conn.Execute(sql)
%>
<SCRIPT>valigntop()</SCRIPT>
<table width=97% cellspacing=1 cellpadding=4 border=0 class=a2 align=center>
  <form method="POST" name="form"action="EditProfile.asp">
<input type=hidden name="menu" value="passok">
<tr>
<td height="20" align="center" colspan="2" valign="bottom" class=a1>
<b>用户密码修改</b></td>
</tr>
<tr class=a4>
    <td height="2" align="right" width="45%"><b> 原密码：</b></td>
    <td height="2" align="left" width="55%"> &nbsp; 
      <input type="password" name="olduserpass" size="40">
</td>
</tr>
<tr class=a3>
    <td height="2" align="right" width="45%"><b> 新密码：</b><br>
      <font color="#FF0000">如不更改密码此处请留空</font></td>
    <td height="2" align="left" width="55%"> &nbsp; 
      <input type="password" name="userpass" size="40" maxlength="16">
</td>
</tr>
<tr class=a4>
    <td height="2" align="right" width="45%"><b>确认新密码：</b><br>
      <font color="#FF0000">请与您的新密码保持一致</font></td>
    <td height="2" align="left" valign="middle" width="55%"> &nbsp; 
      <input type="password" name="userpass2" size="40" maxlength="16"></td>
</tr>
<tr class=a3>
    <td height="1" align="right" width="45%"><b>密码提示问题：</b></td>
    <td height="1" align="left" width="55%"> &nbsp; 
      <input type="text" name="question" size="40" value="<%=rs("question")%>">
</td>
</tr>
<tr class=a4>
    <td height="1" align="right" width="45%"><b>密码提示答案：</b><br>
	<font color="#FF0000">如不更改答案此处请留空</font></td>
    <td height="1" align="left" width="55%" class=a4> &nbsp; 
      <input type="text" name="answer" size="40" value=""></td>
</tr>
<tr class=a3>
    <td align="center" width="100%" colspan="2">
<input type="submit" name="Submit" value=" 确 定 "></td>
</tr>
</table>
<SCRIPT>valignbottom()</SCRIPT>
</form>
<%
end sub

sub passok
username=Request.Cookies("username")
userpass=Trim(Request("userpass"))
olduserpass=Trim(Request("olduserpass"))
userpass2=Trim(Request("userpass2"))
question=HTMLEncode(Request("question"))

sql="select * from [user] where username='"&HTMLEncode(username)&"'"
rs.Open sql,Conn,1,3


if md5(olduserpass)<>rs("userpass") then message=message&"<li>您的原密码错误"

if userpass<>userpass2 then message=message&"<li>您的新密码和确认新密码不同"

if message<>"" then error(""&message&"")

rs("question")=question
if userpass<>empty then rs("userpass")=md5(userpass)
if Request("answer")<>empty then rs("answer")=md5(Request("answer"))

rs("landtime")=now()
rs.update
rs.close
message=message&"<li>修改资料成功<li><a href=usercp.asp>控制面板首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=usercp.asp>")
end sub


htmlend
%>