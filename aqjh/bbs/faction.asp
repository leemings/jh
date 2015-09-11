<!-- #include file="setup.asp" -->
<%
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")

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
if faction<>"" then error2("您已经加入 "&faction&" 了，不能再加入其他帮派！")
factionname=Conn.Execute("Select factionname From faction where id="&id&"")(0)
if conn.execute("Select count(id) from [user] where faction='"&factionname&"'")(0)>99 then error2("该帮派已经超过100名会员，无法再加入新会员")
conn.execute("delete from [message] where id="&int(Request("messageid"))&" and incept='"&Request.Cookies("username")&"'")
conn.execute("update [user] set faction='"&factionname&"' where username='"&Request.Cookies("username")&"'")
error2("加入帮派成功")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="factionout" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")
if faction=empty then error("<li>您目前没有加入任何帮派！")
If not conn.Execute("Select id From [faction] where buildman='"&Request.Cookies("username")&"'").eof Then error("<li>要退出请先解散帮派")
conn.execute("update [user] set faction='',honor='' where username='"&Request.Cookies("username")&"'")
message=message&"<li>退出帮派成功<li><a href=faction.asp>返回社区帮派</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="invite" then
if username="" then error("<li>请填写受邀人的名称")
if Request.Cookies("username")=username then error("<li>不能自己邀请自己")
if Conn.Execute("Select buildman From [faction] where id="&id&"")(0)<>Request.Cookies("username") then error("<li>只有帮主才有权限执行该操作")
if conn.execute("Select count(id) from [user] where faction='"&faction&"'")(0)>99 then error2("帮派已经超过100名会员，无法再加入新会员")
if Conn.Execute("Select faction From [user] where username='"&username&"'")(0)<>"" then error("<li>对方已经加入了其他帮派")
messageid=conn.execute("select Max(ID)+1 From Message")(0)
conn.Execute("insert into [message] (author,incept,content) values ('"&Request.Cookies("username")&"','"&username&"','<form name=factionAdd"&messageid&" method=post action=faction.asp?id="&id&"&messageid="&messageid&"><input type=hidden name=menu value=factionadd></form><font color=0000FF>【系统消息】："&Request.Cookies("username")&" 邀请您加入 "&faction&" 帮派<br><br><center><a href=javascript:factionAdd"&messageid&".submit()>同意</a>　　　<a href=message.asp?menu=del&id="&messageid&">拒绝</a></font>')")
conn.execute("update [user] set newmessage=newmessage+1 where username='"&username&"'")
message=message&"<li>邀请已经成功发出<li><a href=faction.asp>返回社区帮派</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="factiondel" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")
if Conn.Execute("Select buildman From [faction] where id="&id&"")(0)<>Request.Cookies("username")then error("<li>只有帮主才有权限执行该操作")
conn.execute("update [user] set faction='',honor='' where faction='"&faction&"'")
conn.execute("delete from [faction] where id="&id&"")
message=message&"<li>解散帮派成功<li><a href=faction.asp>返回社区帮派</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="factionUserOut" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")
if Request.Cookies("username")=username then error("<li>不能自己开除自己")
if Conn.Execute("Select buildman From [faction] where id="&id&"")(0)<>Request.Cookies("username")then error("<li>只有帮主才有权限执行该操作")
conn.execute("update [user] set faction='',honor='' where username='"&username&"' and faction='"&faction&"'")
message=message&"<li>已经将 "&username&" 从帮派中开除了<li><a href=faction.asp>返回社区帮派</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="factionUserHonor" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")
if Len(Honor)>7 then error("<li>头衔长度不能超多7个字符")
if Conn.Execute("Select buildman From [faction] where id="&id&"")(0)<>Request.Cookies("username")then error("<li>只有帮主才有权限执行该操作")
conn.execute("update [user] set honor='"&Honor&"' where username='"&username&"' and faction='"&faction&"'")
message=message&"<li> "&username&" 已经获得 "&Honor&" 的头衔<li><a href=faction.asp>返回社区帮派</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="addok" then
factionname=HTMLEncode(Request.Form("factionname"))
allname=HTMLEncode(Request.Form("allname"))
tenet=HTMLEncode(Request.Form("tenet"))
if faction<>empty then message=message&"<li>您已经加入了其他帮派！"
if experience< 10000 then message=message&"<li>您的经验值小于 10000 ！"
if money< 10000 then message=message&"<li>您的金币少于 10000 ！"
if factionname="" then message=message&"<li>帮派简称没有填写"
if Len(factionname)>7 then message=message&"<li>帮派简称最多7个字符"
if allname="" then message=message&"<li>帮派全称没有填写"
If not conn.Execute("Select id From [faction] where factionname='"&factionname&"' or buildman='"&Request.Cookies("username")&"'").eof Then  message=message&"<li>社区中已存在同名帮派<li>您已经建立过帮派"
if message<>"" then error(""&message&"")
conn.Execute("insert into faction (factionname,allname,tenet,buildman) values ('"&factionname&"','"&allname&"','"&tenet&"','"&Request.Cookies("username")&"')")
conn.execute("update [user] set faction='"&factionname&"',[money]=[money]-10000 where username='"&Request.Cookies("username")&"'")
message=message&"<li>创建帮派成功<li><a href=faction.asp>返回社区帮派</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="xiuok" then
allname=HTMLEncode(Request.Form("allname"))
tenet=HTMLEncode(Request.Form("tenet"))
if allname="" then message=message&"<li>帮派全称没有填写"
if message<>"" then error(""&message&"")
if Conn.Execute("Select buildman From [faction] where id="&id&"")(0)<>Request.Cookies("username")then error("<li>只有帮主才有权限执行该操作")
conn.execute("update [faction] set allname='"&allname&"',tenet='"&tenet&"' where id="&id&"")
message=message&"<li>修改帮派成功<li><a href=faction.asp>返回社区帮派</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=faction.asp>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
elseif Request("menu")="look" then
sql="select * from faction where id="&id&""
Set Rs=Conn.Execute(sql)
%>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 
<a href="faction.asp">社区帮派</a></td>
</tr>
</table><br>
<table width="82%" border="0" align="center" cellspacing="1" cellpadding="2"  class=a2 height="150">
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>帮派简称:</b></font></div>
</td>
<td width="82%"><%=rs("factionname")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>帮派全称:</b></font></div>
</td>
<td width="82%"><%=rs("allname")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>帮派公告:</b></font></div>
</td>
<td width="82%"><%=rs("tenet")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>创建时间:</b></font></div>
</td>
<td width="82%"><%=rs("addtime")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>帮主名称:</b></font></div>
</td>
<td width="82%"><%=rs("buildman")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>现有会员:</b></font></div>
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
%><%=i%>人</td>
</tr>

<tr bgcolor="FFFFFF">
<td width="15%">
<div align="center"><font color="000066"><b>会员名单:</b></font></div>
</td>
<td width="82%">
<%=list%>
</td>
</tr>

</table>
<br><center><INPUT onclick=history.back(-1) type=button value=" << 返 回 ">
<%

htmlend
end if



%>
<center>

<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 
<a href="faction.asp">社区帮派</a></td>
</tr>
</table><br>
<%
if Request("menu")="add" then

%>
<form method=post name=form action=faction.asp?menu=addok>

<table cellspacing=1 cellpadding=2 width=442 border=0 align="center" class=a2>
<tr class=a1>
<td width=526 colspan="2" align="center" height="25">
创建门派</td>
</tr>
<tr class=a3>
<td width=187 align="right">
<b><font color="0033CC">帮派简称：</font></b></td>
<td width=339>
<input maxlength=7 name=factionname size="10"> 最多7个字符</td>
</tr>
<tr class=a3>
<td width=187>
<div align="right"><b><font color="0033CC">帮派全称： </font></b></div>
</td>
<td width=339>
<input size=30 name=allname>
</td>
</tr>
<tr class=a3>
<td width=187 height=15>
<div align="right"><b><font color="0033CC">帮派公告： </font></b></div>
</td>
<td width=339 height=15>
<input size=40 name=tenet>
</td>
</tr>
<tr class=a3>
<td width=526 colspan=2 height=8>
<div align=center>
<input type=submit value=" 创 建 " name=Submit23>
<input type=reset value=" 重 填 " name=Submit24>
</div>
</td>
</tr>
<tr class=a3>
<td width=526 colspan=2 height=7>
<ol>
创建门派的注意事项：
<li>您的经验值必须 10000 以上
<li>需要扣除您身上 10000 金币作为门派基金 </li>
<li>帮派最多只能容纳 100 名会员</td>
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
<td colspan="2" height="25" align="center" class=a1>　　门派设定</td>
</tr>


<tr class=a3>
<td>　　帮派简称： </td>
<td>
<%=rs("factionname")%></td>
</tr>
<tr class=a3>
<td>　　帮派全称： </td>
<td><input size="30" name="allname" value="<%=rs("allname")%>"> </td>
</tr>
<tr class=a3>
<td>　　帮派公告： </td>
<td><input size="60" name="tenet" value="<%=rs("tenet")%>"> </td>
</tr>
<tr class=a3>
<td colSpan="2">
<div align="center">
<input type="submit" value=" 修 改 " name="Submit">
<input type="reset" value=" 重 填 " name="Submit2">
</div>
</td>
</tr>
</table>
</form>
<%
else
%>


<p>&nbsp;<a href="?menu=add"><img src="images/plus/niu05.gif" width="26" height="26" border="0"><img src="images/plus/cj.gif" width="95" height="26" border="0"></a>　　　　<a href="?menu=factionout&sessionid=<%=session.sessionid%>"><img src="images/plus/niu05.gif" width="26" height="26" border="0"><img src="images/plus/tc.gif" width="95" height="26" border="0"></a><br>
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
if(usernames==undefined){alert('请选择对象');return;}
if(act=="add"){document.location='friend.asp?menu=add&username='+usernames+'';}
if(act=="post"){open('friend.asp?menu=post&incept='+usernames+'','','width=320,height=170')}
if(act=="look"){open('Profile.asp?username='+usernames+'')}
if(act=="out"){document.location='faction.asp?menu=factionUserOut&id=<%=rs("id")%>&sessionid=<%=session.sessionid%>&username='+usernames+'';}
if(act=="Honor"){var id=prompt("请输入该会员的头衔！","");if(id){document.location='faction.asp?menu=factionUserHonor&id=<%=rs("id")%>&sessionid=<%=session.sessionid%>&username='+usernames+'&Honor='+id+'';}}
}

function invite(){
var id=prompt("请输入邀请加入本帮派的会员名称！","");if(id){document.location='faction.asp?menu=invite&id=<%=rs("id")%>&sessionid=<%=session.sessionid%>&username='+id+'';}
}

</SCRIPT>

<tr bgcolor="FFFFFF">
<td width="10%">
<div align="center"><font color="000066"><b>帮派简称</b></font><font color="#000066"><b>:</b></font></div>
</td>
<td width="28%"><%=rs("factionname")%></td>
<td width="10%" align="center"><font color="000066"><b>帮派全称</b></font><font color="#000066"><b>:</b></font></td>
<td width="27%"><%=rs("allname")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="10%">
<div align="center"><font color="000066"><b>创建人</b></font><font color="#000066"><b>:</b></font></div>
</td>
<td width="28%"><%=rs("buildman")%></td>
<td width="10%" align="center"><font color="000066"><b>创建时间</b></font><font color="#000066"><b>:</b></font></td>
<td width="27%"><%=rs("addtime")%></td>
</tr>
<tr bgcolor="FFFFFF">
<td width="10%">
<div align="center"><font color="000066"><b>帮派公告</b></font><font color="#000066"><b>:</b></font></div>
</td>
<td width="82%" colspan="3"><%=rs("tenet")%></td>
</tr>
<form method=post name=faction>
<tr bgcolor="FFFFFF">
<td width="10%">
<div align="center"><font color="000066"><b>会员名单</b></font><font color="#000066"><b>:</b></font><br>
	<font size="1">共 <font color=red><%=i%></font> 人</font></div>
</td>
<td width="65%" colspan="3">

<%=list%>
</td>
</tr>

<tr>
<td width="38%" colspan="2" align="center" class=a1>
会员操作</td>
<td width="37%" colspan="2" align="center" class=a1>
帮主管理</td>
</tr>
<tr bgcolor="FFFFFF">
<td width="38%" colspan="2" align="center">
<a onclick="factionUser(document.faction.username,'add')" href=#>添加好友</a>　　<a onclick="factionUser(document.faction.username,'post')" href=#>发送信息</a>　　<a onclick="factionUser(document.faction.username,'look')" href=#>查看资料</a></td>
<td width="37%" colspan="2" align="center">


<a onclick="invite()" href=#>添加会员</a>　　<a onclick="factionUser(document.faction.username,'out')" href=#>开除会员</a>　　<a onclick="factionUser(document.faction.username,'Honor')" href=#>赐予头衔</a>　　<a href="?menu=xiu&id=<%=rs("id")%>">修改资料</a>　　<a onclick=checkclick('您确定要解散该帮派？') href="?menu=factiondel&sessionid=<%=session.sessionid%>&id=<%=rs("id")%>">解散此帮</a></td>
</tr>
</form>
<%
RS.Close


else
%>
<tr bgcolor="FFFFFF"><td align="center">没有创建或者加入任何帮派</td></tr>
<%
end if
%>
</table>




<SCRIPT>valignbottom()</SCRIPT>

<%
end if


htmlend
%>