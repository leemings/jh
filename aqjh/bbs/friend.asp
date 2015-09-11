<!-- #include file="setup.asp" -->


<%
if Request.Cookies("username")=empty then error2("请登录后才能使用此功能！")
incept=HTMLEncode(Request("incept"))
username=HTMLEncode(Request("username"))


select case Request("menu")
case "add"
add
case "bad"
bad
case "del"
del
case "post"
post
case "look"
look
case "loadlog"
loadlog
case "addpost"
addpost
case ""
index
end select



sub add
if username="" then error2("请输入您要添加的好友名字！")

if username=Request.Cookies("username") then error2("不能添加自己！")

If conn.Execute("Select id From [user] where username='"&username&"'" ).eof Then error2("数据库不存在此用户的资料！")



sql="select friend from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3
if instr(rs("friend"),"|"&username&"|")>0 then error2("此好友已经添加！")
rs("friend")=""&rs("friend")&""&username&"|"
rs.update
rs.close
error2("已经成功添加好友!")
end sub

sub bad
if Request.Cookies("username")=username then error2("不能把过滤自己!")
Response.Cookies("badlist")=""&Request.Cookies("badlist")&""&username&"|"
error2("已经成功将 "&username&" 过滤了!")
end sub


sub del
sql="select friend from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3
rs("friend")=replace(rs("friend"),"|"&username&"|","|")
rs.update
rs.close
index
end sub

sub look

conn.execute("update [user] set newmessage=0 where username='"&Request.Cookies("username")&"'")

page=Request("page")
if page<1 then
disabled="disabled=true"
page=0
end if
count=conn.execute("Select count(id)from message where incept='"&Request.Cookies("username")&"'")(0)
sql="select author,content from message where incept='"&Request.Cookies("username")&"' order by time Desc"
Set Rs=Conn.Execute(sql)
if Count-page<2 then disabled2="disabled=true"

if rs.eof then error2("您没有短讯息！")

RS.Move page
%>


<title>查看消息 - Powered By BBSxp</title>
<body topmargin=0>
<style>
.bt {BORDER-RIGHT: 1px; BORDER-TOP: 1px; FONT-SIZE: 9pt; BORDER-LEFT: 1px; BORDER-BOTTOM: 1px;}
</style>
<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0><TR><TD>&nbsp;昵称：<input readOnly type="text" value="<%=rs("author")%>" size="8"> Email：<input  readOnly type="text" value="<%=Conn.Execute("Select usermail From [user] where username='"&rs("author")&"'")(0)%>" size="10">
</TD><TD align=right><a target=_blank href=Profile.asp?username=<%=rs("author")%>><img border="0" src="<%=Conn.Execute("Select userface From [user] where username='"&rs("author")&"'")(0)%>" width="32" height="32" alt=用户详细资料>
</TD></TR><TR><TD VALIGN=top ALIGN=right colspan="2" bgcolor="F8F8F8">
<script>
document.write ('<iframe id="HtmlEditor"  MARGINHEIGHT=1 MARGINWIDTH=1 height="100" width=100% frameborder=1></iframe>')
frames.HtmlEditor.document.write ('<%=rs("content")%>');
frames.HtmlEditor.document.body.style.fontSize="10pt";
frames.HtmlEditor.document.close();
</script>

</TD></TR></TABLE>
<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0 height="30">
<tr ALIGN=center><TD><input type="button" value="回复讯息" onclick=javascript:open('friend.asp?menu=post&incept=<%=rs("author")%>','_top','width=320,height=170')>
</td><TD><input type="button" value="<<上一条" <%=disabled%> onclick=javascript:open('friend.asp?menu=look&page=<%=page-1%>','_top','')> </td><TD><input type="button" value="下一条>>" <%=disabled2%> onclick=javascript:open('friend.asp?menu=look&page=<%=page+1%>','_top','')>
</td>
</TR></TABLE>
<%
end sub

sub loadlog
sql="select * from message where (author='"&Request.Cookies("username")&"' and incept='"&incept&"') or (author='"&incept&"' and incept='"&Request.Cookies("username")&"') order by time Desc"
Set Rs=Conn.Execute(sql)
do while not rs.eof
content=content&"<font color=red>("&rs("time")&")　　"&rs("author")&"</font><br><font color=0000FF>"&rs("content")&"</font><br>"
rs.movenext
loop
rs.close
if content="" then content="(空)"
%>
<script>
parent.log.innerHTML='<iframe id="HtmlEditor"  MARGINHEIGHT=1 MARGINWIDTH=1 height="100" width=100% frameborder=1></iframe>'
parent.frames.HtmlEditor.document.write ('<%=content%>');
parent.frames.HtmlEditor.document.body.style.fontSize="10pt";
parent.frames.HtmlEditor.document.close();
</script>
<%
end sub

sub post

if incept="" then error2("对不起，您没有输入用户名称！")

sql="select username,userface,usermail from [user] where username='"&incept&"'"
Set Rs=Conn.Execute(sql)
if rs.eof then error2("系统不存在该用户的资料")


%>

<body topmargin=0>
<style>
.bt {BORDER-RIGHT: 1px; BORDER-TOP: 1px; FONT-SIZE: 9pt; BORDER-LEFT: 1px; BORDER-BOTTOM: 1px;}
</style><title>发送消息 - Powered By BBSxp</title>

<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0><TR><form name=form action="friend.asp" method="post">
<input type="hidden" name="menu" value="addpost">
<input type="hidden" name="incept" value="<%=rs("username")%>">
<TD>
&nbsp;昵称：<input readOnly type="text" value="<%=rs("username")%>" size="8"> Email：<input  readOnly type="text" value="<%=rs("usermail")%>" size="10">
</TD><TD align=right><a target=_blank href=Profile.asp?username=<%=rs("username")%>><img border="0" src="<%=rs("userface")%>" width="32" height="32" alt=用户详细资料>
</TD></TR><TR><TD VALIGN=top ALIGN=right colspan="2" bgcolor="F8F8F8">
    <textarea name="content" cols="39" rows="6" onkeydown=presskey() onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')"><%=HTMLEncode(Request("content"))%></textarea>
</TD></TR></TABLE><TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0 height="30">
<tr ALIGN=center><TD><input type="button" value="聊天记录" onclick=looklog(); name=LookLogButton>
</td><TD><input type="reset" value="取消发送" OnClick="window.close();"> </td><TD><input type="submit" value="发送讯息" name=okButton onclick="return check(this.form)" id='btnadd' disabled></td>
</TR></form>
</TABLE>

<span id=log></span>
<script>
function looklog(){
resizeTo(330,330)
document.frames["hiddenframe"].location.replace("?menu=loadlog&incept=<%=incept%>");
document.form.LookLogButton.disabled = true;
}
function check(theForm) {
if (theForm.content.value.lengtd > 255){
alert("对不起，您的讯息不能超过 255 个字节！");
return false;
}
}
function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.form.submit();}}
<%if Request("log")=1 then%>looklog()<%end if%>
</script>
<%
rs.close


end sub

sub addpost
if Len(Request.Form("content"))<2 then error2("内容不能小于 2 字符")
if incept=Request.Cookies("username") then error2("不能给自己发送讯息！")
If conn.Execute("Select id From [user] where username='"&incept&"'" ).eof Then error2("系统不存在"&incept&"的资料")

sql="insert into message(author,incept,content) values ('"&Request.Cookies("username")&"','"&incept&"','"&HTMLEncode(Request.Form("content"))&"')"
conn.Execute(SQL)
conn.execute("update [user] set newmessage=newmessage+1 where username='"&incept&"'")
%>
发送成功！<script>close();</script>
<%
end sub


sub index

top


sql="select username from online where username<>''"
Set Rs=Conn.Execute(sql)
Do While Not RS.EOF
onlinelist=""&onlinelist&""&rs("username")&"|"
rs.movenext
loop
rs.close

%>

<SCRIPT>
var onlinelist= "<%=onlinelist%>";

function add(){
var id=prompt("请输入您要添加的好友名称！","");
if(id){
document.location='friend.asp?menu=add&username='+id+'';
}
}
</SCRIPT>


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
<form method="POST">
<input type=hidden name="menu" value="del">
<SCRIPT>valigntop()</SCRIPT>
<table width="97%" cellSpacing=1 cellPadding=3 align=center border=0 class=a2>
  <tr>
    <td class=a1 align="center">
    昵称 </th>
    <td class=a1 align="center">
    邮件 </th>
    <td class=a1 align="center">
    主页 </th>
    <td class=a1 align="center">
    状态 </th>
    <td class=a1 align="center">
    发短信 </th>
    <td class=a1 align="center">
    操作 </th>
  </tr>
<%

on error resume next '找不到好友资料时候忽略错误

sql="select friend,userface from [user] where username='"&Request.Cookies("username")&"'"
Set Rs=Conn.Execute(sql)
master=split(rs("friend"),"|")
for i = 1 to ubound(master)-1

usermail=Conn.Execute("Select usermail From [user] where username='"&master(i)&"'")(0)
userhome=Conn.Execute("Select userhome From [user] where username='"&master(i)&"'")(0)

%>
  <tr>
    <td vAlign=center align=middle bgcolor="FFFFFF"><a href=Profile.asp?username=<%=master(i)%>><%=master(i)%></a>
    <td align=middle bgcolor="FFFFFF"><a href=mailto:<%=usermail%>><%=usermail%></a>
    <td bgcolor="FFFFFF"><a href=<%=userhome%> target=_blank><%=userhome%></a>
    <td align=middle bgcolor="FFFFFF">
    
<Script>
if(onlinelist.indexOf('<%=master(i)%>|') == -1 ){
document.write("<img border=0 src=images/offline.gif alt='离线'><br>");
}else{
document.write("<img border=0 src=images/online.gif alt='在线'><br>");
}
</Script>
    
    </td>
    <td align=middle bgcolor="FFFFFF"><a style=cursor:hand onclick="javascript:open('friend.asp?menu=post&incept=<%=master(i)%>','','width=320,height=170')">发送</a></td>
    <td align=middle bgcolor="FFFFFF"><INPUT type=radio value=<%=master(i)%> name=username></td>
  </tr>
  
  
  

<%
next
rs.close
%>


  
  

  
  <tr>
    <td vAlign="center" align="right" colSpan="6" bgcolor="FFFFFF">
    <input onclick="javascript:add();" type="button" value="添加好友" name="action">&nbsp;<input onclick="checkclick('确定删除选定的好友吗?');" type="submit" value="删除"></td>
  </tr></form>
</table>
<SCRIPT>valignbottom()</SCRIPT>
<%

htmlend
end sub

responseend
%>