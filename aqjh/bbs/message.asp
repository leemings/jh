<!-- #include file="setup.asp" -->
<%
if Request.Cookies("username")=empty then error2("您还未登录论坛")

id=int(Request("id"))
username=HTMLEncode(Trim(Request("username")))



if Request("menu")="post" then
%>
<title>发送讯息 - Powered By BBSxp</title>
<SCRIPT>
function check(theForm) {

if(theForm.incept.value == "" ) {
alert("昵称不能没有填写！");
return false;
}

if(theForm.incept.value == "<%=Request.Cookies("username")%>" ) {
alert("请输入您要发送的对象，不能发讯息给自己！");
return false;
}

if(theForm.content.value == "" ) {
alert("不能发空讯息！");
return false;
}

if (theForm.content.value.length > 255){
alert("对不起，您的讯息不能超过 255 个字节！");
return false;
}
}

function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.form.submit();}}


function DoTitle(addTitle) {
document.form.incept.value=addTitle
}

function Check(){var Name=document.form.incept.value;
if(Name == ""){alert("昵称不能为空！");return false;}
window.open("friend.asp?menu=post&log=1&incept="+Name+"&content="+document.form.content.value+"","_blank","width=320,height=270");window.close()
}

</SCRIPT><body topmargin=0>
<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0><TR><form name=form action="friend.asp" method="post">
<input type="hidden" name="menu" value="addpost">
<TD height="35">
&nbsp;昵称：<input name="incept" type="text" size="10"></TD>

<TD align=right height="35">

<select onchange=DoTitle(this.options[this.selectedIndex].value)>


<%
if Request("report")="1" then
%>
<option>版主列表</option>
<SCRIPT>
var moderated="<%=Request("moderated")%>"
var list= moderated.split ('|'); 
for(i=0;i<list.length;i++) {
if (list[i] !=""){document.write("<option value="+list[i]+">"+list[i]+"</option>")}
}
</SCRIPT>
<%
else

%>
<option>好友列表</option>
<SCRIPT>
var moderated="<%=Conn.Execute("Select friend From [user] where username='"&Request.Cookies("username")&"'")(0)%>"
var list= moderated.split ('|'); 
for(i=0;i<list.length;i++) {
if (list[i] !=""){document.write("<option value="+list[i]+">"+list[i]+"</option>")}
}
</SCRIPT>
<%
end if

%>



</select>
</TD></TR><TR><TD VALIGN=top ALIGN=right colspan="3" bgcolor="F8F8F8">
    <textarea name="content" cols="39" rows="6" onkeydown=presskey()><%=Request("body")%></textarea>
</TD></TR></TABLE><TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0 height="30">
<tr ALIGN=center><TD>
<input onclick=javascript:Check() type="button" value="聊天记录">
</td><TD><input type="reset" value="取消发送" OnClick="window.close();"> </td><TD><input name=submit1 type="submit" value="发送讯息" onclick="return check(this.form)"></td>
</TR></form>
</TABLE>
<%
responseend

end if


if Request("menu")="del" then
conn.execute("delete from [message] where id="&id&" and incept='"&Request.Cookies("username")&"'")
error2("删除成功")
elseif Request("menu")="alldel" then
conn.execute("delete from [message] where incept='"&Request.Cookies("username")&"'")
error2("已经成功清空收件箱")
end if




top

%>



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

<center>

<SCRIPT>valigntop()</SCRIPT>
<TABLE  cellSpacing=1 border=0 width="97%" class=a2 cellpadding="3">
<TR><TD align="center" bgcolor="FFFFFF" colspan=5>

          <a style="text-decoration: none; color: #000000" href="message.asp">
          <img alt src="images/m_inbox.gif" border="0" dypop="收件箱"></a>&nbsp;
          <a style="text-decoration: none; color: #000000" href="message.asp?send=1">
          <img alt src="images/M_issend.gif" border="0" dypop="已发送邮件"></a>&nbsp;
          <a href="message.asp?menu=alldel" onclick=checkclick('您确定要清空收件箱?')>
          <img alt src="images/recycle.gif" border="0" dypop="清空收件箱"></a>&nbsp;
          <a href="friend.asp">
          <img alt src="images/M_address.gif" border="0" dypop="发送讯息"></a>&nbsp;
          <a style=cursor:hand onclick="javascript:open('?menu=post','','width=320,height=170')"><img alt src="images/m_write.gif" border="0" dypop="发送讯息"></a></td></TR>

  <TR>
          
    <TD vAlign=center align=middle 
          width=5% class=a1 height="20">　</TD>
          
    <TD vAlign=center align=middle 
          width=12% class=a1 height="20"><B><span id=send>发件人</span></B></TD>
          
    <TD vAlign=center align=middle 
          width=50% class=a1 height="20"><B>内容</B></TD>
          
    <TD align=middle 
          width=17% class=a1 height="20"><B>日期</B></TD>
          
    <TD vAlign=center align=middle 
          width=16% class=a1 height="20"><B>操作</B></TD>
  </TR>



<%



if Request("send")="1" then
sql="select * from message where author='"&Request.Cookies("username")&"' order by time Desc"
response.write "<script>send.innerText='收件人'</script>"
else
sql="select * from message where incept='"&Request.Cookies("username")&"' order by time Desc"

end if

rs.Open sql,Conn,1

pagesetup=10 '设定每页的显示数量
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '总页数
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '跳转到指定页数

i=0
Do While Not RS.EOF and i<pagesetup
i=i+1


if Request("send")="1" then
author=rs("incept")
del="<a style=cursor:hand onclick=javascript:open('friend.asp?menu=post&log=1&incept="&author&"','','width=320,height=270')>聊天记录</a>"

else
author=rs("author")
del="<a style=cursor:hand onclick=javascript:open('friend.asp?menu=post&incept="&rs("author")&"','','width=320,height=170')><img alt='回复这个帖子' src=images/replynow.gif border=0></a> <a onclick=checkclick('您确定要删除此条讯息?') href=?menu=del&id="&rs("id")&"><img alt='删除该留言' src=images/del.gif border=0></a>"
end if


%>



        <TR>
          
    <TD vAlign=center align=middle width="5%" bgcolor="FFFFFF"> <IMG 
            src="images/f_norm.gif"></TD>
          
    <TD vAlign=center align=middle width="12%" bgcolor="FFFFFF"> <A href="Profile.asp?username=<%=author%>"><%=author%></A> 
      　</TD>
          
    <TD width="50%" bgcolor="FFFFFF">
<script>document.write(ybbcode("<%=rs("content")%>"))</script>
</TD>
    <TD align="center" width="17%" bgcolor="FFFFFF"><%=rs("time")%></FONT>　</TD>
          
    <TD align="center" width="16%" bgcolor="FFFFFF"> <%=del%> 　</TD>
  </TR>


<%
RS.MoveNext
loop
RS.Close
if newmessage>0 then conn.execute("update [user] set newmessage=0 where username='"&Request.Cookies("username")&"'")
%>
</TD></TR></TABLE>
<SCRIPT>valignbottom()</SCRIPT>
<BR>

Page：[

<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"send=<%=Request("send")%>")
</script>

]


</TABLE><br>

<%htmlend%>