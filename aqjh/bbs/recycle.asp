<!-- #include file="setup.asp" -->
<%
top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")

if membercode < 4 then error2("您的权限不够，无法进入！")


select case Request("menu")

case "deoxidize"
for each ho in request.form("id")
ho=int(ho)
conn.execute("update [forum] set deltopic=0,lasttime='"&now()&"',lastname='"&Request.Cookies("username")&"' where id="&ho&" and deltopic=1")
next


case "clean"

if ""&Request("day")&"" < "2" and membercode < 5 then error2("只有社区区长才有权限清空 1 天前的帖子")

conn.execute("delete from [forum] where deltopic=1 and lasttime<"&SqlNowString&"-"&Request("day")&"")
log("清空回收站内 "&Request("day")&" 天以前的文章！")
error2("已经清空回收站内 "&Request("day")&" 天以前的文章！")
end select



sql="select * from forum where deltopic=1 order by lasttime Desc"
rs.Open sql,Conn,1

pagesetup=20 '设定每页的显示数量
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '总页数
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '跳转到指定页数


%>


<script>
function CheckAll(form){for (var i=0;i<form.elements.length;i++){var e = form.elements[i];if (e.name != 'chkall')e.checked = form.chkall.checked;}}
</script>

<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <a href="recycle.asp">回收站</a></td>
</tr>
</table><br>
<SCRIPT>valigntop()</SCRIPT>
<TABLE cellSpacing=1 cellPadding=1 width=97% align=center border=0 class=a2>
  <form method="POST" action=recycle.asp?menu=deoxidize>
<TR height=25 class=a1>
      <TD align=middle width=3%>　</TD>
      <TD align=middle width=3%>
<input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON"></TD>
      <TD width="46%" height=24 align=middle>主题</TD>
      <TD align=middle width=9% height=24>作者</TD>
      <TD align=middle width=5% height=24>回复</TD>
      <TD align=middle width=4% height=24>点击</TD>
      <TD width=30% height=24>　　 删除时间　　　　 删除者</TD>
    </TR>


<%

i=0
Do While Not RS.EOF and i<pagesetup
i=i+1




if rs("toptopic")=2 then
reimage="<img src=images/top.gif border=0>"
elseif rs("toptopic")=1 then
reimage="<img src=images/f_top.gif border=0>"
elseif rs("locktopic")=1 then
reimage="<img src=images/f_locked.gif border=0>"
elseif rs("polltopic")<>"" then
reimage="<img src=images/f_poll.gif border=0>"
elseif rs("replies")>15 or rs("views") > 150 then
reimage="<img src=images/f_hot.gif border=0>"
elseif rs("replies")>0 then
reimage="<img src=images/f_new.gif border=0>"
else
reimage="<img src=images/f_norm.gif border=0>"
end if


if rs("replies")=0 then
replies="-"
else
replies=rs("replies")
end if

%>

<TR height=25>
      <TD align=middle width=3% class=a4 height=24><a target=_blank href=ShowPost.asp?id=<%=rs("id")%>><%=reimage%></a></TD>
      <TD align=middle width=3% class=a3>
<INPUT type=checkbox value=<%=rs("id")%> name=id></TD>
      <TD width="46%" align=left class=a4>&nbsp;<%if rs("icon")>0 then%><img border=0 src=images/brow/<%=rs("icon")%>.gif> <%end if%><a href=ShowPost.asp?id=<%=rs("id")%>><%=rs("topic")%></a></TD>
      <TD align=middle width=9% class=a3><a href=Profile.asp?username=<%=rs("username")%>><%=rs("username")%>　</TD>
      <TD align=middle width=5% class=a4><%=replies%>　</TD>
      <TD align=middle width=4% class=a3><%=rs("views")%>　</TD>
      <TD align=left width=30% class=a4>&nbsp;<%=rs("lasttime")%> | <a href=Profile.asp?username=<%=rs("lastname")%>><%=rs("lastname")%></a></TD>
    </TR>
    

<%
RS.MoveNext
loop
RS.Close
%>







</TABLE>
<SCRIPT>valignbottom()</SCRIPT>
<TABLE cellSpacing=0 cellPadding=1 width=97% align=center border=0>
  
<TR height=25>
      <TD width="100%" height=2> 
        <TABLE cellSpacing=0 cellPadding=3 width="100%" border=0>

<TR>
<TD height=2 valign="top">
<input type=submit onclick="checkclick('您确定要还原所选的主题?');" value="还 原">
<b>共有
<font color="990000"><%=TotalPage%></font> 页 [
<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"")
</script>]</b></TD>
<TD align=right height=2>
</form>
 <form method="POST" action=?menu=clean>

清空 <input maxLength="1" size="1" value="7" name=day> 天以前的文章 <INPUT  onclick="checkclick('您确定要清空回收站?');"  type=submit value="确定"> 

 </form>




</TD></TR></TABLE></TD></TR></TABLE>




<br>



<center>


<table cellSpacing="4" cellPadding="0" width="80%" border="0">
<tr>
<td noWrap width="200">
<img alt src="images/f_new.gif" border="0">&nbsp;打开主题 (有回复的主题)</td>
<td noWrap width="100">
<img alt src="images/f_hot.gif" border="0">&nbsp;热门主题 </td>
<td noWrap width="100">
<img alt src="images/f_locked.gif" border="0">&nbsp;关闭主题</td>
<td noWrap width="150">
<img src="images/topicgood.gif">
精华主题</td>
</tr>
<tr>
<td noWrap width="200">
<img alt src="images/f_norm.gif" border="0">&nbsp;打开主题 (没有回复的主题)</td>
<td noWrap width="100">
<img alt src="images/f_poll.gif" border="0">&nbsp;投票主题</td>
<td noWrap width="100">
<img alt src="images/f_top.gif" border="0">&nbsp;置顶主题</td>
<td noWrap width="150">
<img src="images/my.gif">
自己发表的主题</td>
</tr>
</table>


</center>
</div>




<%


htmlend
%>