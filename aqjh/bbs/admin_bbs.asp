<!-- #include file="setup.asp" -->
<%
if adminpassword<>session("pass") then response.redirect "admin.asp?menu=login"
log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")


id=HTMLEncode(Request("id"))
bbsid=HTMLEncode(Request("bbsid"))
TimeLimit=HTMLEncode(Request("TimeLimit"))
username=HTMLEncode(Request("username"))


response.write "<center>"



select case Request("menu")
case "applymanage"
applymanage

case "activation"
activation

case "bbsmanage"
bbsmanage

case "bbsmanagexiu"
bbsmanagexiu


case "bbsmanagexiuok"
bbsmanagexiuok

case "bbsadd"
bbsadd

case "bbsaddok"
bbsaddok

case "classs"
classs


case "upclubconfig"
upclubconfig

case "upclubconfigok"
upclubconfigok


case "bbsmanagedel"
conn.execute("delete from [bbsconfig] where id="&id&"")
error2("已经将该论坛的所有数据删除了！")


case "deltopicok"
if bbsid<>"" then BbsIdList="and forumid="&bbsid&""
conn.execute("delete from [forum] where posttime<"&SqlNowString&"-"&TimeLimit&" "&BbsIdList&"")
error2("已经将"&TimeLimit&"天以前的主题删除了！")


case "delforumok"
if bbsid<>"" then BbsIdList="and forumid="&bbsid&""
conn.execute("delete from [forum] where lasttime<"&SqlNowString&"-"&TimeLimit&" "&BbsIdList&"")
error2("已经将"&TimeLimit&"天没有更新过的主题删除了！")


case "delusertopicok"
if username="" then error2("您没有输入用户名！")
if bbsid<>"" then BbsIdList="and forumid="&bbsid&""
conn.execute("delete from [forum] where username='"&username&"' "&BbsIdList&"")
error2("已经将"&username&"发表的主题删除了！")



case "delretopicok"
if bbsid<>"" then BbsIdList="and forumid="&bbsid&""
conn.execute("delete from [reforum"&ReForumName&"] where posttime<"&SqlNowString&"-"&TimeLimit&" "&BbsIdList&"")
error2("已经将"&TimeLimit&"天的回帖删除了！")

case "deluserretopicok"
if username="" then error2("您没有输入用户名！")
conn.execute("delete from [reforum"&ReForumName&"] where username='"&username&"'")
error2("已经将"&username&"发表的回帖删除了！")


case "delbbsconfigok"
conn.execute("delete from [bbsconfig] where hide=1 and lasttime<"&SqlNowString&"-"&TimeLimit&"")
error2("已经将"&TimeLimit&"天没有新帖子的论坛删除了！")



case "delliketopicok"

topic=HTMLEncode(Request("topic"))

if topic="" then error2("您没有输入字符！")
if bbsid<>"" then BbsIdList="and forumid="&bbsid&""
conn.execute("delete from [forum] where topic like '%"&topic&"%' "&BbsIdList&" ")
error2("已经将标题里包含有 "&topic&" 的帖子全部删除了！")


case "uniteok"
hbbs=Request("hbbs")
ybbs=Request("ybbs")
if hbbs = ybbs then error2("不能选择相同论坛！")
if username<>"" then usernamelist="and username='"&username&"'"
conn.execute("update [forum] set forumid="&int(hbbs)&" where forumid="&int(ybbs)&" and lasttime<"&SqlNowString&"-"&TimeLimit&" "&usernamelist&"")
error2("移动论坛资料成功！")

case "activationok"
for each ho in request.form("id")
ho=int(ho)
conn.execute("update [forum] set deltopic=0 where id="&ho&"")
next
error2("已经将激活所选帖子！")



case "delapplication"
Application.Contents.RemoveAll()
error2("已经清除服务器上所有的application缓存！")





end select

sub applymanage
%>

<table cellspacing=1 cellpadding=2 width=97% border=0 class=a2 align=center>
<tr class=a1 id=TableTitleLink>
<td align="center" height="25"><a href="?menu=applymanage&fashion=id">ID</a></td>
<td width="20%" align="center" height="25">
<a href="?menu=applymanage&fashion=bbsname">论坛</a></td>
<td align=center height="25"><a href="?menu=applymanage&fashion=hide">属性</a></td>
<td align=center height="25"><a href="?menu=applymanage&fashion=moderated">版主</a></td>
<td align=center height="25"><a href="?menu=applymanage&fashion=today">今日</a></td>
<td align=center height="25"><a href="?menu=applymanage&fashion=toltopic">主题</a></td>
<td align=center height="25"><a href="?menu=applymanage&fashion=tolrestore">帖子</a></td>
<td align="center" height="25">操作</td>
</tr>
<%

if Request("fashion")=empty then
fashion="tolrestore"
else
fashion=Request("fashion")
end if


sql="select * from bbsconfig order by "&fashion&" Desc"
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
<tr class=a3>
<td align="center" height="25"><%=rs("id")%></td>
<td><a target=_blank href=ShowForum.asp?forumid=<%=rs("id")%>><%=rs("bbsname")%></a></td>
<td align=center><%if ""&rs("hide")&""="1" then%><font color="#FF0000">隐藏</font><%elseif ""&rs("followid")&""="0" then%><font color="#0000FF">类别</font><%else%>正常<%end if%></td>
<td align=center width="200"><%=rs("moderated")%></td>
<td align=center><b><font color=red><%=rs("today")%></font></b></td>
<td align=center><b><font color=red><%=rs("toltopic")%></font></b></td>
<td align=center><b><font color=red><%=rs("tolrestore")%></font></b></td>
<td align="center"><a href=?menu=bbsmanagexiu&id=<%=rs("id")%>>编辑论坛</a> | <a onclick=checkclick('您确定要删除该论坛的所有资料?') href=?menu=bbsmanagedel&id=<%=rs("id")%>>删除论坛</a>
</tr>
<%
RS.MoveNext
loop
RS.Close

%>
</table>

[<b>
<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"menu=applymanage&fashion=<%=Request("fashion")%>")
</script>

</b>]
<%

end sub


sub classs
%>

<table border="0" width="80%">
	<tr>
		<td align="center">
<form name="form" method="post" action="?menu=bbsaddok"><input type=hidden name=classid value=0><input type=hidden name=hide value=0>
类别名称：（例如：电脑网络）<input name="bbsname" onkeyup="ValidateTextboxAdd(this, 'bbsname1')" onpropertychange="ValidateTextboxAdd(this, 'bbsname1')"><input type="submit" value="建立" id='bbsname1' disabled></form>
</td>
		<td align="center">
<form method="post" action="?menu=bbsmanagexiu">
查找论坛：<INPUT size=2 name=id onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')" onfocus="javascript:focusEdit(this)" onblur="javascript:blurEdit(this)" value="ID" helptext="ID">
<input type=submit value="确定" id='btnadd' disabled></form>
</td>
	</tr>
</table>


<table cellspacing=1 cellpadding="2" width="80%" border="0" class="a2" align="center">
<%
sort(0)
%>
</table>


<%

end sub

sub bbsadd

%>


<table cellspacing="1" cellpadding="2" width="80%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan="2">建立论坛资料</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>

<form name="form" method="post" action="?menu=bbsaddok">
<input type=hidden name=classid value=<%=id%>>

论坛名称</td>
    <td class=a3>

<input size="20" name="bbsname"></td></tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛版主</td>
    <td class=a3>

<input size="30" name=moderated> 多版主添加请用|分隔，如：yuzi|裕裕
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛介绍</td>
    <td class=a3>

<textarea rows="5" name="intro" cols="50"></textarea></td></tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛状态</td>
    <td class=a3>

<select size="1" name="pass">
<option value=0>关闭</option>
<option value=1 selected>正常</option>
<option value=2>游客止步</option>
<option value=3>授权浏览</option>
<option value=4>授权发帖</option>
</select>
</td></tr>




   <tr height=25>
    <td class=a3 align=middle>

授权用户名单</td>
    <td class=a3>

<textarea rows="3" name="userlist" cols="40"></textarea> 添加请用|分隔，如：yuzi|裕裕
</td></tr>



   <tr height=25>
    <td class=a3 align=middle>

小图标URL</td>
    <td class=a3>

<input size="30" name="icon"> 显示在社区首页论坛介绍右边
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

大图标URL</td>
    <td class=a3>

<input size="30" name="logo"> 显示在论坛左上角
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

通行密码</td>
    <td class=a3>

<input size="30" name="password"> 如果是公开论坛，此处请留空</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

是否显示在论坛列表　　　　　　　　　　　
    <td class=a3>

<input type="radio" CHECKED value="0" name="hide" id=hide1><label for=hide1>显示</label>
 
<input type="radio" value="1" name="hide" id=hide2><label for=hide2>隐藏</label> </tr>
   <tr height=25>
    <td class=a3 align=middle colspan="2">

　 <input type="submit" value=" 建 立 " name="Submit"><br></td></tr></table>
</form>
<center><br><a href=javascript:history.back()>< 返 回 ></a>


<%
end sub
sub bbsaddok
if Request("bbsname")="" then error2("请输入论坛名称")



master=split(Request("moderated"),"|")
for i = 0 to ubound(master)
If conn.Execute("Select id From [user] where username='"&HTMLEncode(master(i))&"'" ).eof and master(i)<>"" Then error2(""&master(i)&"的用户资料还未注册")
next


userlist=replace(Request("userlist"),vbCrlf,"")

rs.Open "bbsconfig",Conn,1,3
rs.addnew
rs("followid")=Request("classid")
rs("bbsname")=HTMLEncode(Request("bbsname"))
rs("moderated")=Request("moderated")
rs("pass")=Request("pass")
rs("password")=Request("password")
rs("userlist")=userlist
rs("intro")=HTMLEncode(Request("intro"))
rs("icon")=HTMLEncode(Request("icon"))
rs("logo")=HTMLEncode(Request("logo"))
rs("hide")=Request("hide")
rs.update
id=rs("id")

rs.close

classs

end sub




sub bbsmanagexiuok

if Request("bbsname")="" then error2("请输入论坛名称")



master=split(Request("moderated"),"|")
for i = 0 to ubound(master)
If conn.Execute("Select id From [user] where username='"&HTMLEncode(master(i))&"'" ).eof and master(i)<>"" Then error2(""&master(i)&"的用户资料还未注册")
next

userlist=replace(Request("userlist"),vbCrlf,"")

sql="select * from bbsconfig where id="&id&""
rs.Open sql,Conn,1,3


rs("followid")=Request("classid")
rs("SortNum")=int(Request("SortNum"))
rs("bbsname")=HTMLEncode(Request("bbsname"))
rs("moderated")=Request("moderated")
rs("pass")=Request("pass")
rs("password")=Request("password")
rs("userlist")=userlist
rs("intro")=HTMLEncode(Request("intro"))
rs("icon")=HTMLEncode(Request("icon"))
rs("logo")=HTMLEncode(Request("logo"))
rs("hide")=Request("hide")
rs.update

rs.close
%>
编辑成功<br><br><a href=javascript:history.back()>返 回</a>
<%
end sub


sub bbsmanagexiu

sql="select * from bbsconfig where id="&id&""
Set Rs=Conn.Execute(sql)
%>


<form method="post" action="?menu=bbsmanagexiuok" name=form><input type=hidden name=id value=<%=rs("id")%>>
<table cellspacing="1" cellpadding="2" width="80%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan="2">编辑论坛资料</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛名称</td>
    <td class=a3>


<input size="15" name="bbsname" value="<%=rs("bbsname")%>"> &nbsp; 排序 <input size="2" name="SortNum" value="<%=rs("SortNum")%>">
从小到大排序</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛类别</td>
    <td class=a3>


<select name="classid">
<option value="<%=rs("followid")%>">默认</option>
<option value="0">一级分类</option>
<%BBSList(0)%>
</select> 一级分类将当作类别名称</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


论坛版主</td>
    <td class=a3>


<input size="30" name=moderated value="<%=rs("moderated")%>"> 多版主添加请用“|”分隔，如：yuzi|裕裕
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


论坛介绍</td>
    <td class=a3>


<textarea rows="5" name="intro" cols="50"><%=rs("intro")%></textarea></td></tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛状态</td>
    <td class=a3>

<select size="1" name="pass">
<option value=0 <%if rs("pass")=0 then%>selected<%end if%>>关闭</option>
<option value=1 <%if rs("pass")=1 then%>selected<%end if%>>正常</option>
<option value=2 <%if rs("pass")=2 then%>selected<%end if%>>游客止步</option>
<option value=3 <%if rs("pass")=3 then%>selected<%end if%>>授权浏览</option>
<option value=4 <%if rs("pass")=4 then%>selected<%end if%>>授权发帖</option>
</select>
</td></tr>




   <tr height=25>
    <td class=a3 align=middle>

授权用户名单</td>
    <td class=a3>

<textarea rows="3" name="userlist" cols="40"><%=rs("userlist")%></textarea> 添加请用|分隔，如：yuzi|裕裕
</td></tr>

   <tr height=25>
    <td class=a3 align=middle>


小图标URL</td>
    <td class=a3>


<input size="30" name="icon" value="<%=rs("icon")%>"> 显示在社区首页论坛介绍右边
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


大图标URL</td>
    <td class=a3>


<input size="30" name="logo" value="<%=rs("logo")%>"> 显示在论坛左上角</td></tr>



   <tr height=25>
    <td class=a3 align=middle>

通行密码</td>
    <td class=a3>

<input size="30" name="password" value="<%=rs("password")%>"> 如果是公开论坛，此处请留空</td></tr>

   <tr height=25>
    <td class=a3 align=middle>


是否显示在论坛列表</td>
    <td class=a3>


<input type="radio" <%if rs("hide")=0 then%>CHECKED <%end if%>value="0" name="hide" value="0" id=hide1><label for=hide1>显示</label> 
<input type="radio" <%if rs("hide")=1 then%>CHECKED <%end if%>value="1" name="hide" value="1" id=hide2><label for=hide2>隐藏</label> </td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan="2">


<input type="submit" value=" 编 辑 " name="Submit1"></td></tr></table><br></form>
<center><br><a href=javascript:history.back()>< 返 回 ></a>
<Script>
document.form.intro.value = unybbcode(document.form.intro.value);
function unybbcode(temp) {
temp = temp.replace(/<br>/ig,"\n");
return (temp);
}
</Script>
<%
end sub



sub bbsmanage

%>

<script>
var BBSList="<%BBSList(0)%>"
</script>


论坛数：<b><font color=red><%=conn.execute("Select count(id)from bbsconfig")(0)%></font></b>　　主题数：<b><font color=red><%=conn.execute("Select count(id)from forum")(0)%></font></b><br>



<table cellspacing="1" cellpadding="2" width="80%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan="3">批量删除主题</td>
  </tr>
  
  <form method="post" action="?menu=deltopicok">

   <tr height=25>
    <td class=a3 align=middle width="50%">
删除 <INPUT size=3 name=TimeLimit value="180">
天以前的主题</td>
    <td class=a3 align=middle>
<select name="bbsid">
<option value="">所有论坛</option>
<script>document.write(BBSList);</script>
</select></td>
    <td class=a3 align=middle>
 <input type="submit" value=" 确 定 "></td></form></tr>
  
  
  
  
<form method="post" action="?menu=delforumok">
<tr height=25><td class=a3 align=middle width="50%">
删除 <INPUT size=3 name=TimeLimit value="180">
天没有更新的主题</td>
    <td class=a3 align=middle>
<select name="bbsid">
<option value="">所有论坛</option>
<script>document.write(BBSList);</script>
</select></td>
    <td class=a3 align=middle>
 <input type="submit" value=" 确 定 "></td></form></tr>

   <tr>
    <td class=a4 align=middle>
    <form method="post" action="?menu=delusertopicok">
删除 <input size="10" name="username"> 发表的所有主题
</tr>
    <td class=a4 align=middle>
    
<select name="bbsid">
<option value="">所有论坛</option>
<script>document.write(BBSList);</script>
</select></tr>
    <td class=a4 align=middle>
	<input type="submit" value=" 确 定 "></tr>
	</tr></form>

   <tr height=25>
    <td class=a4 align=middle>
    <form method="post" action="?menu=delliketopicok">
删除标题里包含有 <input size="10" name="topic"> 的所有主题
</tr>
    <td class=a4 align=middle>
    
<select name="bbsid">
<option value="">所有论坛</option>
<script>document.write(BBSList);</script>
</select></tr>
    <td class=a4 align=middle>
	<input type="submit" value=" 确 定 "></tr></form>

</table>

<br>

<table cellspacing="1" cellpadding="2" width="80%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>批量删除回帖</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle width="50%">


<form method="post" action="?menu=delretopicok">
删除 <INPUT size=3 name=TimeLimit value="180">
天以前的回帖 <input type="submit" value=" 确 定 "></td></form>

    <td class=a3 align=middle width="50%">
    <form method="post" action="?menu=deluserretopicok">
删除 <input size="10" name="username"> 发表的所有回帖 <input type="submit" value=" 确 定 "></td></form></tr>
    
   </table>


<form method="post" action="?menu=delbbsconfigok">
<table cellspacing="1" cellpadding="2" width="80%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle>批量删除论坛</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>
删除 <INPUT size=3 name=TimeLimit value="90">
天没有新帖子的论坛

<input type="submit" value=" 确 定 "></td>
    </tr>
    </table></form>



<form method="post" action="?menu=uniteok">
<table cellspacing="1" cellpadding="2" width="80%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=3>移动论坛帖子</td>
  </tr>
   <tr>
    <td class=a3 align=middle>

将 <select name="ybbs">
<script>document.write(BBSList);</script>
</select> 移动到
</td>
    <td class=a3 align=middle>

<select name="hbbs">
<script>document.write(BBSList);</script>
</select>
</td>
    <td class=a3 align=middle rowspan="2">

<INPUT type=submit value=" 确 定 " name=submit></td>
	</tr>
   <tr>
    <td class=a3 align=middle width="49%">

仅移动

<input size="2" name="TimeLimit" value="0"> 天前的帖子
</td>
    <td class=a3 align=middle>

仅移动

<input size="8" name="username"> 发表的帖子</td>
    </tr>
   </table></form>






<%
end sub

sub upclubconfig
%>

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>社区资料更新</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
    
此操作将更新论坛资料，修复论坛统计的信息<br>
<a href="?menu=upclubconfigok">点击这里更新论坛统计数据</a><br>


<a href="?menu=delapplication">清除服务器上的application缓存</a><br>



</td></tr></table><br>

<%
end sub

sub upclubconfigok


rs.Open "bbsconfig",Conn
do while not rs.eof

allarticle=conn.execute("Select count(forumid) from forum where deltopic<>1 and forumid="&rs("id")&"")(0)
if allarticle>0 then
allrearticle=conn.execute("Select sum(replies) from forum where deltopic<>1 and forumid="&rs("id")&"")(0)
else
allrearticle=0
end if

conn.execute("update [bbsconfig] set toltopic="&allarticle&",tolrestore="&allarticle+allrearticle&" where ID="&rs("id")&"")


rs.movenext
loop
rs.close

%>
操作成功！<br>
<br>
已经更新社区所有论坛的统计数据<br>

<%
end sub

htmlend

sub activation
%>
<script>
function CheckAll(form){for (var i=0;i<form.elements.length;i++){var e = form.elements[i];if (e.name != 'chkall')e.checked = form.chkall.checked;}}
</script>
  <form method="POST" action=?menu=activationok>

<TABLE cellSpacing=1 cellPadding=1 width=97% align=center border=0 class=a2>
<TR height=25 class=a1>
		<td align="center">激活</td>
		<td align="center">主题</td>
		<td align="center">作者</td>
		<td align="center">发表时间</td></tr>
<%
sql="select * from forum where deltopic=1 order by id Desc"
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
      <TD width="60%" align=left class=a4>&nbsp;<%if rs("icon")>0 then%><img border=0 src=images/brow/<%=rs("icon")%>.gif> <%end if%><a target=_blank href=ShowPost.asp?id=<%=rs("id")%>><%=rs("topic")%></a></TD>
      <TD align=middle width=9% class=a3><a href=Profile.asp?username=<%=rs("username")%>><%=rs("username")%>　</TD>

      <TD align=center width=30% class=a4>&nbsp;<%=rs("posttime")%></TD>
    </TR>

<%

RS.MoveNext
loop
RS.Close


%>


<TR height=25>

      <TD width="6%" class=a3 align="center"><input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON"></TD>

      <TD width="60%" class=a3>		
<b>&nbsp;共有
<font color="990000"><%=TotalPage%></font> 页 [
<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"menu=activation")
</script>]</b>
		
</TD>

      <TD width="48%" class=a3 colspan="2" align="center">
		
		
<input type=submit onclick="checkclick('您确定要激活所选的主题?');" value=" 激 活 "></TD>
    </TR>
    


</table><%
end sub

dim ii
ii=0
sub sort(selec)
	sql="Select * From bbsconfig where followid="&selec&" and hide=0 order by SortNum"
	Set Rs1=Conn.Execute(sql)


	do while not rs1.eof

if selec=0 then
%>
  <tr class=a1 id=TableTitleLink height=25>
<td>　<a target=_blank href=ShowForum.asp?forumid=<%=rs1("id")%>><%=rs1("bbsname")%></a></td>
<td align="right" width="190">
<a href=?menu=bbsadd&id=<%=rs1("id")%>>建立论坛</a> | <a href=?menu=bbsmanagexiu&id=<%=rs1("id")%>>编辑论坛</a> | 
<a onclick=checkclick('您确定要删除该论坛的所有资料?') href=?menu=bbsmanagedel&id=<%=rs1("id")%>>删除论坛</a>
</tr>

<%
else
%>
<tr class=a3 height=25>
<td>　<%=string(ii*2,"　")%><a target=_blank href=ShowForum.asp?forumid=<%=rs1("id")%>><%=rs1("bbsname")%></a></td>
<td align="right">
<a href=?menu=bbsadd&id=<%=rs1("id")%>>建立论坛</a> | <a href=?menu=bbsmanagexiu&id=<%=rs1("id")%>>编辑论坛</a> | 
<a onclick=checkclick('您确定要删除该论坛的所有资料?') href=?menu=bbsmanagedel&id=<%=rs1("id")%>>删除论坛</a>
</tr>
<%
end if
ii=ii+1
	sort rs1("id")
ii=ii-1
	rs1.movenext
	loop
	Set Rs1 = Nothing
end sub

%>