<!-- #include file="setup.asp" --><%

if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")


id=int(Request("id"))

url=HTMLEncode(Request("url"))
name=HTMLEncode(Request("name"))

top


select case Request("menu")
case "add"

If not conn.Execute("Select id From [favorites] where username='"&Request.Cookies("username")&"' and name='"&name&"' and url='"&url&"'" ).eof Then error("<li>收藏夹中已经存在此资料")
conn.execute("insert into favorites(username,name,url)values('"&Request.Cookies("username")&"','"&name&"','"&url&"')")
message="<li>添加成功<li><a href="&Request.ServerVariables("http_referer")&">"&Request.ServerVariables("http_referer")&"</a><li><a href=Default.asp>返回社区首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url="&Request.ServerVariables("http_referer")&">")


case "delweb"
conn.execute("delete from [favorites] where username='"&Request.Cookies("username")&"' and id="&id&"")
message="<li>删除成功<li><a href="&Request.ServerVariables("http_referer")&">"&Request.ServerVariables("http_referer")&"</a><li><a href=Default.asp>返回社区首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url="&Request.ServerVariables("http_referer")&">")


case "del"
conn.execute("delete from [favorites] where username='"&Request.Cookies("username")&"' and url='"&url&"' and name='"&name&"'")
message="<li>删除成功<li><a href="&Request.ServerVariables("http_referer")&">"&Request.ServerVariables("http_referer")&"</a><li><a href=Default.asp>返回社区首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url="&Request.ServerVariables("http_referer")&">")


end select

%>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 
收藏夹</td>
</tr>
</table><br>

<SCRIPT>valigntop()</SCRIPT>
<%

pagesetup=10 '设定每页的显示数量

select case Request("menu")
case ""
%>

<table width="97%" border=0 align=center cellPadding=3 cellSpacing=1 class=a2>
<tr class=a1><td width="69%" align="center" bgcolor="FFFFFF" height="20"><b>标 题</b></td>
	<td width="20%" align="center" bgcolor="FFFFFF" height="20">
		<b>添加时间</b></td>
	<td width="11%" align="center" bgcolor="FFFFFF" height="20"><b>操作</b></td></tr><%

sql="select * from favorites where username='"&Request.Cookies("username")&"' and url<>'topic' and url<>'forum' order by id Desc"
rs.Open sql,Conn,1

rs.pagesize=pagesetup
TotalPage=rs.pagecount  '总页数
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '跳转到指定页数

i=0 
Do While Not RS.EOF and i<pagesetup 
i=i+1 

%> <tr><td bgcolor="FFFFFF"><img src=images/ie.gif> <a target=_blank href="<%=rs("url")%>"><%=rs("name")%></a></td><td align=center bgcolor="FFFFFF"><%=rs("addtime")%><td align=center bgcolor="FFFFFF"><a href="?menu=delweb&id=<%=rs("id")%>"><img src=images/del.gif border=0></a></td></tr><%          
RS.MoveNext
loop
RS.Close      
%> <tr><td vAlign="top" colSpan="3"  class=a1 align="center"><b>&gt;&gt; 添加新链接 &lt;&lt;</b></td></tr><tr><td align="center" colspan="3" bgcolor="FFFFFF"><form method=post name=form action=favorites.asp><input type=hidden name=menu value=add>
<b>名称：</b><INPUT size=20 name=name>　<b>链接地址：</b><INPUT size=40 name=url value="http://">　<input type=submit value="添 加" name="Submit"> </td></form></tr></table>


<%

case "topic"
%>

<script>
var key_word="<%=Request.Cookies("key_word")%>"
var cookieusername="<%=Request.Cookies("username")%>"
</script>

    <table  cellSpacing="1" cellPadding="0" width=97% align="center" border="0" class=a2>
  <tr height="25">
    <td width="3%" class=a1>　</td>
    <td width="3%" class=a1></td>
    <td align="middle" height="24" class=a1 width="45%"> 主题</td>
    <td align="middle" width="9%" height="24" class=a1> 作者</td>
    <td align="middle" width="6%" height="24" class=a1> 回复</td>
    <td align="middle" width="7%" height="24" class=a1> 点击</td>
    <td width="27%" height="24" class=a1 align="center">最后更新</td>
  </tr>

<%
sql="select * from favorites where username='"&Request.Cookies("username")&"' and url='topic' order by id Desc"
rs.Open sql,Conn,1
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '总页数
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '跳转到指定页数


i=0
Do While Not RS.EOF and i<pagesetup
i=i+1


if not isnumeric(rs("name")) then conn.execute("delete from [favorites] where id="&rs("id")&"")

sql="Select * From forum where id="&rs("name")&""
Set Rs1=Conn.Execute(sql)

if Rs1.eof then conn.execute("delete from [favorites] where id="&rs("id")&"")


newtopic=""
if rs1("posttime")+1>now() then newtopic="<img src=images/new.gif>"
%>
<script>ShowForum("<%=rs1("ID")%>","<%=rs1("topic")%>","<%=newtopic%>","<%=rs1("username")%>","<%=rs1("Views")%>","<%=rs1("icon")%>","<%=rs1("toptopic")%>","<%=rs1("locktopic")%>","<%=rs1("pollresult")%>","<%=rs1("goodtopic")%>","<%=rs1("replies")%>","<%=rs1("lastname")%>","<%=rs1("lasttime")%>");</script>
<%
Set Rs1 = Nothing


RS.MoveNext
loop
RS.Close



%>
</table>


<%


case "forum"

%>
<table cellspacing=1 cellpadding=4 width=97% align=center border=0 class=a2>
<%
sql="select * from favorites where username='"&Request.Cookies("username")&"' and url='forum' order by id Desc"
rs.Open sql,Conn,1
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '总页数
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '跳转到指定页数


i=0
Do While Not RS.EOF and i<pagesetup
i=i+1

if not isnumeric(rs("name")) then conn.execute("delete from [favorites] where id="&rs("id")&"")

sql="Select * From bbsconfig where id="&rs("name")&""
Set Rs1=Conn.Execute(sql)

if Rs1.eof then conn.execute("delete from [favorites] where id="&rs("id")&"")


if ""&Rs1("password")&""="" then:password=0:else:password=1:end if
%><script>ShowList("<%=Rs1("pass")%>","<%=Rs1("id")%>","<%=Rs1("bbsname")%>","<%=Rs1("icon")%>","<%=Rs1("intro")%>","<%=Rs1("moderated")%>","<%=Rs1("today")%>","<%=Rs1("toltopic")%>","<%=Rs1("tolrestore")%>","<%=Rs1("lasttime")%>","<%=Rs1("lastname")%>","<%=Rs1("lasttopic")%>","<%=password%>");</script>
<%
Set Rs1 = Nothing


RS.MoveNext
loop
RS.Close



%>
</table>

<%

end select

%>

<SCRIPT>valignbottom()</SCRIPT>


<center><br>Page：[ 

<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"menu=<%=Request("menu")%>")
</script>
 

 ]<br><%
htmlend
%>