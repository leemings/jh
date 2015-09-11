<!-- #include file="setup.asp" -->
<%
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")
username=Trim(HTMLEncode(Request("username")))

if Request.Form("menu")="ok" then

if membercode < 4 then error("<li>您的权限不够，无法抓人入狱！")

If conn.Execute("Select username From [user] where username='"&username&"'" ).eof Then error("<li>"&username&"的资料不存在")
If not conn.Execute("Select username From [prison] where username='"&username&"'" ).eof Then error("<li>"&username&"已经被关进监狱")


if username="" then error("<li>犯人的名称没有添写")


causation=HTMLEncode(Request("causation"))
prisonday=HTMLEncode(Request("prisonday"))

if causation="" then error("<li>您没有输入入狱原因")
if prisonday>1000 then error("<li>坐牢时间不能超过1000天")

sql="insert into prison (username,causation,constable,prisonday) values ('"&username&"','"&causation&"','"&Request.Cookies("username")&"','"&prisonday&"')"
conn.Execute(SQL)



end if

if Request("menu")="release" then
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")

if membercode < 4 then error("<li>您的权限不够，无法释放犯人！")

conn.execute("delete from [prison] where username='"&username&"'")

log("将 "&username&" 释放出监狱！")

error2("已经将 "&username&" 释放出监狱！")
end if


if Request("menu")="look" then
sql="select * from prison where username='"&username&"'"
Set Rs=Conn.Execute(sql)


%>
<title>探 监 - Powered By BBSxp</title>
<b><%=rs("username")%></b>
<SCRIPT>
var tips=["斜着眼睛瞟了一眼看守,嘟哝着:最近的点儿太背!","两眼汪汪的说:都是我不好!对不起大家了!","脸上露出诡异的笑容:嘿嘿!要不要进来看看!","感慨万分道:一失足,成千古恨!我一定重新做人!","望着布满电网和铁丝网的高墙,摇头叹息着!"]
index = Math.floor(Math.random() * tips.length);
document.write("" + tips[index] + "");
  </SCRIPT><br><br>
入狱原因：<%=rs("causation")%><br><br>
入狱时间：<%=rs("cometime")%><br><br>
出狱时间：<%=rs("cometime")+rs("prisonday")%><br><br>
执行警官：<%=rs("constable")%>

<%
rs.close
responseend

end if

top

conn.execute("delete from [prison] where cometime<"&SqlNowString&"-prisonday")
%>

<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 
<a href="prison.asp">社区监狱</a></td>
</tr>
</table><br>

<SCRIPT>valigntop()</SCRIPT>

<TABLE cellSpacing=1 cellPadding=3 border=0 width=97% align=center class=a2><TR class=a1 height="25">
	<TD align=center width="10%"><b>用户名</b></TD>
<TD align=center><b>入狱原因</b></TD>
<TD align=center width="20%"><b>入狱时间</b></TD>
<TD align=center width="10%"><b>执行警官</b></TD>
<TD align=center width="15%"><b>动作</b></TD>
</TR>

<%
rs.Open "prison order by cometime Desc",Conn,1
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


response.write "<tr class=a4><TD align=center><a href=Profile.asp?username="&rs("username")&">"&rs("username")&"</a></TD><TD>"&rs("causation")&"</TD><TD align=center>"&rs("cometime")&"</TD><TD align=center><a href=Profile.asp?username="&rs("constable")&">"&rs("constable")&"</a></TD><TD align=center><a href=?menu=release&sessionid="&session.sessionid&"&username="&rs("username")&">释 放</a> | <a href=# onClick=javascript:open('prison.asp?menu=look&username="&rs("username")&"','','resizable,scrollbars,width=220,height=180')>探 监</a></TD></tr>"


RS.MoveNext
loop
RS.Close

%>

<form METHOD=POST><input type=hidden name=menu value=ok><tr>
  <TD align=center colspan="5" class=a3>将
<input name="username" size="13"> 抓入监狱　　坐牢天数：<input name="prisonday" size="2" value="15"><br>
入狱原因：<input name="causation" size="33"> <input type="submit" value=" 确 定 "></TD>
  			</tr></form>
</TABLE>
<SCRIPT>valignbottom()</SCRIPT><center>
[<b>
<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"")
</script>

</b>]

<%
htmlend
%>