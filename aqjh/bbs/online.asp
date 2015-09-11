<!-- #include file="setup.asp" -->


<%
top
count=conn.execute("Select count(sessionid)from online")(0)
regonline=conn.execute("Select count(sessionid)from online where username<>''")(0)
toltopic=conn.execute("Select SUM(toltopic)from bbsconfig")(0)
tolretopic=conn.execute("Select SUM(tolrestore)from bbsconfig")(0)
%>

<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 查看论坛状态</td>
</tr>
</table><br>

<SCRIPT>valigntop()</SCRIPT>
<TABLE cellspacing="1" cellpadding="4" width="97%" align="center" border="0" class="a2">
  
  <TR class="a1" id=TableTitleLink>
      <Td width="16%" align="center" height="10"><b><font color="FFFFFF"><a href="online.asp"> 
        在线情况</a></font></TH> </b> 
      <TD width="16%" align="center" height="10"><b><font color="FFFFFF"> <a href="online.asp?menu=cutline">在线图例</TH> 
        </a> </font> </b> 
      <TD width="16%" align="center" height="10">
		<b>
		<a href="online.asp?menu=sex">性别图例</a></b><TD width="16%" align="center" height="10">
		<b><a href="online.asp?menu=todaypage">今日图例</a></b><TD width="16%" align="center" height="10"><b><font color="FFFFFF"><a href="online.asp?menu=board"> 
        主题图例</a></font></TH> </b> 
      <TD width="16%" align="center" height="10"><b><font color="FFFFFF"> <a href="online.asp?menu=tolrestore">帖子图例</TH> 
        </a> </font> </b> 
      <TR class=a3>
      <Td width="48%" align="center" height="10" colspan="3">总帖数 <%=tolretopic%> 篇。其中主题 <%=toltopic%> 篇，回帖 <%=tolretopic-toltopic%> 
		篇。<Td width="48%" align="center" height="10" colspan="3">总在线 <%=count%> 人。其中注册用户 <%=regonline%> 人，访客 <%=count-regonline%> 人。</TABLE>
<SCRIPT>valignbottom()</SCRIPT>
<br>
<%

select case Request("menu")
case ""
index
case "cutline"
cutline

case "board"
board

case "tolrestore"
tolrestore

case "todaypage"
todaypage

case "sex"
sex

end select

sub index

ip=HTMLEncode(Request.Form("ip"))
username=HTMLEncode(Request.Form("username"))

if ip<>"" then whereip=" where ip='"&ip&"'"
if username<>"" then whereusername=" where username='"&username&"' and eremite='0'"


sql="select * from online "&whereip&" "&whereusername&" order by lasttime Desc"
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




if membercode<4 then
ips=split(rs("ip"),".")
ip=""&ips(0)&"."&ips(1)&".*.*"
else
ip=""&rs("ip")&""
end if


if rs("username")="" then
username="访 客"
elseif rs("eremite")=1 and membercode<4 then
username="隐 身"
else
username="<a href=Profile.asp?username="&rs("username")&">"&rs("username")&"</a>"
end if



place2=""
if rs("act")<>"" then
place2 = "<a href="&rs("acturl")&">"&rs("act")&"</a>"
place = "『 "&rs("bbsname")&" 』"
else
place = "『 <a href="&rs("acturl")&">"&rs("bbsname")&"</a> 』"
end if
allline=""&allline&"<TR align=middle class=a4><TD height=24>"&ip&"</TD><TD height=24>"&rs("cometime")&"</TD><TD height=24>"&username&"</TD><TD height=24>"&place&"</TD><TD height=24>"&place2&"</TD><TD height=24>"&rs("lasttime")&"</TD></TR>"

rs.movenext
loop
rs.close



%>

<SCRIPT>valigntop()</SCRIPT>
<TABLE cellSpacing=1 cellPadding=1 width=97% align=center border=0 class=a2>



<TR align=middle class=a1>

<TD height=23>IP地址</TD>
<TD height=23>登录时间</TD>
<TD height=23>用户名</TD>
<TD height=23>所在论坛</TD>
<TD height=23>所在主题</TD>
<TD height=23>活动时间</TD>

</TR>

<%=allline%>

</TABLE>
<SCRIPT>valignbottom()</SCRIPT>
<center>

<table border=0 width=97%>
	<tr>
		<td align="center"><form action="online.asp" method="post">查询ＩＰ：<input size="15" value="<%=Request.Form("ip")%>" name="ip"> <input type="submit" value=" 确定 "></td></form>
		<td align="center" valign="bottom">Page：[
<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"")
</script>
 ]</td>
		<td align="center"><form action="online.asp" method="post">查询用户：<input size="15" value="<%=Request.Form("username")%>" name="username"> <input type="submit" value=" 确定 "></td></form>
	</tr>
</table>
<p><br>
</p>




<%


end sub


sub cutline


sql="select * from bbsconfig where followid<>0 and hide=0"
Set Rs=Conn.Execute(sql)



%>


<SCRIPT>valigntop()</SCRIPT>
<table class="a2" cellSpacing="1" cellPadding="4" width="97%" align="center" border="0">
		<tr>
			<td class="a1" vAlign="bottom" align="middle" height="20">论坛名称</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">图形比例</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">在线人数</td>
		</tr>
		


<%

i=0
Do While Not RS.EOF
onlinemany=conn.execute("Select count(sessionid)from online where forumid="&rs("id")&"")(0)

%>

		<tr class="a3">
			<td width="21%" height="2" align="center">
			<a href="ShowForum.asp?forumid=<%=rs("id")%>"><%=rs("bbsname")%></a></td>
			<td width="65%" height="2"><IMG height=8 src="images/bar/<%=i%>.gif" width="<%=onlinemany/count*100%>%"></td>
			<td align="center" width="12%" height="2"><%=onlinemany%></td></tr>
			
			
			<%
			
i=i+1
if i=10 then i=0

RS.MoveNext
loop
RS.Close   


%></table>

<SCRIPT>valignbottom()</SCRIPT>
<%


end sub

sub board

sql="select * from bbsconfig where followid<>0 and hide=0 order by toltopic Desc"
Set Rs=Conn.Execute(sql)
%>
<SCRIPT>valigntop()</SCRIPT>
<table class="a2" cellSpacing="1" cellPadding="4" width="97%" align="center" border="0">
		<tr>
			<td class="a1" vAlign="bottom" align="middle" height="20">论坛名称</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">图形比例</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">主题数</td>
		</tr>
		


<%
i=0
Do While Not RS.EOF
%>

		<tr class="a3">
			<td width="21%" height="2" align="center"><a href="ShowForum.asp?forumid=<%=rs("id")%>"><%=rs("bbsname")%></a></td>
			<td width="65%" height="2"><IMG height=8 src="images/bar/<%=i%>.gif" width="<%=rs("toltopic")/toltopic*100%>%"></td>
			<td align="center" width="12%" height="2"><%=rs("toltopic")%></td></tr><%


i=i+1
if i=10 then i=0


			
RS.MoveNext
loop
RS.Close   


%></table>
<SCRIPT>valignbottom()</SCRIPT>
<%
end sub

sub tolrestore

sql="select * from bbsconfig where followid<>0 and hide=0 order by tolrestore Desc"
Set Rs=Conn.Execute(sql)
%>
<SCRIPT>valigntop()</SCRIPT>
<table class="a2" cellSpacing="1" cellPadding="4" width="97%" align="center" border="0">
		<tr>
			<td class="a1" vAlign="bottom" align="middle" height="20">论坛名称</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">图形比例</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">帖子数</td>
		</tr>
<%
i=0
Do While Not RS.EOF
%>

		<tr class="a3">
			<td width="21%" height="2" align="center"><a href="ShowForum.asp?forumid=<%=rs("id")%>"><%=rs("bbsname")%></a></td>
			<td width="65%" height="2"><IMG height=8 src="images/bar/<%=i%>.gif" width="<%=rs("tolrestore")/tolretopic*100%>%"></td>
			<td align="center" width="12%" height="2"><%=rs("tolrestore")%></td></tr><%
			
i=i+1
if i=10 then i=0

			
RS.MoveNext
loop
RS.Close   


%></table>
<SCRIPT>valignbottom()</SCRIPT>
<%
end sub



sub todaypage
toltoday=conn.execute("Select SUM(today)from bbsconfig")(0)
if toltoday=0 then toltoday=1
sql="select * from bbsconfig where followid<>0 and hide=0 order by today Desc"
Set Rs=Conn.Execute(sql)
%>
<SCRIPT>valigntop()</SCRIPT>
<table class="a2" cellSpacing="1" cellPadding="4" width="97%" align="center" border="0">
		<tr>
			<td class="a1" vAlign="bottom" align="middle" height="20">论坛名称</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">图形比例</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">今日帖数</td>
		</tr>
		


<%
i=0
Do While Not RS.EOF
%>

		<tr class="a3">
			<td width="21%" height="2" align="center"><a href="ShowForum.asp?forumid=<%=rs("id")%>"><%=rs("bbsname")%></a></td>
			<td width="65%" height="2"><IMG height=8 src="images/bar/<%=i%>.gif" width="<%=rs("today")/toltoday*100%>%"></td>
			<td align="center" width="12%" height="2"><%=rs("today")%></td></tr><%
			
i=i+1
if i=10 then i=0

			
RS.MoveNext
loop
RS.Close   


%></table>
<SCRIPT>valignbottom()</SCRIPT>
<%
end sub





sub sex
count=conn.execute("Select count(id)from [user]")(0)
male=conn.execute("Select count(id)from [user] where sex='male'")(0)
female=conn.execute("Select count(id)from [user] where sex='female'")(0)

%>

<SCRIPT>valigntop()</SCRIPT>
<table class="a2" cellSpacing="1" cellPadding="4" width="97%" align="center" border="0">
		<tr>
			<td class="a1" vAlign="bottom" align="middle" height="20">性别</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">图形比例</td>
			<td class="a1" vAlign="bottom" align="middle" height="20">人数</td>
		</tr>
		



		<tr class="a3">
			<td width="10%" height="2" align="center">
			<img src=images/male.gif></td>
			<td width="75%" height="2"><IMG height=8 src="images/bar/7.gif" width="<%=male/count*100%>%"></td>
			<td align="center" width="12%" height="2"><%=male%></td></tr>		
			
		<tr class="a3">
			<td width="10%" height="2" align="center">
			<img src=images/female.gif></td>
			<td width="75%" height="2"><IMG height=8 src="images/bar/0.gif" width="<%=female/count*100%>%"></td>
			<td align="center" width="12%" height="2"><%=female%></td></tr>		
			
		<tr class="a3">
			<td width="10%" height="2" align="center">未知</td>
			<td width="75%" height="2"><IMG height=8 src="images/bar/2.gif" width="<%=(count-male-female)/count*100%>%"></td>
			<td align="center" width="12%" height="2"><%=count-male-female%></td></tr>		
	

</table>
<SCRIPT>valignbottom()</SCRIPT>
<%
end sub


htmlend
%>