<!-- #include file="setup.asp" -->
<%
top
%>

<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 管理团队</td>
</tr>
</table><br>

<SCRIPT>valigntop()</SCRIPT>

<table cellpadding=3 cellspacing=1 border=0 width=97% height="25" class=a2 align=center><tr>
<td align=center style="font-size: 9pt" class=a1 height="17" width="30%">
<b>管 理 人 员</b></td>
<td align=center style="font-size: 9pt" class=a1 height="17">
<b>详 细 信 息</b></td></tr>
<%

sql="select * from [user] where membercode > 3 order by membercode Desc,landtime Desc"
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
<tr class=a3><td width="30%" align="center" valign="top">
<script>
if("<%=rs("userphoto")%>"!=""){
document.write("<a target=_blank href=<%=rs("userphoto")%>><img src=<%=rs("userphoto")%> border=0 onload='javascript:if(this.width>200)this.width=200'></a>")
}
</script>

</td><td valign=top>
<table cellSpacing=0 cellPadding=5 border=0 style="border-collapse: collapse" style="TABLE-LAYOUT: fixed">
<tr><td align=middle width="50">
<img src=<%=rs("userface")%> width="32" height="32"></td>
<td width="80"><b>昵 称</b></td>
<td align=left width="120"><a href=Profile.asp?username=<%=rs("username")%>><%=rs("username")%></a>
<td align=left width="30">
<img src=images/add.gif align=abscenter border=0 width="16" height="16"></td>
<td align=left width="80"><a href=friend.asp?menu=add&username=<%=rs("username")%>> 加为好友</a></td>
<td align=left>
<img src=images/msg.gif width="16" height="16"> <a href=# onclick=javascript:open('friend.asp?menu=post&incept=<%=rs("username")%>','','width=320,height=170')> 发送讯息</a></td>
</tr><tr><td align=middle width="50">
<img src=images/level.gif border=0 width="16" height="16"></td>
<td width="80"><b>用户等级</b></td>
<td align=left width="120"><Script>document.write(level(<%=rs("experience")%>,<%=rs("membercode")%>,'','')+levelname);</Script></td>
<td align=left width="30">
<img src=images/mail.gif border=0 width="16" height="16"></td>
<td align=left width="80"><b>EMAIL</b></td><td align=left><a href=mailto:<%=rs("usermail")%>><%=rs("usermail")%></a>
</tr><tr><td align=middle width="50">
<img src=images/registered.gif border=0 width="16" height="16"></td>
<td width="80"><b>注册日期</b></td><td align=left width="120"><%=rs("regtime")%>
<td align=left width="30">
<img src=images/home.gif border=0 width="16" height="16"></td>
<td align=left width="80"><b>主页地址</b></td>
<td align=left><a target="_blank" href="<%=rs("userhome")%>"><%=rs("userhome")%></a></tr><tr>
<td align=middle width="50">
<img src=images/posts.gif border=0 width="16" height="16"></td>
<td width="80"><b>总发贴数</b></td>
<td align=left width="120"><%=rs("posttopic")+rs("postrevert")%>
<td align=left width="110" colspan="2">
<b>最后登录时间</b></td>
<td align=left><%=rs("landtime")%>
</td>
</tr><tr>
<td>
</td>
<td colspan="5"><%=personal%>
　</td>
</tr>
</table></td></tr>
<%
RS.MoveNext
loop
RS.Close

%></table><center>
<SCRIPT>valignbottom()</SCRIPT>
[<b>
<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"")
</script>

</b>]<br><%
htmlend
%>