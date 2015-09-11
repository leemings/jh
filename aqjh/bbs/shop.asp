<!-- #include file="setup.asp" -->
<%
top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")

sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3
money=rs("money")
experience=rs("experience")
rs.close


select case Request("menu")

case "experience"
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")
how=int(Request("how"))
if how<1 then error("<li>数量不能小于1")
if money<300*how then error("<li>您的金币不够！")
Randomize
d1=fix(rnd*300)+1
sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3
rs("experience")=rs("experience")+d1*how
rs("money")=rs("money")-300*how
rs.update
rs.close

error2("已经增加了 "&d1*how&" 点经验值！")




case "VIPUser"
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")
if membercode>1 then error("<li>必须是普通会员才可购买！")
if experience<200000 then error("<li>您的经验小于200000")
if money<10000 then error("<li>您的金币不够！")
conn.execute("update [user] set membercode=2,[money]=[money]-10000 where username='"&Request.Cookies("username")&"'")
error2("您已经成功升级为贵宾会员！")


case "colony"
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")
content=HTMLEncode(Request("content"))

howmoney=conn.execute("Select count(sessionid)from online where username<>''")(0)*10

if content="" then error("<li>请填写群发内容!")
if money<howmoney then error("<li>您的金币不够！<li>本次共需 "&howmoney&" 金币")

conn.execute("update [user] set [money]=[money]-"&howmoney&" where username='"&Request.Cookies("username")&"'")

sql="select username from online where username<>''"
Set Rs=Conn.Execute(sql)
do while not rs.eof
conn.Execute("insert into message (author,incept,content) values ('"&Request.Cookies("username")&"','"&rs("username")&"','<font color=0000FF>【会员广播】："&content&"</font>')")
conn.execute("update [user] set newmessage=newmessage+1 where username='"&rs("username")&"'")
rs.movenext
loop
Set Rs = Nothing
error2("发送成功！\n\n共花费 "&howmoney&" 金币")



case "thew"
if money<100 then error("<li>您的金币不够！")

userlife=Conn.Execute("Select userlife From [user] where username='"&Request.Cookies("username")&"'")(0)
if userlife=100 then error("<li>体力已满，无需增加！")

conn.execute("update [user] set userlife=100,[money]=[money]-100 where username='"&Request.Cookies("username")&"'")
error2("您的体力已经全满！")

case "flowers"
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")

if money<50 then error("<li>您的金币不够！")

vs=HTMLEncode(Request("vs"))
if vs=Request.Cookies("username") then error("<li>不能自己送自己！")
if Application(CacheName&"flowers")=vs then error("<li>请勿对同一个用户进行操作！")
If conn.Execute("Select id From [user] where username='"&vs&"'" ).eof Then error("<li>系统不存在"&vs&"的资料")

conn.execute("update [user] set [money]=[money]-50 where username='"&Request.Cookies("username")&"'")


Randomize
d1=fix(rnd*50)+1

conn.execute("update [user] set experience=experience+"&d1&" where username='"&vs&"'")

sql="insert into message(author,incept,content) values ('"&Request.Cookies("username")&"','"&vs&"','<font color=0000FF>【系统消息】："&Request.Cookies("username")&"送您1束鲜花，您增加了"&d1&"点经验值！</font>')"
conn.Execute(SQL)
conn.execute("update [user] set newmessage=newmessage+1 where username='"&vs&"'")

Application(CacheName&"flowers") = vs

error2(""&vs&"已经增加了"&d1&"点经验值！")



case "egg"
if ""&Request("sessionid")&""<>""&session.sessionid&"" then error("<li>效验码错误<li>请重新返回刷新后再试")
if money<50 then error("<li>您的金币不够！")

vs=HTMLEncode(Request("vs"))
if vs=Request.Cookies("username") then error("<li>不能自己送自己！")

if Application(CacheName&"egg")=vs then error("<li>请勿对同一个用户进行操作！")

If conn.Execute("Select id From [user] where username='"&vs&"'" ).eof Then error("<li>系统不存在"&vs&"的资料")

If conn.Execute("Select experience From [user] where username='"&vs&"'" )(0)<50 Then error("<li>对方经验值少于50，您不能再向他（她）丢鸡蛋了")

conn.execute("update [user] set [money]=[money]-50 where username='"&Request.Cookies("username")&"'")


Randomize
d1=fix(rnd*50)+1

conn.execute("update [user] set experience=experience-"&d1&" where username='"&vs&"'")

sql="insert into message(author,incept,content) values ('"&Request.Cookies("username")&"','"&vs&"','<font color=0000FF>【系统消息】："&Request.Cookies("username")&"送您1粒鸡蛋，您减少了"&d1&"点经验值！</font>')"
conn.Execute(SQL)
conn.execute("update [user] set newmessage=newmessage+1 where username='"&vs&"'")

Application(CacheName&"egg") = vs
error2(""&vs&"已经减少了"&d1&"点经验值！")

end select

%>
<title>社区商场</title>

<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <a href="shop.asp">社区商店</a></td>
</tr>
</table><br>


<SCRIPT>
function colony(){
var id=prompt("请输入您要群发的讯息！","");if(id){document.location='?menu=colony&sessionid=<%=session.sessionid%>&content='+id+'';}
}

function experience(){
var id=prompt("请输入您要购买的数量！","1");if(id){document.location='?menu=experience&sessionid=<%=session.sessionid%>&how='+id+'';}
}

function flowers(){
var id=prompt("请输入对方用户名！","");if(id){document.location='?menu=flowers&sessionid=<%=session.sessionid%>&vs='+id+'';}
}

function egg(){
var id=prompt("请输入对方用户名！","");if(id){document.location='?menu=egg&sessionid=<%=session.sessionid%>&vs='+id+'';}
}

function VIPUser(){
document.location='?menu=VIPUser&sessionid=<%=session.sessionid%>';
}

</SCRIPT>
<center>



<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="111111" width="90%">
<tr>
<td width="100%">
<div align="center">
<center>
<table border="0" cellpadding="3" cellspacing="1" width="100%" class=a2>
<tr class=a1>
<td width="50%" height="5" align="center" colspan="2"><b>
体力药丸</b></td>
<td width="50%" height="5" align="center" colspan="2"><b>
经验药丸</b></td>
</tr>
<tr>
<td width="15%" height="4" align="center" bgcolor="FFFFFF" rowspan="2"><FONT color=cccccc>
<IMG src="images/plus/life.gif" border=1></FONT></td>
<td width="25%" height="2" bgcolor="FFFFFF">名称：体力药丸<br>
价格：100 <b>金币</b><br>
功能：快速补充体力</td>
<td width="15%" height="4" align="center" bgcolor="FFFFFF" rowspan="2"><FONT color=cccccc>
<IMG src="images/plus/experience.gif" border=1></FONT></td>
<td width="25%" height="2" bgcolor="FFFFFF">
  名称：经验药丸<br>
  价格：300 <b>金币</b><br>
  功能：提升经验</td>
</tr>
<tr>
<td width="25%" height="2" bgcolor="FFFFFF"><form method=POST action=?menu=thew>
<input type="submit" value="我要购买" name="Submit"></td></form>
<td width="25%" height="2" bgcolor="FFFFFF">

<input type="submit" value="我要购买" name="Submit" onclick="experience()"></td>
</tr>
</table>
</center>
</div><br>
</td>
</tr>


<tr>
<td width="100%" height="97"><div align="center">
<center>
<table border="0" cellpadding="3" cellspacing="1" width="100%" class=a2>
<tr class=a1>
<td width="50%" height="5" align="center" colspan="2"><b>
鲜</b>　<b>花</b></td>
<td width="50%" height="5" align="center" colspan="2"><b>
鸡</b>　<b>蛋</b></td>
</tr>
<tr>
<td width="15%" height="4" align="center" rowspan="2" bgcolor="FFFFFF"><FONT color=cccccc>
<IMG src="images/plus/flowers.gif" border=1></FONT></td>
<td width="25%" height="2" bgcolor="FFFFFF">名称：鲜花<br>
价格：50 <b>金币</b><br>
功能：增加对方的经验值</td>
<td width="15%" height="4" align="center" rowspan="2" bgcolor="FFFFFF"><FONT color=cccccc>
<IMG src="images/plus/egg.gif" border=1></FONT></td>
<td width="25%" height="2" bgcolor="FFFFFF">名称：鸡蛋<br>
价格：50 <b>金币</b><br>
功能：降低对方的经验值</td>
</tr>
<tr>
<td width="25%" height="2" bgcolor="FFFFFF">
<input type="submit" value="我要购买" name="Submit1" onclick="flowers()"></td>
<td width="25%" height="2" bgcolor="FFFFFF">
<input type="submit" value="我要购买" name="Submit2"  onclick="egg()"></td>
</tr>
</table>
</center>
</div><br>
</td>
</tr>


<tr>
<td width="100%" height="97"><div align="center">
<center>
<table border="0" cellpadding="3" cellspacing="1" width="100%" class=a2>
<tr class=a1>
<td width="50%" height="5" align="center" colspan="2"><b>
讯息群发器</b></td>
<td width="50%" height="5" align="center" colspan="2"><b>贵宾凭证</b></td>
</tr>
<tr>
<td width="15%" height="4" align="center" rowspan="2" bgcolor="FFFFFF"><FONT color=cccccc>
<IMG src="images/plus/colony.gif" border=1></FONT></td>
<td width="25%" height="2" bgcolor="FFFFFF">名称：讯息群发器<br>
价格：在线会员数 × 10<b>金币</b><br>
功能：群发讯息给在线会员</td>
<td width="15%" height="4" align="center" rowspan="2" bgcolor="FFFFFF">
<FONT color=cccccc>
<IMG src="images/plus/vipuser.gif" border=1></FONT><td width="25%" height="2" bgcolor="FFFFFF">
名称：贵宾凭证<br>
价格：10000 <b>金币</b><br>
功能：升级为贵宾会员<br>
（经验必须达到200000以上者才可购买）</tr>
<tr>
<td width="25%" height="2" bgcolor="FFFFFF">
<input type="submit" value="我要购买" name="Submit3" onclick="colony()"></td>
<td width="25%" height="2" bgcolor="FFFFFF">

<input type="submit" value="我要购买" name="Submit4" onclick="VIPUser()"></tr>
</table>
</center>
</div>
</td>
</tr>



</table>




<%htmlend%>