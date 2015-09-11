<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
name=LCase(trim(Request.QueryString("name")))
if name="大家" or name="" then name=aqjh_name
Set conn = Server.CreateObject("ADODB.CONNECTION")
Set rs = Server.CreateObject("ADODB.RecordSet")
conn.Open (Application("aqjh_usermdb"))
rs.Open "select house.*,用户.姓名,用户.性别,用户.配偶,用户.门派,用户.身份,用户.名单头像,用户.等级,用户.会员等级,用户.银两 from house inner join 用户 on house.h_拥有者=用户.姓名 where house.h_拥有者='"& name &"'", conn, 1, 1
if rs.Eof and rs.Bof then
	rs.Close
	if name<>aqjh_name then
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Write "<script language=JavaScript>{alert('提示:此账号交没有购买房屋,不能查看！');}</script>"
		Response.End
	end if
	rs.Open "select house.*,用户.姓名,用户.性别,用户.门派,用户.配偶,用户.身份,用户.名单头像,用户.等级,用户.会员等级,用户.银两 from house inner join 用户 on house.h_拥有者=用户.姓名 where house.h_拥有者2='"& name &"'", conn, 1, 1
	if rs.Eof and rs.Bof then
		rs.close
		Set rs = Nothing
		conn.Close
		Set conn = Nothing
		Response.Write "<script language=JavaScript>{alert('提示:您还没有购买房屋,不能进行操作,请去购买!');top.window.close();}</script>"
		Response.End
	end if
	session("myroom")="h_拥有者2"
	info=name & "回到了与["& rs("姓名") &"]的家"
else
	if name=aqjh_name then
		session("myroom")="h_拥有者"
		info=name & "回到了自己的家"
	elseif rs("h_拥有者2")=aqjh_name then
		name=aqjh_name
		session("myroom")="h_拥有者2"
		info=name & "回到了与["& rs("姓名") &"]的家"
	else
		session("myroom")="h_拥有者"
		info=name & "在参观["& rs("姓名") &"]的家"
	end if
end if
if name<>aqjh_name and rs("h_参观")=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示:此房屋已经上锁禁止参观!');top.window.close();}</script>"
	Response.End
end if
mywife=rs("配偶")
%>
<HTML><HEAD><TITLE><%=info%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/efeelingsbbs.css" rel=stylesheet>
<SCRIPT language=JavaScript>
function changeImage(an_image, newSource) {var theimg=eval("window.document."+an_image);theimg.src = newSource;}
function shLayers(n){if(n.style.visibility=="visible"){n.style.visibility="hidden";}else if(n.style.visibility=="hidden"){n.style.visibility="visible";}}
function showinfo(){shLayers(photo);shLayers(xinxi);}
</SCRIPT>
<STYLE>.gray {	FILTER: gray}</STYLE>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-- 显示人物背景图 -->
<DIV id=Layer1 style="Z-INDEX: 2; LEFT: 3px; WIDTH: 395px; POSITION: absolute; TOP: 43px; HEIGHT: 283px"><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0><TBODY><TR><TD width=245><IMG height=373 src="images/outdoorb.jpg" width=245></TD><TD vAlign=top width=150><IMG height=104 src="images/outdoorbg01.gif" width=150><IMG height=122 src="images/outdoorbg02.gif" width=150><IMG height=147 src="images/outdoorbg03.gif" width=150></TD></TR></TBODY></TABLE></DIV>
<%if name=session("aqjh_name") or name=session("aqjh_name") then%>
<!-- 主人留言 -->
<A href="#" onClick="window.open('../letter/write.asp')"><IMG onmouseover="changeImage('Image1','images/outdoorb1f.gif')" style="Z-INDEX: 7; LEFT: 260px; POSITION: absolute; TOP: 130px" onmouseout="changeImage('Image1','images/outdoorb1.gif')" src="images/outdoorb1.gif" border=none name=Image1></A>
<!-- 房间 -->
<A onClick="window.open('myhome.asp?name=<%=name%>','myhome','scrollbars=no,resizable=no,width=630,height=450,menubar=no,top=100,left=200')" href="#"><IMG onmouseover="changeImage('Image4','images/outdoorb4f.gif')" style="Z-INDEX: 6; LEFT: 255px; POSITION: absolute; TOP: 302px" onmouseout="changeImage('Image4','images/outdoorb4.gif')" src="images/outdoorb4.gif" border=none name=Image4></A>
<%else%>
<!-- 主人留言 -->
<A href="#" onClick="window.open('../letter/write.asp','mess','scrollbars=yes,resizable=yes,width=500,height=400')"><IMG onmouseover="changeImage('Image1','images/outdoorb1f.gif')" style="Z-INDEX: 7; LEFT: 260px; POSITION: absolute; TOP: 130px" onmouseout="changeImage('Image1','images/outdoorb1.gif')" src="images/outdoorb1.gif" border=none name=Image1></A>
<!-- 房间 -->
<A href="javascript:alert('小屋内部不是配偶禁止进入!');"><IMG onmouseover="changeImage('Image4','images/outdoorb4f.gif')" style="Z-INDEX: 6; LEFT: 255px; POSITION: absolute; TOP: 302px" onmouseout="changeImage('Image4','images/outdoorb4.gif')" src="images/outdoorb4.gif" border=none name=Image4></A>
<%end if%>
<!-- 主人信息 -->
<A href="#" onClick="window.open('SEEME.ASP')"><IMG onmouseover="changeImage('Image2','images/outdoorb2f.gif')" style="Z-INDEX: 8; LEFT: 277px; POSITION: absolute; TOP: 200px" onmouseout="changeImage('Image2','images/outdoorb2.gif')" src="images/outdoorb2.gif" border=none name=Image2></A>
<!-- 花园 -->
<A onclick="window.open('../garden/hua.asp?name=<%=name%>','garden','menubar=no,scrollbars=no,width=640,height=420')" href="#"><IMG onmouseover="changeImage('Image3','images/outdoorb3f.gif')" style="Z-INDEX: 5; LEFT: 266px; POSITION: absolute; TOP: 250px" onmouseout="changeImage('Image3','images/outdoorb3.gif')" src="images/outdoorb3.gif" border=none name=Image3></A>
<!-- 房子类型图片 -->
<IMG style="Z-INDEX: 9; LEFT: 70px; POSITION: absolute; TOP: 330px" src="images/outdoorbg04.gif">
<IMG style="Z-INDEX: 10; LEFT: 271px; WIDTH: 112px; POSITION: absolute; TOP: 70px; HEIGHT: 49px" src="images/outdoorbg05.gif">
<!-- 心灵感应 -->
<DIV id=Layer11 style="Z-INDEX: 13; LEFT: 20px; WIDTH: 29px; POSITION: absolute; TOP: 305px; HEIGHT: 30px"><A onclick="window.open('../mess.asp','心灵感应','width=201,height=152','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no')" href="#"><IMG height=30 src="images/nolttr.gif" width=30 border=0></A></DIV>
<DIV id=Layer12 style="Z-INDEX: 14; LEFT: 394px; WIDTH: 40px; POSITION: absolute; TOP: 363px; HEIGHT: 36px"><TABLE><TBODY><TR><TD></TD></TR></TBODY></TABLE></DIV>
<%
myname=rs("姓名")
mysex=rs("性别")
Response.Write "<IMG title=点我可以进入海天大地里的本村 style=""Z-INDEX: 1; LEFT: 299px; CURSOR: hand; POSITION: absolute; TOP: 35px""  src=""pic/hut"& rs("h_类型") & rs("h_级别") &".jpg"" name=hutcover width=""480"" height=""360"">"
'<!-- 所在地区 -->
Response.Write "<DIV id=Layer9 style=""Z-INDEX: 11; LEFT: 280px; OVERFLOW: hidden; WIDTH: 96px; POSITION: absolute; TOP: 87px; HEIGHT: 31px""><TABLE cellSpacing=0 cellPadding=0 width=""100%"" border=0><TBODY><TR><TD><FONT id=town color=#330000 size=2>"& rs("h_小区名") & rs("h_级别") &"级</FONT></TD></TR></TBODY></TABLE></DIV>"

'<!-- 显示用户名及性别 -->
Response.Write "<DIV id=Layer6 style=""Z-INDEX: 14; LEFT: 165px; OVERFLOW: hidden; WIDTH: 235px; POSITION: absolute; TOP: 56px; HEIGHT: 14px""><FONT id=hostinfo color=#330000>"& rs("姓名") &"&nbsp;性别:"& rs("性别") &"&nbsp;"
if rs("会员等级")>0 then 
	Response.Write rs("会员等级") &"会员&nbsp;"
else
	Response.Write "<a href=""../chat/hy.asp"">非会员</a>&nbsp;"
end if

Response.Write "</FONT></DIV>"

%>

</BODY></HTML>