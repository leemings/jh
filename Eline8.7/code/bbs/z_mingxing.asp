<%
'明星之水王首页显示
sub mingxing()
	'头部文件
	response.write("<tr>")
	dim sql,rs,F_or_M,username,article,mingxing_1_type,mingxing_2_type

'显示第一列的头像
select case forum_setting(17)
case 1
	mingxing_1_type="今日发帖"
	sql="select top 1 username,count(username) from "&NowUseBbs&" where DateAndTime>Date() group by username order by count(username) desc"
case 2
	mingxing_1_type="本周发帖"
	sql="select top 1 username,count(username) from "&NowUseBbs&" where DateAndTime>Date()-weekday(date(),2) group by username order by count(username) desc"
case 3
	mingxing_1_type="本月发帖"
	sql="select top 1 username,count(username) from "&NowUseBbs&" where year(DateAndTime)=year(date()) and Month(DateAndTime)=month(Date()) group by username order by count(username) desc"
case 4
	mingxing_1_type="今年发帖"
	sql="select top 1 username,count(username) from "&NowUseBbs&" where lockuser=0 and yyear(DateAndTime)=year(Date()) group by username order by count(username) desc"
case 5
	mingxing_1_type="总计发帖"
	sql="select top 1 username,article from [user] where lockuser=0 order by article desc"
case 6
	mingxing_1_type="最佳男发帖"
	sql="select top 1 username,article from [user] where sex='1' and UserGroupID>3 order by Article desc"
case 7
	mingxing_1_type="最佳女发帖"
	sql="select top 1 username,article from [user] where sex='0' and UserGroupID>3 order by Article desc"
case 8
	mingxing_1_type="今日得分"
	sql="select top 1 username,sum(score) from "&NowUseBbs&" where not isnull(score) and scoretime>Date() group by username order by sum(score) desc"
case else
	mingxing_1_type="今日发帖"
	sql="select top 1 username,count(username) from "&NowUseBbs&" b where DateAndTime>Date() group by username order by count(username) desc"
end select
set rs=conn.execute(sql)
if rs.eof then
	response.write("<TD vAlign=middle align=middle width='5%' class=tablebody2> <IMG height=27 src='pic/FeMale.gif' width=25 align=absMiddle>")
	response.write("</TD>")
	response.write("<TD vAlign=top align=middle width='45%' class=tablebody1>")
	response.write("<table border=0 width='100%' cellspacing=0 cellpadding=0>")
	response.write("<tr>")
	response.write("<td width=44% valign=top> <strong>")
	response.write("<font color=#ff0000>"&mingxing_1_type&"</font><font color=#DA9136>状元</font>")
	response.write("</strong> <font color=#9999ff>形象-=></font>")
	response.write("<p align=left>姓　　名：<font color=#ff0000>等你来改写</font> <br>")
	'======================================================================================
	response.write "社区等级："
	response.write "<br><font color=red>"&mingxing_1_type & "</font>："
	response.write "<br>个人现金："
	response.write "<br>个人魅力："
	response.write "<br>社区经验："
	'========================================================================================
	response.write("</td><td width=50% valign=middle align=center>&nbsp;</td></tr></table></TD>")
else
	username=rs(0)
	article=rs(1)
	'================================================================
	sql="select userclass,userWealth,useremail,face,width,height,sex,userEP,userCP from [user] where username='"&username&"'"
	'================================================================
	set rs=conn.execute(sql)
	if rs(6)="1" then
		F_or_M="Male.gif"
	else
		F_or_M="FeMale.gif"
	end if
	response.write("<TD vAlign=middle align=middle width='5%' class=tablebody2> <IMG height=27 src='pic/"&F_or_M&"' width=25 align=absMiddle>")
	response.write("</TD>")
	response.write("<TD vAlign=top align=middle width='45%' class=tablebody1>")
	response.write("<table border=0 width='100%' cellspacing=0 cellpadding=0>")
	response.write("<tr>")
	response.write("<td width=44% valign=top> <strong>")
	response.write("<font color=#ff0000>"&mingxing_1_type&"</font><font color=#DA9136>状元</font>")
	response.write("</strong> <font color=#9999ff>形象-=></font>")
	response.write("<p align=left>姓　　名：<a href=dispuser.asp?name="&username&" target=_blank><font color=blue>"&username&"</font></a> <br>")
	'==============================================================
	response.write "社区等级：" & rs(0) & "<br>"
	response.write "<font color=red>"&mingxing_1_type & "</font>：<font color=red>" & Article & "</font>"
	response.write "<br>个人现金："&rs(1)
	response.write "<br>个人魅力："&rs(8)
	response.write "<br>社区经验："&rs(7)
	'==============================================================	
	response.write("</td>")
	response.write "<td width=50% valign=middle align=center>"
	call ShowUserVisual(UserName,99,"tablebody1",(cint(left(forum_setting(72),1))=1))
	response.write "</td></table></TD>"
end if
	
'显示第二列的头像
select case forum_setting(18)
case 1
	mingxing_2_type="今日发帖"
	sql="select top 2 username,count(username) from "&NowUseBbs&" where DateAndTime>Date() group by username order by count(username) desc"
case 2
	mingxing_2_type="本周发帖"
	sql="select top 2 username,count(username) from "&NowUseBbs&" where DateAndTime>Date()-weekday(date(),2) group by username order by count(username) desc"
case 3
	mingxing_2_type="本月发帖"
	sql="select top 2 username,count(username) from "&NowUseBbs&" where year(DateAndTime)=year(date()) and Month(DateAndTime)=month(Date()) group by username order by count(username) desc"
case 4
	mingxing_2_type="今年发帖"
	sql="select top 2 username,count(username) from "&NowUseBbs&" where year(DateAndTime)=year(Date()) group by username order by count(username) desc"
case 5
	mingxing_2_type="总计发帖"
	sql="select top 2 username,article from [user] where lockuser=0 order by article desc"
case 6
	mingxing_2_type="最佳男发帖"
	sql="select top 2 username,article from [user] where sex='1' and UserGroupID>3 order by Article desc"
case 7
	mingxing_2_type="最佳女发帖"
	sql="select top 2 username,article from [user] where sex='0' and UserGroupID>3 order by Article desc"
case 8
	mingxing_2_type="今日得分"
	sql="select top 2 username,sum(score) from "&NowUseBbs&" where not isnull(score) and scoretime>Date() group by username order by sum(score) desc"
case else
	mingxing_2_type="今日发帖"
	sql="select top 2 username,count(username) from "&NowUseBbs&" where DateAndTime>Date() group by username order by count(username) desc"
end select
set rs=conn.execute(sql)
if (not rs.eof) and forum_setting(17)=forum_setting(18) then
	rs.movenext
end if	
if rs.eof then
	response.write("<TD vAlign=middle align=middle width='5%' class=tablebody2> <IMG height=27 src='pic/FeMale.gif' width=25 align=absMiddle>")
	response.write("</TD>")
	response.write("<TD vAlign=top align=middle width='45%' class=tablebody1>")
	response.write("<table border=0 width='100%' cellspacing=0 cellpadding=0>")
	response.write("<tr>")
	response.write("<td width=44% valign=top> <strong>")
	if forum_setting(17)=forum_setting(18) then
		response.write("<font color=#ff0000>"&mingxing_2_type&"</font><font color=#DA9136>榜眼</font>")
	else
		response.write("<font color=#ff0000>"&mingxing_2_type&"</font><font color=#DA9136>状元</font>")
	end if
	response.write("</strong> <font color=#9999ff>形象-=></font>")
	response.write("<p align=left>姓　　名：<font color=#ff0000>等你来改写</font> <br>")
	'======================================================================================
	response.write "社区等级："
	response.write "<br><font color=red>"&mingxing_2_type & "</font>："
	response.write "<br>个人现金："
	response.write "<br>个人魅力："
	response.write "<br>社区经验："
	'========================================================================================
	response.write("</td><td width=50% valign=middle align=center>&nbsp;</td></tr></table></TD>")
else
	username=rs(0)
	article=rs(1)
	'================================================================
	sql="select userclass,userWealth,useremail,face,width,height,sex,userEP,userCP from [user] where username='"&username&"'"
	'================================================================
	set rs=conn.execute(sql)
	if rs(6)="1" then
		F_or_M="Male.gif"
	else
		F_or_M="FeMale.gif"
	end if
	response.write("<TD vAlign=middle align=middle width='5%' class=tablebody2> <IMG height=27 src='pic/"&F_or_M&"' width=25 align=absMiddle>")
	response.write("</TD>")
	response.write("<TD vAlign=top align=middle width='45%' class=tablebody1>")
	response.write("<table border=0 width='100%' cellspacing=0 cellpadding=0>")
	response.write("<tr valign=top>")
	response.write("<td width=44% > <strong>")
	if forum_setting(17)=forum_setting(18) then
		response.write("<font color=#ff0000>"&mingxing_2_type&"</font><font color=#DA9136>榜眼</font>")
	else
		response.write("<font color=#ff0000>"&mingxing_2_type&"</font><font color=#DA9136>状元</font>")
	end if
	response.write("</strong> <font color=#9999ff>形象-=></font>")
	response.write("<p align=left>姓　　名：<a href=dispuser.asp?name="&username&" target=_blank><font color=blue>"&username&"</font></a> <br>")
	'==============================================================
	response.write "社区等级：" & rs(0) & "<br>"
	response.write "<font color=red>"&mingxing_2_type & "</font>：<font color=red>" & Article & "</font>"
	response.write "<br>个人现金："&rs(1)
	response.write "<br>个人魅力："&rs(8)
	response.write "<br>社区经验："&rs(7)
	'==============================================================	
	response.write("</td>")
	response.write "<td width=50% align=center valign=middle>"
	call ShowUserVisual(UserName,99,"tablebody1",(cint(left(forum_setting(72),1))=1))
	response.write "</td></table></TD>"
end if
'尾部文件
response.write("</TR>")
rs.close
set rs=nothing
end sub 

function mingxingtop(topmax)
	dim sql,rs
	sql="select top "&topmax&" username,count(username) from bbs1 where DateAndTime>Date() group by username order by count(username) desc"
	set rs=conn.execute(sql)
	if rs.eof then
		mingxingtop="今日暂时无人发贴" 
	else
		mingxingtop=""
		dim c
		c=0
		do while not rs.eof
			mingxingtop=mingxingtop&"<font face=Wingdings color=blue>v</font> <a href=dispuser.asp?name="&htmlencode(rs(0))&" target=_blank)>"&rs(0)&"</a>[<font color=red>"&rs(1)&"</font>]&nbsp;&nbsp;"
			rs.movenext
			c=c+1
			if c=topmax then exit do
		loop
	end if
end function
'明星之水王首页显示代码结束
%>