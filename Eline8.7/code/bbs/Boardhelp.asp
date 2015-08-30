<!--#include file=conn.asp-->
<!-- #include file=inc/const.asp -->
<%if boardid=0 then
	stats="论坛总帮助"
	call nav()
	call head_var(2,0,"","")
else
	stats="版面帮助"
	call nav()
	call head_var(1,BoardDepth,0,0)
end if
call main()
call footer()

sub main()%>
	<table cellspacing=1 cellpadding=3 align=center class=tableborder1>
		<tr> 
			<td class=tablebody2 width=100% align=center><A HREF=#point>积分设置</A> | <A HREF=#grade>等级设置</A> | <A HREF=#ubb>UBB语法</A> | <A HREF=#ubb1>UBB设置</A> | <A HREF=#price>价格设置</A></td>
		</tr>
	  <tr> 
	    <th align=left>&nbsp;&nbsp;A. <A name=point>积分设置</A></th>
	  </tr> 
		<tr> 
			<td width=100% class=tablebody1>&nbsp;&nbsp;该论坛注册、登录、发帖、回帖、加入精华、删除帖子等操作对用户分值的影响，版主可根据用户发帖表现自行增减以下默认值，总版主可对用户威望进行直接操作：<BR>
				<BR>&nbsp;&nbsp;（一）<B>金钱</B><BR>&nbsp;&nbsp;注册金钱数：<font color=<%=Forum_body(8)%>><%=Forum_user(0)%></font>&nbsp;登录增加金钱：<font color=<%=Forum_body(8)%>><%=Forum_user(4)%></font>&nbsp;发帖增加金钱：<font color=<%=Forum_body(8)%>><%=Forum_user(1)%></font>&nbsp;跟帖增加金钱：<font color=<%=Forum_body(8)%>><%=Forum_user(2)%></font>&nbsp;精华增加金钱：<font color=<%=Forum_body(8)%>><%=Forum_user(15)%></font>&nbsp;删帖减少金钱：<font color=<%=Forum_body(8)%>><%=Forum_user(3)%></font>&nbsp;投票增加金钱：<font color=<%=Forum_body(8)%>><%=Forum_user(40)%></font><BR>
				<BR>&nbsp;&nbsp;（二）<B>经验</B><BR>&nbsp;&nbsp;注册经验数：<font color=<%=Forum_body(8)%>><%=Forum_user(5)%></font>&nbsp;登录增加经验：<font color=<%=Forum_body(8)%>><%=Forum_user(9)%></font>&nbsp;发帖增加经验：<font color=<%=Forum_body(8)%>><%=Forum_user(6)%></font>&nbsp;跟帖增加经验：<font color=<%=Forum_body(8)%>><%=Forum_user(7)%></font>&nbsp;精华增加经验：<font color=<%=Forum_body(8)%>><%=Forum_user(17)%></font>&nbsp;删帖减少经验：<font color=<%=Forum_body(8)%>><%=Forum_user(8)%></font>&nbsp;投票增加经验值：<font color=<%=Forum_body(8)%>><%=Forum_user(41)%></font><BR>
				<BR>&nbsp;&nbsp;（三）<B>魅力</B><BR>&nbsp;&nbsp;注册魅力数：<font color=<%=Forum_body(8)%>><%=Forum_user(10)%></font>&nbsp;登录增加魅力：<font color=<%=Forum_body(8)%>><%=Forum_user(14)%></font>&nbsp;发帖增加魅力：<font color=<%=Forum_body(8)%>><%=Forum_user(11)%></font>&nbsp;跟帖增加魅力：<font color=<%=Forum_body(8)%>><%=Forum_user(12)%></font>&nbsp;精华增加魅力：<font color=<%=Forum_body(8)%>><%=Forum_user(16)%></font>&nbsp;删帖减少魅力：<font color=<%=Forum_body(8)%>><%=Forum_user(13)%></font>&nbsp;投票增加魅力值：<font color=<%=Forum_body(8)%>><%=Forum_user(42)%></font><BR>
			</td>
		</tr>
	  <tr> 
	    <th align=left>&nbsp;&nbsp;B. <A name=grade>等级设置</A></th>
	  </tr> 
	  <tr>
	  	<td class=tablebody1>
	  		<table cellspacing=1 cellpadding=6 border=0>
				  <tr> 
				    <td colspan=2>以下为该论坛相应等级所需文章，以及相应的等级图片：</td>
				  </tr>
					<%set rs=conn.execute("select * from usertitle where not Minarticle=-1 order by MinArticle")
					do while not rs.eof
						response.write "<tr><td valign=middle>升级到 " & rs("usertitle") & " 需要 " & rs("MinArticle") & " 的文章 等级标志为 </td><td valign=middle><img src=pic/"&rs("titlepic")&"></td></tr>" 
						rs.movenext
					loop
					rs.close
					set rs=nothing%>
					<tr><td valign=middle>贵宾  为管理员设定，可以进入特定版面。</td><td valign=middle><img src=pic/<%=conn.execute("select top 1 TitlePic from usertitle where usertitleid=18")(0)%>></td></tr>
					<tr><td valign=middle>版主  为管理员设定，可以对其管理的论坛帖子进行管理。</td><td valign=middle><img src=pic/<%=conn.execute("select top 1 TitlePic from usertitle where usertitleid=19")(0)%>></td></tr>
					<tr><td valign=middle>超级版主  为管理员设定，可以对所有论坛帖子进行管理。</td><td valign=middle><img src=pic/<%=conn.execute("select top 1 TitlePic from usertitle where usertitleid=20")(0)%>></td></tr>
					<tr><td valign=middle>管理员  ，拥有论坛全部权限。</td><td valign=middle><img src=pic/<%=conn.execute("select top 1 TitlePic from usertitle where usertitleid=21")(0)%>></td></tr>
				</table>
			</td>
	  </tr> 
	  <tr> 
	    <th align=left>&nbsp;&nbsp;C. <A name=ubb>UBB语法</A></th>
	  </tr> 
	  <tr> 
	    <td class=tablebody1>
				<p>&nbsp;&nbsp;论坛可以由管理员设置是否支持UBB标签，UBB标签就是不允许使用HTML语法的情况下，通过论坛的特殊转换程序，以至可以支持少量常用的、无危害性的HTML效果显示。以下为具体使用说明：
				<p>&nbsp;&nbsp;<font color=red>[B]</font><b>文字</b><font color=red>[/B]</font>：在文字的位置可以任意加入您需要的字符，显示为粗体效果。
				<p>&nbsp;&nbsp;<font color=red>[I]</font><i>文字</i><font color=red>[/I]</font>：在文字的位置可以任意加入您需要的字符，显示为斜体效果。
				<p>&nbsp;&nbsp;<font color=red>[U]</font><u>文字</u><font color=red>[/U]</font>：在文字的位置可以任意加入您需要的字符，显示为下划线效果。
				<p>&nbsp;&nbsp;<font color=red>[align=center]</font>文字<font color=red>[/align]</font>：在文字的位置可以任意加入您需要的字符，center位置center表示居中，left表示居左，right表示居右。
				<p>&nbsp;&nbsp;<font color=red>[URL]</font><A HREF=HTTP://BBS.happyjh.com>HTTP://BBS.happyjh.com</A><font color=red>[/URL]</font>
				<P>&nbsp;&nbsp;<font color=red>[URL=HTTP://BBS.happyjh.com]</font><A HREF=HTTP://BBS.happyjh.com>E线论坛</A><font color=red>[/URL]</font>：有两种方法可以加入超级连接，可以连接具体地址或者文字连接。
				<p>&nbsp;&nbsp;<font color=red>[EMAIL]</font><A HREF=mailto:Eline_Email@etang.com>Eline_Email@etang.com</A><font color=red>[/EMAIL]</font>
				<P>&nbsp;&nbsp;<font color=red>[EMAIL=MAILTO:Eline_Email@etang.com]</font><A HREF=mailto:Eline_Email@etang.com>回首当年</A><font color=red>[/EMAIL]</font>：有两种方法可以加入邮件连接，可以连接具体地址或者文字连接。
				<P>&nbsp;&nbsp;<font color=red>[img]</font><img src=http://zhzx.jjedu.org/logo.gif><font color=red>[/img]</font>：在标签的中间插入图片地址可以实现插图效果。
				<P>&nbsp;&nbsp;<font color=red>[flash]</font>Flash连接地址<font color=red>[/Flash]</font>：在标签的中间插入Flash图片地址可以实现插入Flash。
				<P>&nbsp;&nbsp;<font color=red>[code]</font>文字<font color=red>[/code]</font>：在标签中写入文字可实现html中编号效果。
				<P>&nbsp;&nbsp;<font color=red>[quote]</font>引用<font color=red>[/quote]</font>：在标签的中间插入文字可以实现HTMl中引用文字效果。
				<P>&nbsp;&nbsp;<font color=red>[list]</font>文字<font color=red>[/list]</font> <font color=red>[list=a]</font>文字<font color=red>[/list]</font>  <font color=red>[list=1]</font>文字<font color=red>[/list]</font>：更改list属性标签，实现HTML目录效果。
				<P>&nbsp;&nbsp;<font color=red>[fly]</font>文字<font color=red>[/fly]</font>：在标签的中间插入文字可以实现文字飞翔效果，类似跑马灯。
				<P>&nbsp;&nbsp;<font color=red>[move]</font>文字<font color=red>[/move]</font>：在标签的中间插入文字可以实现文字移动效果，为来回飘动。
				<P>&nbsp;&nbsp;<font color=red>[glow=255,red,2]</font>文字<font color=red>[/glow]</font>：在标签的中间插入文字可以实现文字发光特效，glow内属性依次为宽度、颜色和边界大小。
				<P>&nbsp;&nbsp;<font color=red>[shadow=255,red,2]</font>文字<font color=red>[/shadow]</font>：在标签的中间插入文字可以实现文字阴影特效，shadow内属性依次为宽度、颜色和边界大小。
				<P>&nbsp;&nbsp;<font color=red>[color=颜色代码]</font>文字<font color=red>[/color]</font>：输入您的颜色代码，在标签的中间插入文字可以实现文字颜色改变。
				<P>&nbsp;&nbsp;<font color=red>[size=数字]</font>文字<font color=red>[/size]</font>：输入您的字体大小，在标签的中间插入文字可以实现文字大小改变。
				<P>&nbsp;&nbsp;<font color=red>[face=字体]</font>文字<font color=red>[/face]</font>：输入您需要的字体，在标签的中间插入文字可以实现文字字体转换。
				<P>&nbsp;&nbsp;<font color=red>[DIR=500,350]</font>http://<font color=red>[/DIR]</font>：为插入shockwave格式文件，中间的数字为宽度和长度
				<P>&nbsp;&nbsp;<font color=red>[RM=500,350]</font>http://<font color=red>[/RM]</font>：为插入realplayer格式的rm文件，中间的数字为宽度和长度
				<P>&nbsp;&nbsp;<font color=red>[MP=500,350]</font>http://<font color=red>[/MP]</font>：为插入为midia player格式的文件，中间的数字为宽度和长度
				<P>&nbsp;&nbsp;<font color=red>[QT=500,350]</font>http://<font color=red>[/QT]</font>：为插入为Quick time格式的文件，中间的数字为宽度和长度
			</td>
	  </tr> 
	  <tr> 
	    <th align=left>&nbsp;&nbsp;D. <A name=ubb1>UBB设置</A></th>
	  </tr> 
	  <tr> 
	    <td class=tablebody1>&nbsp;&nbsp; 下面为本论坛的UBB语法设置，通过这些设置，您可以知道在本版面发言中有哪些语句是不能使用的，这里还包括了控制用户签名里使用的UBB选项。<BR><%
	    	if boardid>0 then
	    		%>&nbsp;&nbsp;<B>用户发帖</B>：<ul><!--#include file="z_Contentflag.asp"--><BR><BR></ul>&nbsp;&nbsp;<B>特殊内容</B><BR>
	    		<ul>
					<li><%=iif(Board_Setting(54),"会员帖：可用","会员帖：不可用")%>
					<li><%=iif(Board_Setting(59),"门派帖：可用","门派帖：不可用")%>
					<li><%=iif(Board_Setting(58),"年龄帖：可用","年龄帖：不可用")%>
					<li><%=iif(Board_Setting(56),"性别帖：可用","性别帖：不可用")%>
					<li><%=iif(Board_Setting(57),"高级帖：可用","高级帖：不可用")%>
					<li><%=iif(Board_Setting(53),"秘密帖：可用","秘密帖：不可用")%>
					<li><%=iif(Board_Setting(55),"定人帖：可用","定人帖：不可用")%>
					<li><%=iif(Board_Setting(10),"金钱帖：可用","金钱帖：不可用")%>
					<li><%=iif(Board_Setting(11),"积分帖：可用","积分帖：不可用")%>
					<li><%=iif(Board_Setting(12),"魅力帖：可用","魅力帖：不可用")%>
					<li><%=iif(Board_Setting(13),"威望帖：可用","威望帖：不可用")%>
					<li><%=iif(Board_Setting(14),"文章帖：可用","文章帖：不可用")%>
					<li><%=iif(Board_Setting(60),"显示帖：可用","显示帖：不可用")%>
					<li><%=iif(Board_Setting(61),"隐藏帖：可用","隐藏帖：不可用")%>
					<li><%=iif(Board_Setting(15),"回复帖：可用","回复帖：不可用")%>
					<li><%=iif(Board_Setting(23),"出售帖：可用","出售帖：不可用")%>
					</ul>
				<%end if
				%><BR>&nbsp;&nbsp;<B>用户签名</B>：
				<ul>
				<li>HTML标签： <%=iif(Forum_Setting(66),"可用","不可用")%>
				<li>UBB标签： <%=iif(Forum_Setting(65),"可用","不可用")%>
				<li>帖图标签(包括图片、flash、多媒体)： <%=iif(Forum_Setting(67),"可用","不可用")%>
				</ul>说明：这里html标签指是否允许使用html语法，帖图和flash以及表情字符转换都属于UBB语法内容，其使用方法可查看UBB语法
			</td>
	  </tr> 
	  <tr> 
	    <th align=left>&nbsp;&nbsp;E. <A name=price>价格设置</A></th>
	  </tr> 
	  <tr> 
	    <td class=tablebody1>
				<p>&nbsp;&nbsp;普通会员给全体会员点歌： <font color=<%=forum_body(8)%>><%=forum_user(18)%></font> 元/次
				<p>&nbsp;&nbsp;普通会员给单一会员点歌： <font color=<%=forum_body(8)%>><%=forum_user(19)%></font> 元/人
				<p>&nbsp;&nbsp;VIP 会员给全体会员点歌： <font color=<%=forum_body(8)%>><%=forum_user(20)%></font> 元/次
				<p>&nbsp;&nbsp;VIP 会员给单一会员点歌： <font color=<%=forum_body(8)%>><%=forum_user(21)%></font> 元/人
				<p>&nbsp;&nbsp;普通会员修改头像： <font color=<%=forum_body(8)%>><%=forum_user(22)%></font> 元/次
				<p>&nbsp;&nbsp;VIP 会员修改头像： <font color=<%=forum_body(8)%>><%=forum_user(23)%></font> 元/次
				<p>&nbsp;&nbsp;普通会员修改签名： <font color=<%=forum_body(8)%>><%=forum_user(24)%></font> 元/次
				<p>&nbsp;&nbsp;VIP 会员修改签名： <font color=<%=forum_body(8)%>><%=forum_user(25)%></font> 元/次
				<p>&nbsp;&nbsp;普通会员发送短信： <font color=<%=forum_body(8)%>><%=forum_user(26)%></font> 元/人
				<p>&nbsp;&nbsp;VIP 会员发送短信： <font color=<%=forum_body(8)%>><%=forum_user(27)%></font> 元/人
				<p>&nbsp;&nbsp;普通会员回复后短信通知： <font color=<%=forum_body(8)%>><%=forum_user(28)%></font> 元/帖
				<p>&nbsp;&nbsp;VIP 会员回复后短信通知： <font color=<%=forum_body(8)%>><%=forum_user(29)%></font> 元/帖
				<p>&nbsp;&nbsp;普通会员回复后邮件通知： <font color=<%=forum_body(8)%>><%=forum_user(30)%></font> 元/帖
				<p>&nbsp;&nbsp;VIP 会员回复后邮件通知： <font color=<%=forum_body(8)%>><%=forum_user(31)%></font> 元/帖
				<p>&nbsp;&nbsp;普通会员短信引起注意： <font color=<%=forum_body(8)%>><%=forum_user(32)%></font> 元/帖
				<p>&nbsp;&nbsp;VIP 会员短信引起注意： <font color=<%=forum_body(8)%>><%=forum_user(33)%></font> 元/帖
				<p>&nbsp;&nbsp;普通会员上传头像： <font color=<%=forum_body(8)%>><%=forum_user(34)%></font> 元/KB
				<p>&nbsp;&nbsp;VIP 会员上传头像： <font color=<%=forum_body(8)%>><%=forum_user(35)%></font> 元/KB
				<p>&nbsp;&nbsp;普通会员上传文件： <font color=<%=forum_body(8)%>><%=forum_user(36)%></font> 元/KB
				<p>&nbsp;&nbsp;VIP 会员上传文件： <font color=<%=forum_body(8)%>><%=forum_user(37)%></font> 元/KB
				<p>&nbsp;&nbsp;普通会员上传照片： <font color=<%=forum_body(8)%>><%=forum_user(38)%></font> 元/KB
				<p>&nbsp;&nbsp;VIP 会员上传照片： <font color=<%=forum_body(8)%>><%=forum_user(39)%></font> 元/KB
			</td>
		</tr>
	</table>
	<%call activeonline()
end sub
%>