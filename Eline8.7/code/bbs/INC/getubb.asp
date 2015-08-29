<SELECT onchange="if(this.options[this.selectedIndex].value!=''){showfont(this.options[this.selectedIndex].value);this.options[0].selected=true;}else {this.selectedIndex=0;}" name=font>
<option value="宋体" selected>字体类型</option>
<OPTION>========</OPTION>
<option value="宋体">宋体</option>
<option value="楷体_GB2312">楷体</option>
<option value="黑体">黑体</option>
<option value="隶书">隶书</option>
<OPTION value=Arial>Arial</OPTION> 
<OPTION value="Century Gothic">Century Gothic</OPTION> 
<OPTION value="Comic Sans MS">Comic Sans MS</OPTION>
<OPTION value="Courier New">Courier New</OPTION>
<OPTION value=Impact>Impact</OPTION>
<OPTION value=Tahoma>Tahoma</OPTION>
<OPTION value="Times New Roman">Times New Roman</OPTION>
<OPTION value="Script MT Bold">Script MT Bold</OPTION>
</SELECT>&nbsp;
<select name="size" onChange="if(this.options[this.selectedIndex].value!=''){showsize(this.options[this.selectedIndex].value);this.options[0].selected=true;}else {this.selectedIndex=0;}">
<option value="3" selected>字体大小</option>
<OPTION>=====</OPTION>
<option value="1">1</option>
<option value="2">2</option>
<option value="3">3</option>
<option value="4">4</option>
<option value="5">5</option>
<option value="6">6</option>
<option value="7">7</option>
</select><!--&nbsp;
<select name="dongzuo" onChange="insertsmilie(this.options[this.selectedIndex].value)">
<OPTION value="" selected>动作</OPTION>
<OPTION>===</OPTION>
<OPTION value="/无聊">无聊</OPTION>
<OPTION value="/负责">负责</OPTION>
<OPTION value="/生气">生气</OPTION>
<OPTION value="/招呼">招呼</OPTION>
<OPTION value="/鼓掌">鼓掌</OPTION>
<OPTION value="/等待">等待</OPTION>
<OPTION value="/反对">反对</OPTION>
<OPTION value="/唱歌">唱歌</OPTION>
<OPTION value="/不要">不要</OPTION>
<OPTION value="/去死">去死</OPTION>
<OPTION value="/狂笑">狂笑</OPTION>
<OPTION value="/痛哭">痛哭</OPTION>
<OPTION value="/道别">道别</OPTION>
<OPTION value="/跳舞">跳舞</OPTION>
<OPTION value="/失望">失望</OPTION>
<OPTION value="/浪漫">浪漫</OPTION>
<OPTION value="/害羞">害羞</OPTION>
<OPTION value="/比酷">比酷</OPTION>
<OPTION value="/救命">救命</OPTION>
<OPTION value="/找死">找死</OPTION>
<OPTION value="/狂妄">狂妄</OPTION>
<OPTION value="/拳击">拳击</OPTION>
<OPTION value="/我踩">我踩</OPTION>
<OPTION value="/饶命">饶命</OPTION>
<OPTION value="/我踢">我踢</OPTION>
<OPTION value="/傻笑">傻笑</OPTION>
<OPTION value="/臭美">臭美</OPTION>
<OPTION value="/灌水">灌水</OPTION>
<OPTION value="/变态">变态</OPTION>
<OPTION value="/拼酒">拼酒</OPTION>
<OPTION value="/深情">深情</OPTION>
<OPTION value="/恶心">恶心</OPTION>
<OPTION value="/高兴">高兴</OPTION>
<OPTION value="/怀疑">怀疑</OPTION>
<OPTION value="/拉勾">拉勾</OPTION>
<OPTION value="/经典">经典</OPTION>
<OPTION value="/同意">同意</OPTION>
<OPTION value="/KISS">KISS</OPTION>
<OPTION value="/生日">生日</OPTION>
<OPTION value="/晕倒">晕倒</OPTION>
<OPTION value="/气你">气你</OPTION>
<OPTION value="/错啦">错啦</OPTION>
<OPTION value="/加油">加油</OPTION>
<OPTION value="/恭喜">恭喜</OPTION>
<OPTION value="/简单">简单</OPTION>
<OPTION value="/鼓励">鼓励</OPTION>
<OPTION value="/过奖">过奖</OPTION>
<OPTION value="/原谅">原谅</OPTION>
<OPTION value="/感动">感动</OPTION>
<OPTION value="/道谢">道谢</OPTION>
<OPTION value="/摇头">摇头</OPTION>
<OPTION value="/点头">点头</OPTION>
<OPTION value="/拥抱">拥抱</OPTION>
<OPTION value="/ＯＫ">ＯＫ</OPTION>
<OPTION value="/欢迎">欢迎</OPTION>
</select>--><%
if boardid>0 or ucase(right(Request.ServerVariables("SCRIPT_NAME"),14))="Z_FASTPOST.ASP" then
if boardid=0 then
	redim Board_Setting(63)
	dim Referer
	Referer=(ucase(right(Request.ServerVariables("SCRIPT_NAME"),14))="Z_FASTPOST.ASP")
	Board_Setting(10)=Referer:Board_Setting(11)=Referer
	Board_Setting(12)=Referer:Board_Setting(13)=Referer
	Board_Setting(14)=Referer:Board_Setting(15)=Referer
	Board_Setting(23)=Referer:Board_Setting(53)=Referer
	Board_Setting(54)=Referer:Board_Setting(55)=Referer
	Board_Setting(56)=Referer:Board_Setting(57)=Referer
	Board_Setting(58)=Referer:Board_Setting(59)=Referer
	Board_Setting(60)=Referer:Board_Setting(61)=Referer
	Board_Setting(62)=Referer:Board_Setting(63)=Referer
end if%>&nbsp;
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;this.options[0].selected=true;}" name="teshutie">
<OPTION selected>特殊内容</OPTION>
<OPTION>=====</OPTION>
<%if Board_Setting(54) then%>
	<OPTION value="javascript:usermem()">会员帖</OPTION>
<%end if
if Board_Setting(59) then%>
	<OPTION value="javascript:usergroup()">门派帖</OPTION>
<%end if
if Board_Setting(58) then%>
	<OPTION value="javascript:userage()">年龄帖</OPTION>
<%end if
if Board_Setting(56) then%>
	<OPTION value="javascript:sex()">性别帖</OPTION>
<%end if
if Board_Setting(57) then%>
	<OPTION value="javascript:uservip()">高级帖</OPTION>
<%end if
if Board_Setting(53) then%>
	<OPTION value="javascript:sec()">秘密帖</OPTION>
<%end if
if Board_Setting(55) then%>
	<OPTION value="javascript:spec()">定人帖</OPTION>
<%end if
if Board_Setting(10) then%>
	<OPTION value="javascript:money()">金钱帖</OPTION>
<%end if
if Board_Setting(11) then%>
	<OPTION value="javascript:point()">积分帖</OPTION>
<%end if
if Board_Setting(12) then%>
	<OPTION value="javascript:usercp()">魅力帖</OPTION>
<%end if
if Board_Setting(13) then%>
	<OPTION value="javascript:power()">威望帖</OPTION>
<%end if
if Board_Setting(14) then%>
	<OPTION value="javascript:article()">文章帖</OPTION>
<%end if
if Board_Setting(60) then%>
	<OPTION value="javascript:timel()">显示帖</OPTION>
<%end if
if Board_Setting(61) then%>
	<OPTION value="javascript:timex()">隐藏帖</OPTION>
<%end if
if Board_Setting(15) then%>
	<OPTION value="javascript:replyview()">回复帖</OPTION>
<%end if
if Board_Setting(23) then%>
	<OPTION value="javascript:usemoney()">出售帖</OPTION>
<%end if%>
</select>&nbsp;
<select name=fastpost onchange=insertsmilie(this.options[this.selectedIndex].value)>
<OPTION selected>快速发贴选项</OPTION>
<OPTION>========</OPTION>
<%dim rsfastpost
set rsfastpost=connfastpost.execute("select LabelName,LabelContent from fastpost where username='"&trim(membername)&"'")
if rsfastpost.eof or rsfastpost.bof then%>
	<OPTION>暂无快速发贴</OPTION>
<%else
	do while not rsfastpost.eof%>
		<OPTION value="<%=rsfastpost(1)%>"><%=rsfastpost(0)%></OPTION>
		<%rsfastpost.movenext
	loop
end if
rsfastpost.close
set rsfastpost=nothing%>
</select>
<%end if%>
<br>
<img onclick=Cbold() src="<%=Forum_info(7)&Forum_ubb(0)%>" width="22" height="22" alt="粗体" border="0"><img onclick=Citalic() src="<%=Forum_info(7)&Forum_ubb(1)%>" width="23" height="22" alt="斜体" border="0"><img onclick=Cunder() src="<%=Forum_info(7)&Forum_ubb(2)%>" width="23" height="22" alt="下划线" border="0"><IMG onclick=Csup() height=22 alt=上标字 src="<%=Forum_info(7)%>sup.gif" width=23 border=0><IMG onclick=Csub() height=22 alt=下标字 src="<%=Forum_info(7)%>sub.gif" width=23 border=0><img onclick=Ccenter() src="<%=Forum_info(7)&Forum_ubb(3)%>" width="22" height="22" alt="居中" border="0"><img onclick=Curl() src="<%=Forum_info(7)&Forum_ubb(4)%>" width="22" height="22" alt="超级连接" border="0"><img onclick=Cemail() src="<%=Forum_info(7)&Forum_ubb(5)%>" width="23" height="22" alt="Email连接" border="0"><img onclick=Cimage() src="<%=Forum_info(7)&Forum_ubb(6)%>" width="23" height="22" alt="图片" border="0"><img onclick=Cswf() src="<%=Forum_info(7)&Forum_ubb(7)%>" width="23" height="22" alt="Flash图片" border="0"><img onclick=Cdir() src="<%=Forum_info(7)&Forum_ubb(8)%>" width="23" height="22" alt="Shockwave文件" border="0"><img onclick=Crm() src="<%=Forum_info(7)&Forum_ubb(9)%>" width="23" height="22" alt="realplay视频文件" border="0"><img onclick=Cwmv() src="<%=Forum_info(7)&Forum_ubb(10)%>" width="23" height="22" alt="Media Player视频文件" border="0"><img onclick=Cmov() src="<%=Forum_info(7)&Forum_ubb(11)%>" width="23" height="22" alt="QuickTime视频文件" border="0"><img onclick=Cquote() src="<%=Forum_info(7)&Forum_ubb(12)%>" width="23" height="22" alt="引用" border="0"><IMG onclick=Cmark() height=22 alt=防复制水印 src="<%=Forum_info(7)%>water.gif" width=23 border=0><IMG onclick=Cfly() height=22 alt=飞行字 src="<%=Forum_info(7)&Forum_ubb(13)%>" width=23 border=0><IMG onclick=Cmarquee() height=22 alt=移动字 src="<%=Forum_info(7)&Forum_ubb(14)%>" width=23 border=0><IMG onclick=Cguang() height=22 alt=发光字 src="<%=Forum_info(7)&Forum_ubb(15)%>" width=23 border=0><IMG onclick=Cying() height=22 alt=阴影字 src="<%=Forum_info(7)&Forum_ubb(16)%>" width=23 border=0><IMG onclick=openScript('smiles.asp',500,600) src="<%=Forum_info(7)&Forum_ubb(17)%>" border=0 alt=查看更多的心情图标><img onclick=Csound() src="<%=Forum_info(7)&Forum_ubb(18)%>" width="23" height="22" alt="背景音乐" border="0">
<br>
<SCRIPT language=JavaScript src="inc/z_color.js"></SCRIPT>
<SCRIPT language=JavaScript>
	var height1 = 10;	// define the height of the color bar
	var pas = 36;			// define the number of color in the color bar
	var width1=2;
										//define the width of the color bar here (for forum with little width put 1)
</SCRIPT>
<TABLE id=ColorPanel cellSpacing=0 cellPadding=0 border=0 >
	<TBODY>
	<TR><TD id=ColorUsed onclick="if(this.bgColor.length > 0) insertTag(this.bgColor)" vAlign=center align=middle><SCRIPT language=JavaScript>document.write('<IMG height='+height1+' width=20 border=1>');</SCRIPT></TD><SCRIPT language=JavaScript>rgb(pas,width1,height1)</SCRIPT></TR>
	</TBODY>
</TABLE>
<script language=JavaScript>
	function checkbbs(flag) {
		if(flag) {
			if(frmAnnounce.subject) {
				if(frmAnnounce.subject.value=="") {
					alert("帖子的主题不应为空。");
					return false;
				} else if(strLength(frmAnnounce.subject.value)>50) {
					alert("帖子主题长度不能超过50。");
					return false;
				}
			}
		}
		if(frmAnnounce.Content.value=="") {
			alert("没有填写内容。");
			return false;
		} else if(strLength(frmAnnounce.Content.value)><%
			if boardid>0 then
				response.write Clng(Board_Setting(16))
			else
				response.write "16240"
			end if%>) {
			alert("发言内容不得大于<%
			if boardid>0 then
				response.write Clng(Board_Setting(16))
			else
				response.write "16240"
			end if%>bytes。");
			return false;
		}
		return true;
	}
</script>