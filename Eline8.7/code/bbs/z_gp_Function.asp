<%
function iif(expression,returntrue,returnfalse)
	if expression=0 then
		iif=returnfalse
	else
		iif=returntrue
	end if
end function

'取两个数中最大那个
function Max(expression1,expression2)
	if expression1+0>=expression2+0 then
		Max=expression1
	else
		Max=expression2
	end if
end function

Rem 过滤HTML代码
function HTMLEncode(fString)
	if not isnull(fString) then
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
	
		fString = Replace(fString, CHR(32), "&nbsp;")
		fString = Replace(fString, CHR(9), "&nbsp;")
		fString = Replace(fString, CHR(34), "&quot;")
		fString = Replace(fString, CHR(39), "&#39;")
		fString = Replace(fString, CHR(13), "")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
		fString = Replace(fString, CHR(10), "<BR> ")
	
		HTMLEncode = fString
	end if
end function
'-------------------------------信息提示-------------------------------
'flag=0 显示关闭页面	flag=1 显示返回股票交易大厅		flag=2	显示返回上一页
sub endinfo(flag) 
%>
	<tr><th valign=center align=middle height=23><b>股票操作信息提示</b></th>
	<tr><td class=tablebody1>
<%
	if errmess<>"" then
		response.Write "<b>产生错误的可能原因：</b><br><br><li>您是否仔细阅读了股市指南，可能您还没有登录或者不具有使用当前功能的权限"	
		response.Write "<font color=red>"&errmess&"</font>"
	else
		response.Write "<b>操作成功：</b><br><br><li>欢迎光临"&forum_info(0)&"股票交易中心，请返回进行其它操作"
		response.Write "<font color=navy>"&sucmess&"</font>"
	end if		
%>	
	<br></td></tr> 
	<TR><Td class=tablebody2 height=25 align="center">
<%
	if flag=0 then
		response.Write "<a href=""../index.asp"">[返回论坛]</a><a href=""#"" onclick=""window.close()"">[关闭本页]</a>"
	else
    dim rUrl
		if rUrl="" or isnull(rurl) then
			rUrl=Request.ServerVariables("HTTP_REFERER")
		end if
		response.Write "<a href=""z_gp_GuPiao.asp"" class=cblue>[返回股市大厅]</a>&nbsp;&nbsp;"	
		response.Write "<a href="""&rUrl&""" class=cblue>[返回上一页]</a>"
	end if
%></Td></TR> 
<%
end sub

sub AdminHead()%>
	<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
		<tr>
			<td align=left valign=middle height="25" class=tablebody2>&nbsp;<a href="z_gp_Gupiao.asp">股票交易大厅</a> ｜ <a href="z_gp_Admin_Setting.asp">基本设置</a> ｜ <a href="z_gp_Admin_Gupiao.asp">股票管理</a> ｜ <a href="z_gp_Admin_User.asp">客户管理</a> ｜ <a href="z_gp_Announcements.asp">公告管理</a> ｜ <a href="z_gp_Admin_Data.asp">数据库管理</a> ｜ <a href=javascript:history.go(-1)>返回上一页</a></td>
		</tr> 
	</table>
	<br>
<%end sub

'-------------------------------股票收市了-------------------------------
sub CloseGuPiao(expression)
	if expression=1 then%>
		<table cellspacing=1 cellpadding=3 align=center border=0 class=tableborder1>
			<tr>
				<th height=25>股票收市通知</th>
			</tr>
			<tr>
				<td width="100%" height="50" valign="middle" class=tablebody1><br>亲爱的股民 <font color=blue><%=membername%></font>，您好！<br>&nbsp;&nbsp;&nbsp;&nbsp;现在股票市场已经收市了，请于每天的 <%=split(Gupiao_Setting(4),"||")(0)%>:00～<%=split(Gupiao_Setting(4),"||")(1)%>:00 再来。谢谢合作！<br><BR></td>
			</tr>
			<tr>
				<td align=center height=25 class=tablebody2><a href="index.asp">[返回论坛]</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="#" onClick="window.close()">[关闭本页]</a></td>
			</tr> 
		</table>
	<%elseif expression=0 then%>
		<table cellspacing=1 cellpadding=3 align=center border=0 class=tableborder1>
			<tr>
				<th height=25>欢迎光临 <%=Gupiao_Setting(5)%></th>
			</tr>
			<tr>
				<td width="100%" height="50" valign="middle" class=tablebody1><br><font color=<%=forum_body(8)%>>现在股市正处于关闭状态中，原因如下：</font><br><br>&nbsp;&nbsp;&nbsp;<%=StopGPReadme%><br><br></td>
			</tr>
			<tr>
				<td align=center height=25 class=tablebody2><a href="index.asp">[返回论坛]</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="#" onClick="window.close()">[关闭本页]</a></td>
			</tr> 
		</table>
	<%else%>
		<table cellspacing=1 cellpadding=3 align=center border=0 class=tableborder1>
			<tr>
				<th height=25>防刷新机制</th>
			</tr>
			<tr>
				<td width="100%" height="50" valign="middle" class=tablebody1><br>&nbsp;&nbsp;本页面起用了防刷新机制，请不要在 <font color=<%=forum_body(8)%>><%=Gupiao_Setting(7)%></font> 秒内连续刷新本页面<br><br>&nbsp;&nbsp;<font color=<%=forum_body(8)%>><%=Gupiao_Setting(7)%></font> 秒之后将会自动打开页面，请稍后……<br><br></td>  
			</tr> 
			<tr>
				<td align=center height=25 class=tablebody2 id="TdReFlash"></td>
			</tr> 
		</table>
		<meta HTTP-EQUIV=REFRESH CONTENT="<%=Gupiao_Setting(7)%>">
		<script language="VBScript"> 
		<!--
			TimeLimit=<%=Gupiao_Setting(7)%>
			call GetSec()
			function GetSec()
				TimeLimit=TimeLimit-1
				TdReFlash.innerHTML = "<font color=blue>"&TimeLimit&"</font>"
				if TimeLimit<= 0 then
					TdReFlash.innerHTML="<font color=red>正在打开页面……</font>"
				else
					setTimeout "GetSec()",1000 
				end if
			end function
		//-->	
		</script>		
	<%end if
end sub 

sub History()
	dim rst,TongJi,explainsplit
	set rst=gp_conn.execute("select TongJi,DangQianJiaGe,sid,ShengYuGuFen,explain from [GuPiao] order by sid")
	do while not rst.eof
		explainsplit=split(rst("explain"),"|")
		if ubound(explainsplit)<1 then
			redim preserve explainsplit(1)
			explainsplit(1)=1000
		end if
		if rst(3)>clng(explainsplit(1)) then
			sql="JiaoYiLiang="&explainsplit(1)&" "
		elseif rst(3)>0 then
			sql="JiaoYiLiang="&rst(3)
		else
			sql="JiaoYiLiang=1,ZhuangTai='封' "
		end if	
		TongJi=rst(0)
		TongJi=mid(TongJi,instr(TongJi,"|")+1)&"|"&formatnumber(rst(1),2,-1)
		gp_conn.execute("update [GuPiao] set TongJi='"&TongJi&"',KaiPanJiaGe=DangQianJiaGe,ChengJiaoLiang=0,MaiRuBiShu=0,MaiChuBiShu=0,TodayWave=0,"&sql&" where sid="&rst(2))
		rst.movenext
	loop
end sub
%>