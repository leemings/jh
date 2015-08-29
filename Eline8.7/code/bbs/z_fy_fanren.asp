<!--#include file="conn.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="z_fy_const.asp"-->
<!--#include file="z_plus_check.asp"-->
<%
'=========================================================
' Plusname: 法院监狱第三版
' File: z_fy_fanren.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
call nav()
stats="社区监狱"
call head_var(0,0,"社区法院","z_fy_fayuan.asp")
if not founduser then
   Errmsg=Errmsg+"<br>"+"<li>您没有进入本社区监狱的权限。"
call dvbbs_error()
else
   call getfyconfig()%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr><th valign=middle colspan=2 align=center height=25><b>社区监狱</b></td></tr>
<tr>
	<td valign=middle class=tablebody1 height=100>
  	<TABLE cellSpacing=0 cellPadding=0 border=0 align=center><TBODY>
  		<TR>
  			<TD align=middle> <IMG src=Images/img_fy/1-1.jpg width=81 height=70></TD>
  			<TD align=middle valign=middle> <IMG src=Images/img_fy/1-2.jpg width=358 height=70></TD>
				<TD align=middle> <IMG src=Images/img_fy/1-3.jpg width=82 height=70></TD>
			</TR>
			<TR>
				<TD align=center valign=middle background=Images/img_fy/2-1.jpg width=81>&nbsp;</TD>
				<TD valign=top bgcolor=#FFFBFF width=358><%
					sql="select username from [user] where lockuser=1"
					set rs=conn.execute(sql)
					dim Numcount,fyrs
					if rs.eof and rs.bof then
					  response.write("<p align=center><br><br><font color=blue>本社区治安良好,监狱所长没事干:)</font><br><br><table><tr>")
					else
						if master then
						 response.write("<p align=center><a href=z_fy_fyaction.asp?faction=allout><font color=red><b>【大赦天下】</b></font></a><br>")
						end if
						response.write("<p align=center>由于以下人等严重扰乱社区秩序,<br><font color=red><b>特关押在此!!!以敬效尤!!!</b></font><br><br>")
						dim unum
						unum=0
						response.write "<table cellSpacing=0 cellPadding=0 border=0 height=300 width=100% align=center><tr height=25>"
						do while not rs.eof
							unum=unum+1
							response.write "<td align=left>【<a href=dispuser.asp?name="&rs(0)&" target=_blank title='我是["&rs(0)&"]呀，谁来救我！'><font color=blue>"&left(rs(0),4)&"</font></a>】</td><td align=right>"
							if master then
								response.write "<a href=# title='探望 ["&rs(0)&"]' onClick=""javascript:window.open('z_fy_tanjian.asp?name="&rs(0)&"','','scrollbars=yes,width=450,height=510,top=0,left=300');"">"
								response.write "<b>探</b></a> <a href=admin_lockuser.asp?action=lock_3&name="&rs(0)&" title='释放 "&rs(0)&"'><b>放</b></a></td>"
							else
								if tj_set(0)="1" then
									response.write "<a href=# title='探望 ["&rs(0)&"]' onClick=""javascript:window.open('z_fy_tanjian.asp?name="&rs(0)&"','','scrollbars=yes,width=450,height=510,top=0,left=300');"">"
									response.write "<b>探</b></a>"
								end if
								if bs_set(0)="1" then  
									response.write " <a href=z_fy_fyaction.asp?faction=baoshi&name="&rs(0)&" title='"
									if jymaster then
										response.write "释放 ["&rs(0)&"]'><b>放</b>"
									else
										response.write "花费您"&bs_set(1)&"元保释 ["&rs(0)&"]么？'><b>保</b>"
									end if
									response.write "</a></td>"
								end if
							end if
							if (unum mod 3=0) then response.write "</tr><tr height=25>"
							rs.movenext
						loop
					end if
					response.write "</tr><tr height=*><td colspan=3>&nbsp;</td></tr></table>"
				%></TD>
				<TD align=center valign=middle background=Images/img_fy/2-2.jpg width=82>&nbsp;</TD>
			</TR>
			<TR>
				<TD align=middle> <IMG src=Images/img_fy/3-1.jpg width=81 height=84></TD>
				<TD align=middle> <IMG src=Images/img_fy/3-2.jpg width=358 height=84></TD>
				<TD align=middle> <IMG src=Images/img_fy/3-3.jpg width=82 height=84></TD>
			</TR>
		</TBODY></TABLE>
	</td>
	<td class=tablebody2 width=30% valign=top >☆ 典狱长：<%
		if not isnull(jyname) and jyname<>"" and jyname<>"无" then response.write "<a href=dispuser.asp?name="&jyname&" target=_blank>"&jyname&"</a>(<font color=red>特邀</font>) "
		sql="select username from [user] where usergroupid=1"    
		set rs=conn.execute(sql)    
		do while Not RS.Eof 
			response.write "<a href=dispuser.asp?name="&rs(0)&" target=_blank>"&checkstr(rs(0))&"</a> "	
			rs.movenext
		loop 
		response.write "<br><br>" 
		response.write "<b>当前监狱设置如下：</b><br><br>"   
		if bs_set(0)="1" then
			response.write "<b>保</b> 保释：允许，保释金为"&bs_set(1)&"元。<br>"
		else
			response.write "<b>保</b> 保释：不允许。<br>"
		end if
		if tj_set(0)="1" then
			response.write "<b>探</b> 探监：允许，如留言需要"&tj_set(1)&"元。<br><br>"
		else
			response.write "<b>探</b> 探监：不允许探视囚犯。<br><br>"
		end if
		%><table  cellspacing=1 cellpadding=2 width=100% border=2 >
			<tr  align=center>
				<td height=25 ><font color=#ff0000>最新事件</font></td>
			</tr>
			<tr bgcolor=#444444>
				<td><MARQUEE onmouseover=this.stop() onmouseout=this.start() scrollAmount=1 scrollDelay=10 direction=up width=100% height=200 align=middle border=0><%
					set fyrs=server.createobject("adodb.recordset")
					sql="select top 20 username,dateandtime,tousername,title,content from log where class='监狱'  order by dateandtime desc"
					fyrs.open sql,connfy
					do while Not fyRS.Eof
						response.write "&nbsp;<font color=#ffffff>"&fyrs(1)&"</font><br><div align=left>"
						select case trim(fyrs(3))
						case "抢劫"
							response.write "<font color=#FFFFFF>抢劫："&fyrs(4)&"</font>"
						case "探监" 
							response.write "<font color=#FFFFFF>探监："&fyrs(0)&"探视 "&fyrs(2)&"留言说:"&fyrs(4)&"</font>"
						case "入狱"
							response.write "<font color=yellow>入狱："&fyrs(2)&" "&fyrs(4)&"</font>"
						case "出狱" 
							response.write "<font color=green>出狱：囚犯"&fyrs(2)&" "&fyrs(4)&"</font>"
						case "保释" 
							response.write "<font color=green>出狱："&fyrs(0)&" "&fyrs(4)&"</font>"
						end select
						response.write "&nbsp;</div><br>"
						fyRS.MoveNext
					Loop
				%></marquee></td>
			</tr>
		</table>
		<br><br>
		<b>查看留言说明：</b><br>
		<li>当前囚犯留言（事件）请点击探视图<br>
		<li><a href=# title='查看已出狱人员留言'  onClick="javascript:window.open('z_fy_looktj.asp?action=list','','scrollbars=yes,width=450,height=500,top=10,left=300');"><font color=red>点这里查看已出狱人员留言簿</font></a><br>
		<li><a href=# title='查看我的留言簿'  onClick="javascript:window.open('z_fy_looktj.asp?name=<%=membername%>','','scrollbars=yes,width=450,height=500,top=10,left=300');"><font color=red>点这里查看您的留言（事件）簿</font></a><br><br><br><br><%
		if qj_set(0)="1" then 
			response.write "<li>监狱状态：在押囚犯权益没什么保障，<br>&nbsp;&nbsp;&nbsp;随时可能被抢劫哦~~~"
			if qj_set(1)="1" then response.write "<br>&nbsp;&nbsp;&nbsp;抢劫别人扣"&qj_set(2)&"点魅力"
		else
			response.write "<li>监狱状态：在押囚犯权益有保障。"
		end if
	%></tr>
</table>
<%end if
call activeonline()
call footer()
%>
