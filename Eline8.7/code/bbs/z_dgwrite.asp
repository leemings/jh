<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'---------------------
' write by 绿水青山
'---------------------
	stats="撰写点歌"

	call nav()
	call head_var(0,0,"点歌控制面板","z_dglistall.asp")
		
if not founduser then
	errmsg=errmsg+"<br>"+"<li>客人不可以查阅点歌列表"
	errmsg=errmsg+"<br>"+"<li>请先<a href=login.asp><font color=blue>登录</font></a>,还没有<a href=reg.asp><font color=blue>注册</font></a>"
	founderr=true
end if
if founderr then
	call dvbbs_error()
else
%>
	<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
		<tr>
			<th><a href=z_dglistall.asp><font color=white><b>所有点歌列表</b></font></a></th>
			<th><a href=z_dglistme.asp><font color=white><b>我的点歌列表</b></font></a></th>
			<th><a href=z_dgwrite.asp><font color=white><b>我要点歌</b></font></a></th>
		</tr>
		<tr>
			<td class=tablebody2 align=center valign=middle colspan="3">点首歌曲祝福你的好友</td>
		</tr>
	</table>
		<br>
	<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1 style="width:55%">
		<tr> 
			<th colspan="2"><b>请完整输入下列信息</b></th> 
		</tr> 
			<tr>
				<td colspan="2" class=tablebody2><%
				if isnull(myvip) or myvip<>1 then
					response.write "<font color=forum_info(8)>您当前还有现金："&mymoney&"元，为全体会员点歌需要现金"&forum_user(18)&"元，为单一会员点歌需要现金"&forum_user(19)&"元</font>"
				else
					response.write "<font color=forum_info(8)>您当前还有现金："&mymoney&"元，为全体会员点歌需要现金"&forum_user(20)&"元，为单一会员点歌需要现金"&forum_user(21)&"元</font>"
				end if%></td>
			</tr>		
		<form name=codeform action="z_dgsave.asp" method=post> 
			<tr> 
				<td class=tablebody2 width="25%" valign=middle><b>送给：</b><br>支持多用户</td>
				<td class=tablebody2 width="75%" valign=middle> 
					<input name="incept" type=text id="incept" value="<%=Trim(Request.QueryString("name"))%>" size=40>
			&nbsp;<font color="#FF0000">*</font> <br>使用逗号&quot;|&quot;分开，最多5位用户		<input name="alluser" type=checkbox id="alluser" value="" onclick="{if(document.forms[0].incept.value=='全体会员') {document.forms[0].incept.value='';return;} document.forms[0].incept.value='全体会员';return;}">全体会员</td> 
			</tr> 
			<tr> 
				<td class=tablebody2 width="25%" height="24" valign=top ><b>歌曲名字：</b><br>50字以内</td>
				<td class=tablebody2 width="75%" valign=middle> 
					<input name="medianame" type=text id="medianame" size=40 maxlength=50> &nbsp;<font color="#FF0000">*</font> 
				</td>
			</tr> 
			<tr> 
				<td class=tablebody2 valign=top><b>音乐地址：</b></td>
				<td class=tablebody2 valign=middle>
					<input name="url" type="text" id="url" value="http://" size="50" maxlength="100">&nbsp;<font color="#FF0000">*</font><br>
					<font color=ff0000><b>→</b></font><a href=http://search.sogua.com/search/search.asp target=_blank title=搜索音乐><font color=blue><b><u>点击这里去找歌</u></b></font></a><br>
        支持格式:mp3、mid、wav、wma、swf、mpg、rm、ra、asf<br>
        支持协议:http、rtsp、mms、pnm </td>
			</tr> 
			<tr> 
				<td class=tablebody2 valign=top><b>祝福内容：</b><br>200字以内
					<br>
					<br> 
				</td>
				<td class=tablebody2  valign=middle>
					<textarea name="content" cols="50" rows="5" id="content"></textarea> &nbsp;<font color="#FF0000">*</font>
				</td>
			</tr> 
			<tr> 
				<td class=tablebody2  valign=middle colspan=2 align=center>
					<INPUT TYPE="submit" NAME="Submit" VALUE="送出">　　　
					<INPUT TYPE="reset" NAME="Submit2" VALUE="重写">
				</td>
			</tr>
			<tr>
				<td colspan="2" class=tablebody2>&nbsp;</td>
			</tr>			
		</FORM>
	</table>
	<br>
	<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
		<tr>
			<th><font color=white>--== 论坛点歌台-撰写点歌 ==--</font></th>
		</tr>
	</table>
<%
call activeonline()
call footer() 
end if	
%>	