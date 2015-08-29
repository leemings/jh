<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file=z_down_conn.asp-->
<!--#include file=z_down_function.asp-->
<%
stats="修改删除软件"
call nav()
call head_var(0,0,"下载中心","z_down_default.asp")
if not Founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入下载中心管理的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	call dvbbs_error()
elseif flagdown>2 or flagdown="" or flagdown<=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有管理软件的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	call dvbbs_error()
else%>
	<script src="inc/ubbcode.js"></script>
	<table border="0" cellpadding="3" cellspacing="1" align=center class=tableborder1>
		<form method="POST" name=frmAnnounce action="z_down_savesoft.asp?id=<%=request("id")%>&action=edit">
		<tr>
			<th height="25" colspan=2>修 改 下 载 程 序</td>
		</tr>
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>软件类型：</b>&nbsp;</td>
			<td width="75%" class=tablebody1>选择大类：<select class="smallSel" name="classid" size="1" onChange="changelocation(document.frmAnnounce.classid.options[document.frmAnnounce.classid.selectedIndex].value)"><%
				dim DClassID,ClassID,NClassID
				classid=request("classid")
				Nclassid=request("Nclassid")
				set rs=server.createobject("adodb.recordset")
				sql="select * from Dclass"
				rs.open sql,conndown,1,1
				if rs.eof and rs.bof then
					response.write "<option value="""" name=classid>没有分类</option>"
				else
					DClassID=rs("classID")
					do while not rs.eof
						response.write "<option"
						if clng(classid)=rs("classid") then
							DClassID=rs("classID")
							response.write " selected"
						end if	
						response.write " value='"&CStr(rs("classID"))&"' name=classid>"&rs("class")&"</option>"&vbNewLine
						rs.movenext
					loop
				end if	
				rs.close
				%></select>&nbsp;&nbsp;选择小类：<select name="Nclassid" size="1"><%
				sql="select * from DNclass where classid="&DClassID&" order by Nclassid"
				rs.open sql,conndown,1,1
				if rs.eof and rs.bof then
					response.write "<option value="""" name=Nclassid>没有分类</option>"
				else
					do while not rs.eof
						response.write "<option"
						if clng(Nclassid)=rs("Nclassid") then
							response.write " selected"
						end if
						response.write " value='" + Cstr(rs("Nclassid")) + "' name=Nclassid>" + rs("Nclass") + "</option>"
						rs.MoveNext
					Loop
				end if
				rs.close
			%></select>&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr>
		<%sql="select * from download where id="&request("id")
		rs.open sql,conndown,1,1%>
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>下载属性：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><input type=radio name="downshow" value="1" <%if rs("downshow")=1 then%>checked<%end if%>>开放下载 <input type=radio name="downshow" value="2" <%if rs("downshow")=2 then%>checked<%end if%>>会员下载<%if vipshow=2 then%> <input type=radio name="downshow" value="3" <%if rs("downshow")=3 then%>checked<%end if%>>VIP下载<%end if%> <input type=radio name="downshow" value="4" <%if rs("downshow")=4 then%>checked<%end if%>>特约下载 <input type=radio name="downshow" value="5" <%if rs("downshow")=5 then%>checked<%end if%>>付费下载 <input name="money" size=6 value="<%=rs("point")%>"></td>
		</tr> 
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>软件属性：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><select name="softsx" size="1"><option value="0" name="softsx" <%if rs("softsx")=0 then%>selected<%end if%>>普通软件</option><option value="1" name="softsx" <%if rs("softsx")=1 then%>selected<%end if%>>RealPlay播放</option><option value="2" name="softsx" <%if rs("softsx")=2 then%>selected<%end if%>>MediaPlay播放</option><option value="3" name="softsx" <%if rs("softsx")=3 then%>selected<%end if%>>FLASH动画</option></select>&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr>
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>显示名称：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><input type="text" name="txtshowname" size="103" class="smallinput" maxlength="100" value="<%=rs("showname")%>">&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr> 
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>下载地址1：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><input type="text" name="txtfilename" size="103" class="smallinput" maxlength="100" value="<%=rs("filename")%>">&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font>&nbsp;<font color=red>是否防盗</font><input type="checkbox" name="localdown" value="on" <%if rs("ldown")=1 then%>checked<%end if%>></td>
		</tr>
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>下载地址2：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><input type="text" name="txtfilename1" size="103" class="smallinput" maxlength="100" value="<%=rs("filename1")%>"></td>
		</tr>
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>下载地址3：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><input type="text" name="txtfilename2" size="103" class="smallinput" maxlength="100" value="<%=rs("filename2")%>"></td>
		</tr>
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>每日下载次数限制：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><input type="text" name="daydown" size="15" class="smallinput" maxlength="100" value="<%=rs("daydown")%>">人次（填写“0”表示无限制）&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr>
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>下载有效期：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><select name="usetime" size="1"><option value="0" name="usetime" <%if rs("usetime")=0 then%>selected<%end if%>>无限期</option><option value="1" name="usetime" <%if rs("usetime")=1 then%>selected<%end if%>>一天</option><option value="2" name="usetime" <%if rs("usetime")=2 then%>selected<%end if%>>二天</option><option value="3" name="usetime" <%if rs("usetime")=3 then%>selected<%end if%>>三天</option><option value="7" name="usetime" <%if rs("usetime")=7 then%>selected<%end if%>>一周</option><option value="14" name="usetime" <%if rs("usetime")=14 then%>selected<%end if%>>二周</option><option value="30" name="usetime" <%if rs("usetime")=30 then%>selected<%end if%>>一个月</option><option value="60" name="usetime" <%if rs("usetime")=60 then%>selected<%end if%>>二个月</option><option value="90" name="usetime" <%if rs("usetime")=90 then%>selected<%end if%>>三个月</option><option value="180" name="usetime" <%if rs("usetime")=180 then%>selected<%end if%>>半年</option><option value="360" name="usetime" <%if rs("usetime")=360 then%>selected<%end if%>>一年</option></select>&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr>
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>推荐度：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><select name="hot" size="1"><option value="1" name="hot" <%if rs("hot")=1 then%>selected<%end if%>>一星级</option><option value="2" name="hot" <%if rs("hot")=2 then%>selected<%end if%>>二星级</option><option value="3" name="hot" <%if rs("hot")=3 then%>selected<%end if%>>三星级</option><option value="4" name="hot" <%if rs("hot")=4 then%>selected<%end if%>>四星级</option><option value="5" name="hot" <%if rs("hot")=5 then%>selected<%end if%>>五星级</option></select>&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr>
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>运行环境：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><select name="system" size="1"><option value="ASP环境" name="system" <%if rs("system")="ASP环境" then%>selected<%end if%>>ASP环境</option><option value="PHP环境" name="system" <%if rs("system")="PHP环境" then%>selected<%end if%>>PHP环境</option><option value="CGI环境" name="system" <%if rs("system")="CGI环境" then%>selected<%end if%>>CGI环境</option><option value="JSP环境" name="system" <%if rs("system")="JSP环境" then%>selected<%end if%>>JSP环境</option><option value="Windows环境" name="system" <%if rs("system")="Windows环境" then%>selected<%end if%>>Windows环境</option><option value="Linux环境" name="system">Linux环境</option></select>&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr>
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>软件大小：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><input type="text" name="size1" size="20" class="smallinput" maxlength="100" value="<%=rs("size")%>"></td>
		</tr>
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>软件来源：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><input type="text" name="fromurl" size="103" class="smallinput" maxlength="100" value="<%=rs("fromurl")%>"></td>
		</tr>
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>授权形式：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><select name="order" size="1"><option value="共享软件" name="order" <%if rs("orders")="共享软件" then%>selected<%end if%>>共享软件</option><option value="付费软件" name="order" <%if rs("orders")="付费软件" then%>selected<%end if%>>付费软件</option><option value="商业软件" name="order" <%if rs("orders")="商业软件" then%>selected<%end if%>>商业软件</option></select>&nbsp;是否推荐<input type="checkbox" name="hots" value="on" <%if rs("hots")=1 then%>checked<%end if%>>&nbsp;是否隐藏<input type="checkbox" name="hide" value="on" <%if rs("stop")=1 then%>checked<%end if%>>&nbsp;是否固顶推荐<input type="checkbox" name="gudin" value="on" <%if rs("gudin")=1 then%>checked<%end if%>></td>
		</tr>
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>上传图片：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><iframe name=ad frameborder=0 width=400 height=24 scrolling=no src=z_down_picupload.asp></iframe></td>
		</tr>
		<tr>
			<td width="25%" align="right" height="30" class=tablebody1><b>软件图片：</b>&nbsp;</td>
			<td width="75%" height="30" class=tablebody1><input type="text" name="pic" size="103" class="smallinput" maxlength="100" value="<%=rs("pic")%>"></td>
		</tr>
		<tr>
			<td width="25%" align="right" valign="top" class=tablebody1><br><b>程式简介：</b>&nbsp;</td>
			<td width="75%" class=tablebody1><!--#include file="inc/getubb.asp"--><textarea rows="6" name="Content" cols="103" class="smallarea"><%=server.htmlencode(replace(replace(rs("note"),"<br>",chr(13)),"&nbsp;",""))%></textarea>&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr>
		<tr>
			<td colspan=2 height="30" align=center class=tablebody2><input type="submit" value=" 修 改 " name="cmdok">&nbsp; <input type="reset" value=" 清 除 " name="cmdcancel"></td>
		</tr>
		</form>
  </table>
	<%rs.close
	set rs=nothing
end if
call activeonline()
call footer()
set rs=server.createobject("adodb.recordset")
sql = "select * from DNclass order by Nclassid asc"
rs.open sql,conndown,1,1%>
<SCRIPT language = "JavaScript">
	var onecount;
	onecount=0;
	subcat = new Array();
	<%dim count
	count = 0
	do while not rs.eof%>
		subcat[<%=count%>] = new Array("<%= trim(rs("Nclass"))%>","<%=cstr(rs("classid"))%>","<%=cstr(rs("Nclassid"))%>");
		<%count = count + 1
		rs.movenext
	loop
	rs.close%>
	onecount=<%=count%>;

	function changelocation(locationid) {
		document.frmAnnounce.Nclassid.length = 0; 
		var locationid=locationid;
		var i;
		for (i=0;i < onecount; i++) {
			if (subcat[i][1] == locationid) { 
				document.frmAnnounce.Nclassid.options[document.frmAnnounce.Nclassid.length] = new Option(subcat[i][0], subcat[i][2]);
			} 
		}
	} 
</SCRIPT> 
