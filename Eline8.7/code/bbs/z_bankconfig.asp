<%
	dim culi,daili,zhuangli,chubei,daitian,dmoney
	dim Chen_StartLixiDay		'开始计算利息的存款天数		绿水青山  2002.12.25
	dim Chen_BusinessTimeSlice  '银行营业时间				绿水青山  2002.12.26
	dim bankstate,bank_setting,bankuser_setting
	dim log_setting	   '银行操作事件,记录全部事件,保存最新事件条数,记录管理事件
	dim content

    set rs=server.createobject("adodb.recordset")
    sql="select * from [bankconfig]"
    rs.open sql,conn1,1,3
    culi=rs("savedayli")     '存款利息
    daili=rs("ddayli")       '贷款利息
    zhuangli=rs("zzli")		 '转帐利息 
    chubei=rs("chubei")		 '银行储备 
    daitian=rs("dkday")		 '贷款天数限制   
	bankstate=rs("state")    '银行当前状态   1-营业   0-关闭       我来了添加 2002.11.30
	bank_setting=split(rs("bank_setting"),",")    ' 我来了添加 2002.11.30
	log_setting=split(rs("log_setting"),",")    ' 我来了添加 2002.11.30
	Chen_StartLixiDay=cint(rs("StartLixiDay"))		' 绿水青山  2002.12.25
	Chen_BusinessTimeSlice=split(bank_setting(7),"||")	'银行营业起始时间   绿水青山 2002.12.25
    rs.close
	
sub bankhead()
%>
<br>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
	<tr>
		<td align=left valign=middle class=TopLighNav1> <a href=Z_bank.asp><font color=navy>银行营业大厅</font></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=Z_bank.asp?menu=1><font color=navy>社区二十大富翁</font></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=Z_bank.asp?menu=2><font color=navy>银行二十大储户</font></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=Z_fuwu.asp><font color=navy>社区服务</font></a></td>
		<%if master then%><td align=right valign=middle class=TopLighNav1> <a href="Z_bank.asp?menu=8"><font color=red>银行管理中心</font></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=Z_bank_user.asp><font color=red>储户管理</font></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=Z_fuwu.asp?action=admin><font color=red>服务管理中心</font></a><%end if%><%if master or cint(log_setting(6))=1 then%><%if not master then%><td align=center valign=middle class=TopLighNav1><%end if%>&nbsp;&nbsp;&nbsp;<a href=Z_bank.asp?menu=13 title=查看银行服务事件记录><font color=blue>事件</font><%end if%></a></td> 
	</tr> 
</table>
<br>
<%
end sub
'-------------------------------出错信息-------------------------------
sub bank_err() 
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>银行错误信息</th></tr>
	<tr><td width="100%" class=tablebody1><b>产生错误的可能原因：</b><br><br><li>您是否仔细阅读了<a href="boardhelp.asp?boardid=<%=boardid%>">帮助文件</a>，可能您还没有登录或者不具有使用当前功能的权限
	<font color=<%=Forum_body(8)%>><%=errmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href="javascript:history.go(-1)">返回上一页</a></td></tr>
</table>
<%
end sub
'--------------------银行关闭告示----------------------
sub BankClose()
%>
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:97%">
	<tr>
		<th height=25><%=stats%></td>
	</tr>
	<tr>
		<td class=tablebody2 height=1>
			<%call bankhead()%>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
				<tr><th height=25>银行已经关门了</th></tr>
				<tr><td width="100%" class=tablebody1>
					亲爱的客户 <font color=blue><%=membername%></font>，您好！<br>
					&nbsp;&nbsp;&nbsp;&nbsp;现在银行已经关门了，银行的营业时间是 <%=Chen_BusinessTimeSlice(0)%>:00～<%=Chen_BusinessTimeSlice(1)%>:00 。请在营业时间内进行您的银行操作，谢谢合作！ 
				</td></tr>
				<tr><td align=center height=26 class=tablebody2><a href="index.asp">返回论坛</a></td></tr> 
			</table>
		</td>
	</tr>
	</table>		
<%
end sub
'-------------------------------成功信息-------------------------------
sub bank_suc() 
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>银行操作成功</th></tr> 
	<tr><td width="100%" class=tablebody1><b>操作成功：</b><br><br><li>欢迎光临<%=Forum_info(0)%>电子银行
	<font color=navy><%=sucmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href=<%=Request.ServerVariables("HTTP_REFERER")%>>返回上一页</a></td></tr>
</table>
<%
end sub

sub logs(lclass,title,userN)
'我来了 2002.12.01 写入事件
	if UserID="" or (not isnumeric(UserID)) then UserID=0
	if content="" then content="可疑的操作"
	conn1.execute("insert into log ([class],UserName,UserID,Title,content,DateAndTime,IP) values('"&lclass&"','"&userN&"',"&UserID&",'"&Title&"','"&content&"',now(),'"&Request.ServerVariables("REMOTE_ADDR")&"')")
end sub
%>

