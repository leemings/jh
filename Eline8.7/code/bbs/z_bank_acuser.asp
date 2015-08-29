<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="Z_bankconn.asp"-->
<!--#include file="Z_bankconfig.asp"-->
<%
'chentitle="电子银行"
'chenbackasp="bank.asp"
stats="电子银行 储户用户管理"
call nav()
call head_var(0,0,"电子银行","Z_bank.asp")

dim rs1

if not master then
	Errmsg=Errmsg+"<br>"+"<li>本页面为银行行长专用，请登录后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
	call main()
	set rs1=nothing
	conn1.close
	set conn1=nothing
end if
	
call activeonline()
call footer()

sub main()
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
	 <tr>
		<th  height=25>欢迎 <%=membername%> 进入银行管理页面</th>
	</tr> 
  	<tr>
    	<td class=tablebody2 width="100%">
			<%call bankhead()%>
      		<table cellspacing="0" width="100%"  cellpadding="0" class=tableborder1 align=center style="width:97%">
        		<tr>
          			<td width="100%" class=tablebody2>
					<br>
						<%
						if request("action")="save" then
							call update()
						else
							call userinfo()
						end if
						%>
		   			</td>
				</tr>                                         
      		</table>
		</td>
   	</tr>
  	<tr>
		<td align="center" class=tablebody2><a href=Z_bank.asp>返回银行首页</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=javascript:history.go(-1)>返回上一页</a></td>
	</tr>
</table>
<%
end sub

sub userinfo()
	dim money
	set rs=server.createobject("adodb.recordset")
	sql="select userwealth from [user] where username='"&checkStr(trim(request("name")))&"' "
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then 
		rs.close
		errmsg=errmsg+"<br>"+"<li>该帐户名在论坛中已经不存在了，可以删除账号"
		call dvbbs_error()
		exit sub
	else		
		money=rs(0)
	end if
	rs.close	
	
	sql="Select * from [bank] where username='"&checkStr(trim(request("name")))&"'"
	rs.open sql,conn1,1,3
	if rs.eof and rs.bof then
		rs.close
		errmsg=errmsg+"<br>"+"<li>该帐户名不存在。"
		call dvbbs_error()
		exit sub
	else
		dim dkdu,saveday,smoney
		saveday=datediff("d",rs("date"),date())
		smoney=rs("savemoney")
		bankuser_setting=split(rs("bankuser_setting"),",")
		
		'计算贷款信誉额
		if rs("daikuang")>0 then
			dkdu=0
		elseif cint(bankuser_setting(0))=1 then
			dkdu=int(money*bankuser_setting(1))         
		else
			dkdu=int(money*bank_setting(4))	
		end if		
%>
<form method="POST" action=Z_bank_acuser.asp?action=save>
<table cellspacing="1" cellpadding="3" class=tableborder1 align="center" style="width:100%">
	<tr > 
		<th height="20" colspan="2" ><%=(rs("username"))%> 的银行信息</th>
	</tr>
	<tr> 
		<td width="41%" height="7" class=tablebody2>帐户名</td>
		<td width="59%" height="7" class=tablebody1><%=rs("username")%>
		<input type="hidden" name="name" size="35" value="<%=rs("username")%>">
		</td>
	</tr>
	<tr> 
		<td width="41%" height="9" class=tablebody2>存款</td>
		<td width="59%" height="9" class=tablebody1>
		<input type="text" name="savemoney" size="20" value="<%=rs("savemoney")%>"> 元
		</td>
	</tr>
	<tr>
		<td width="41%" height="-1" class=tablebody2>贷款</td>
		<td width="59%" height="-1" class=tablebody1>
		<input type="text" name="daikuang" size="20" value="<%=rs("daikuang")%>"> 元
		</td>
	</tr>
	<tr>
		<td width="41%" height="-1" class=tablebody2>信誉系数(注：自定义只对本用户生效)</td>
		<td width="59%" height="-1" class=tablebody1>
		<input type=radio name="credit" value=0 <%if cint(bankuser_setting(0))=0 then%>checked<%end if%>>默认&nbsp; 
		<input type=radio name="credit" value=1 <%if cint(bankuser_setting(0))=1 then%>checked<%end if%>>自定义&nbsp;	
		自定义信誉系数 <input type="text" name="creditvalue" size="10" value="<%=bankuser_setting(1)%>">  
		</td>
	</tr>	
	<tr>
		<td width="41%" height="-1" class=tablebody2>存款利息</td>
		<td width="59%" height="-1" class=tablebody1><%if saveday<Chen_StartLixiDay then%>0<%else%><%=clng((formatnumber(smoney)*(formatnumber(culi)/100))*saveday)%><%end if%> 元
		</td>
	</tr>
	<tr>
		<td width="41%" height="-1" class=tablebody2>信贷度</td>
		<td width="59%" height="-1" class=tablebody1><%=dkdu%> 元
		</td>
	</tr>
	<tr> 
		<td width="41%" height="18" class=tablebody2><font color=red>冻结帐户</font></td>
		<td width="59%" height="18" class=tablebody1> 
		<input type=radio name=lockac value="1" <%if rs("lockac") then%>checked<%end if%>>是
		<input type=radio name=lockac value="0" <%if not rs("lockac") then%>checked<%end if%>>否
		</td>
	</tr>
	
	<tr > 
		<td height="25" colspan="2" align="center" class=tablebody2> 
		<input type="submit" name="Submit" value="提交">
		</td>
	</tr>
</table>
</form>
<%
	end if
	rs.close
	set rs=nothing
end sub

sub update()
	if request.form("savemoney")="" or (not isnumeric(request.form("savemoney"))) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>存款金额必须为数字"
	elseif request.form("savemoney")<0 then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>存款金额不能是负数"		
	end if	
	if request.form("daikuang")="" or (not isnumeric(request.form("daikuang"))) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>贷款金额必须为数字"	
	elseif request.form("daikuang")<0 then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>贷款金额不能是负数"	
	end if
	if request.form("creditvalue")="" or (not isnumeric(request.form("creditvalue"))) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>信誉系数必须为数字"
	elseif request.form("creditvalue")<0 then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>贷款金额不能是负数"
	end if
	
	if founderr then
		call bank_err()
		exit sub
	end if 				
	set rs1=server.createobject("adodb.recordset")
	sql="Select * from [bank] where username='"&checkStr(trim(request("name")))&"'"
	rs1.open sql,conn1,1,3
	if rs1.eof and rs1.bof then
		errmsg=errmsg+"<br>"+"<li>该帐户名不存在。"
		call dvbbs_error()
		exit sub
	else 
    	rs1("savemoney")=abs(request.form("savemoney"))
		rs1("daikuang")=abs(request.form("daikuang"))
		rs1("bankuser_setting")=request.form("credit")&","&abs(request.form("creditvalue"))	
		if request.form("lockac")="1" then
			rs1("lockac")=true
		else
			rs1("lockac")=false
		end if
			    
		rs1.update
        rs1.close
	end if
	set rs1=nothing
	
	sucmsg=""
	if cint(log_setting(0))=1 and cint(log_setting(3))=1 then
		content="更新用户[<font color=navy>"&checkStr(trim(request("name")))&"</font>]在银行的资料"
		call logs("管理","储户信息管理",membername)
		sucmsg=sucmsg+"<br>"+"<li>您的操作信息已经记录在案"
	end if	
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>银行操作成功</th></tr>
	<tr><td width="100%" class=tablebody1><b>操作成功：</b><br><br><li>欢迎光临<%=Forum_info(0)%>电子银行
	<font color=navy><li>更新储户资料成功<%=sucmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href=<%=Request.ServerVariables("HTTP_REFERER")%>>返回上一页</a></td></tr>
</table>
<%
end sub
%>
