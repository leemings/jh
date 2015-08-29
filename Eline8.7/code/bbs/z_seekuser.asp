<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_plus_check.asp"-->
<% 
dim nam
stats="会员登录情况"
call nav() 
call head_var(2,0,"","")
if not founduser or membername="" then
	errmsg=errmsg+"<br>"+"<li>您还没有<a href=login.asp>登录论坛</a>，不能使用这个功能。。如果您还没有<a href=reg.asp>注册</a>，请先<a href=reg.asp>注册</a>！"
	founderr=true
	call dvbbs_error()
else%>
<table class=tableborder1 cellspacing=1 cellpadding=3 align=center>
	<tr> 
		<th height="25" align="center" colspan=7>会员查询</th>
	</tr>
	<form method="POST" action="?">
	<tr> 
		<td class=tablebody1 valign=middle height=50 align=center>输入您要查询的会员: <input type="text" name="nam1" size="12">  <input type="submit" value="提交" name="B1"></td>
	</tr>
	</form>
</table>
<br>
<%nam=checkstr(trim(request("nam1")))
if nam<>"" then
	call hy(nam)
end if
end if
call activeonline()
call footer()

sub hy(nam)
	dim page,pagecount,totalrec,pagesize
	dim GlTime,sex
	
	pagesize=20
	page=trim(request("page"))
	if page="" then
		page=1
	else
		page=cint(page)
	end if
	set rs = server.CreateObject ("adodb.recordset")
	sql="select username,adddate,Sex,Userclass,UserGroup,lastlogin from [user] where username like '%"&nam&"%' order by datediff('d',lastlogin,date()) "
	rs.open sql,conn,1,1
	totalrec=rs.recordcount
	if totalrec mod pagesize=0 then
		pagecount=totalrec \ pagesize
	else
		pagecount=totalrec \ pagesize+1
	end if
	if page > pagecount then page = pagecount
	if page < 1 then page=1
	rs.PageSize = pagesize%>
	<table Class=TableBorder1 cellspacing="1" cellpadding="1" align="center" >
		<tr>
			<th height="25" align="center" colspan=7>会员登录情况</td>
		</tr>
  	<%if rs.EOF or rs.BOF then%>
  		<tr><td align="center" height="25" class=tablebody1 colspan=7>没有你说的这样的会员哦！！</td></tr>
  	<%else%>
   		<tr height=23 align=center> 
    		<td class=tablebody2>【会员姓名】</td> 
    		<td class=tablebody2>【注册日期】</td>
    		<td class=tablebody2>【消失天数】</td>
				<td class=tablebody2>【会员性别】</td>
    		<td class=tablebody2>【会员等级】</td>
				<td class=tablebody2>【会员门派】</td>
    		<td class=tablebody2>【发送信息】</td>
			</tr>
  		<%rs.AbsolutePage=page
  		i=0
  		do while not rs.EOF%>
  			<tr height=23 align=center> 
  				<td class=tablebody1> <a href="dispuser.asp?name=<%=rs(0)%>" target="_blank" title="查看<%=rs(0)%>的个人资料"><%=rs(0)%></a> </td>
  				<td class=tablebody1><%=rs(1)%></td>
					<% GlTime= datediff("d",rs(5),date()) %>
  				<td width="10%" class=tablebody1><b><font color=red><%=GlTime%></font></b> 天</td>
					<% if rs(2)=1 then
   					sex="<font color=#0066CC>男</font>"
   				elseif rs(2)=0 then
   					sex="<font color=red>女</font>" 
					end if %>
  				<td class=tablebody1><%=sex%></td>
  				<td class=tablebody1><%=rs(3)%></td>
  				<td class=tablebody1><%=rs(4)%></td>	
  				<td class=tablebody1><a target="_blank" href="messanger.asp?action=new&touser=<%=rs("username")%>"><img src="pic/message.gif" alt=给他留言 border="0"></a></td>
				</tr>
				<%	i=i+1
				rs.movenext
				if i>= rs.PageSize then exit do
			loop  
		end if %>
		<tr align="center"> 
			<td height="25" align=center class=tablebody2><a href="z_viewmanage.asp">[版主工作]</a></td>
			<td height="25" align=left class=tablebody2>&nbsp;共 <b><%=rs.RecordCount%></b> 人 第 <b><%=page%></b> / <b><%=rs.PageCount%></b> 页</td>
			<td height="25" align=right class=tablebody2 colspan=5><%call disppagenum(page,pagecount,"?page=","&nam1="&nam)%>&nbsp;</td>
		</tr>
	</table>
	<%rs.Close
	set rs = nothing 
end sub
%>
