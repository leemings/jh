<!--#include file="conn.asp"-->
<!--#include file="chkmaster.asp"-->
<% 
if session("sjjh_name")="" then
	response.redirect "../login.asp"
else
 	if not master then
		call endinfo("你没有资格操作")
	else
		gname=Request.Form ("gname")
		gtot=Request.Form ("gtot")
		yp=Request.Form ("yp")
		np=Request.Form ("np")
		sql="insert into 股票 (企业,流通股票,开盘价格,当前价格,日期) values ('"&gname&"',"&gtot&","&yp&","&np&",'"&date()&"')"
		conn.execute sql
    		CloseDatabase		'关闭数据库
		
		response.redirect "stock.asp"
	end if 
end if

sub endinfo(message) 
'-------------------------------信息提示-------------------------------
%>
<table width="97%" border=0 cellspacing=1 cellpadding=3 align=center>
<tr><td align=center height=26><b>信息提示</b></td></tr>
<tr><td align=center height=70><%=message%><br></td></tr>
<tr><td align=center height=26 ><a href="stock.asp">返回</a></td></tr></table>
<%
end sub
%>