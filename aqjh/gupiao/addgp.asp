<!--#include file="conn.asp"-->
<!--#include file="chkmaster.asp"-->
<% 
if session("aqjh_name")="" then
	response.redirect "../login.asp"
else
 	if not master then
		call endinfo("��û���ʸ����")
	else
		gname=Request.Form ("gname")
		gtot=Request.Form ("gtot")
		yp=Request.Form ("yp")
		np=Request.Form ("np")
		sql="insert into ��Ʊ (��ҵ,��ͨ��Ʊ,���̼۸�,��ǰ�۸�,����) values ('"&gname&"',"&gtot&","&yp&","&np&",'"&date()&"')"
		conn.execute sql
    		CloseDatabase		'�ر����ݿ�
		
		response.redirect "stock.asp"
	end if 
end if

sub endinfo(message) 
'-------------------------------��Ϣ��ʾ-------------------------------
%>
<table width="97%" border=0 cellspacing=1 cellpadding=3 align=center>
<tr><td align=center height=26><b>��Ϣ��ʾ</b></td></tr>
<tr><td align=center height=70><%=message%><br></td></tr>
<tr><td align=center height=26 ><a href="stock.asp">����</a></td></tr></table>
<%
end sub
%>