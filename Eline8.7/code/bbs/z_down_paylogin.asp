<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_down_conn.asp"-->
<!--#include file="md5.asp"-->
<%
stats="�����û��������"
call nav()
call head_var(0,0,"��������","z_down_default.asp")
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>��û�н��븶���û���������Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
	call dvbbs_error()
else
	dim UserIP
	dim username
	dim userclass
	dim password
	dim article
	dim rsp
	UserIP=Request.ServerVariables("REMOTE_ADDR")
	
	if request("password")="" then
		errmsg=errmsg+"<br>"+"<li>�������������롣"
		call dvbbs_error()
	else
		password=md5(trim(checkStr(request("password"))))
		set rs=conn.execute("select * from [user] where username='"&membername&"'")
		sql="select * from [user] where username='"&membername&"' and apply=1"
		set rsp=conndown.execute(sql)
		if rsp.eof and rsp.bof then
			errmsg=errmsg+"<br>"+"<li>�����Ǹ��ѻ�Ա���������˺Ż�û�еõ���׼"
			call dvbbs_error()
		else
			if rs("userpassword")<>password then 
				errmsg=errmsg+"<br>"+"<li>�����û����������ڣ���������������󣬻��������ʺ��ѱ�����Ա������"
				call dvbbs_error()
			else
				session("payuser")=membername
				rsp.close
				set rsp=nothing
				rs.close
				set rs=nothing
				response.redirect Request.ServerVariables("HTTP_REFERER")
			end if
		end if
		rsp.close
		set rsp=nothing
		rs.close
		set rs=nothing
	end if
end if
call footer()%>