<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_down_conn.asp"-->
<!--#include file="md5.asp"-->
<%
stats="付费用户控制面版"
call nav()
call head_var(0,0,"下载中心","z_down_default.asp")
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入付费用户控制面版的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
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
		errmsg=errmsg+"<br>"+"<li>请输入您的密码。"
		call dvbbs_error()
	else
		password=md5(trim(checkStr(request("password"))))
		set rs=conn.execute("select * from [user] where username='"&membername&"'")
		sql="select * from [user] where username='"&membername&"' and apply=1"
		set rsp=conndown.execute(sql)
		if rsp.eof and rsp.bof then
			errmsg=errmsg+"<br>"+"<li>您不是付费会员或是您的账号还没有得到批准"
			call dvbbs_error()
		else
			if rs("userpassword")<>password then 
				errmsg=errmsg+"<br>"+"<li>您的用户名并不存在，或者您的密码错误，或者您的帐号已被管理员锁定。"
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