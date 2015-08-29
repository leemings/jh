<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<!--#include file="md5.asp"-->
<%dim sid,pwd,poind,rsa,rss,point,rsb
stats="购买软件"
call nav()
call head_var(0,0,"下载中心","z_Down_default.asp")
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入下载中心的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	call dvbbs_error()
else
	sid=request("sid")
	if sid="" then
		Errmsg=Errmsg+"<br>"+"<li>没有这个软件！"
		call dvbbs_error()
	else
		set rsb=server.createobject("adodb.recordset")
		sql="select * from [download] where id="&sid
		rsb.open sql,conndown,1,3
		point=rsb("point")
		set rs=server.createobject("adodb.recordset")
		sql="select * from [user] where username='"&membername&"' and lockuser=1 and apply=1"
		rs.open sql,conndown,1,3
		if rs.eof and rs.bof then
			rs.close
			set rs=nothing
			rsb.close
			set rsb=nothing
			Errmsg=Errmsg+"<br>"+"<li>您不是付费用户或是您的账号已经被锁定！"
			call dvbbs_error()
		else
			set rss=server.createobject("adodb.recordset")
			sql="select * from [user] where username='"&membername&"'"
			rss.open sql,conn,1,3
			pwd=request("password")
			if pwd="" then
				rs.close
				set rs=nothing
				rsb.close
				set rsb=nothing
				rss.close
				set rss=nothing
				Errmsg=Errmsg+"<br>"+"<li>请输入密码！"
				call dvbbs_error()
			elseif rss("userpassword")<>md5(pwd) then
				rs.close
				set rs=nothing
				rsb.close
				set rsb=nothing
				rss.close
				set rss=nothing
				Errmsg=Errmsg+"<br>"+"<li>您输入的账号密码不正确！"
				call dvbbs_error()
			elseif rss("userwealth")<point then
				rs.close
				set rs=nothing
				rsb.close
				set rsb=nothing
				rss.close
				set rss=nothing
				Errmsg=Errmsg+"<br>"+"<li>您账户的金额不够！"
				call dvbbs_error()
			else
				set rsa=server.createobject("adodb.recordset")
				sql="select * from [downevent]"
				rsa.open sql,conndown,1,3
				rss("userwealth")=rss("userwealth")-point
				rss.update
				rs("allpoint")=rs("allpoint")+point
				rs.update
				rsb("buyuser")=rsb("buyuser")&","&rs("id")
				rsb.update
				rsa.addnew
				rsa("username")=rs("username")
				rsa("buydown")=rsb("showname")
				rsa("addtime")=date()
				rsa("buypoint")=point
				rsa.update
				rsa.close
				set rsa=nothing
				rsb.close
				set rsb=nothing
				rs.close
				set rs=nothing
				rss.close
				set rss=nothing
				sucmsg=sucmsg+"<br>"+"<li>您购买成功，现在可以下载了！"
				call dvbbs_suc()
			end if
		end if
	end if
end if
call footer()%>
	
