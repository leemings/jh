<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<!--#include file="md5.asp"-->
<%dim sid,pwd,poind,rsa,rss,point,rsb
stats="�������"
call nav()
call head_var(0,0,"��������","z_Down_default.asp")
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>��û�н����������ĵ�Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
	call dvbbs_error()
else
	sid=request("sid")
	if sid="" then
		Errmsg=Errmsg+"<br>"+"<li>û����������"
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
			Errmsg=Errmsg+"<br>"+"<li>�����Ǹ����û����������˺��Ѿ���������"
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
				Errmsg=Errmsg+"<br>"+"<li>���������룡"
				call dvbbs_error()
			elseif rss("userpassword")<>md5(pwd) then
				rs.close
				set rs=nothing
				rsb.close
				set rsb=nothing
				rss.close
				set rss=nothing
				Errmsg=Errmsg+"<br>"+"<li>��������˺����벻��ȷ��"
				call dvbbs_error()
			elseif rss("userwealth")<point then
				rs.close
				set rs=nothing
				rsb.close
				set rsb=nothing
				rss.close
				set rss=nothing
				Errmsg=Errmsg+"<br>"+"<li>���˻��Ľ�����"
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
				sucmsg=sucmsg+"<br>"+"<li>������ɹ������ڿ��������ˣ�"
				call dvbbs_suc()
			end if
		end if
	end if
end if
call footer()%>
	
