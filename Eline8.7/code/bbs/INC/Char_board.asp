<%
Rem ==========论坛版面函数=========
Rem 判断认证论坛进入用户
function chkboardlogin(boardid,username)
	dim boarduser
	chkboardlogin=false
	if master then
		chkboardlogin=true
	else
		sql="select boarduser from board where boardid="&boardid
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			chkboardlogin=false
		else
			if isnull(rs(0)) or rs(0)="" then
				chkboardlogin=false
			else
			boarduser=split(rs(0), ",")
			for i = 0 to ubound(boarduser)
				if trim(boarduser(i))=trim(username) then
					chkboardlogin=true
					exit for
				end if
			next
			end if
		end if
		rs.close
		set rs=nothing		
	end if
end function

Rem ==========论坛版面函数=========
Rem 判断VIP是否可以进入特殊论坛
function chkviplogin(uname)
	dim rs_vip,sql_vip
	chkviplogin=false
if master then
		chkviplogin=true
else
		sql_vip="select vip from [user] where username='"&uname&"'"
		set rs_vip=server.createobject("adodb.recordset")
		rs_vip.open sql_vip,conn,1,1
		if rs_vip.eof and rs_vip.bof then
			chkviplogin=false
		else
			if int(rs_vip("vip"))=1 then
			chkviplogin=true
			else
			chkviplogin=false
			end if
		end if
	rs_vip.close
	set rs_vip=nothing
end if
end function

Rem ==========论坛版面函数=========
Rem 判断用户性别是否可以进入特殊论坛
function chksexlogin(sexnum,uname)
	dim rs_sex,sql_sex
	chksexlogin=false
if master then
		chksexlogin=true
else
		sql_sex="select sex from [user] where username='"&uname&"'"
		set rs_sex=server.createobject("adodb.recordset")
		rs_sex.open sql_sex,conn,1,1
		if rs_sex.eof and rs_sex.bof then
			chksexlogin=false
		else
			if int(rs_sex("sex"))<>int(sexnum-1) then
			chksexlogin=false
			else
			chksexlogin=true
			end if
		end if
	rs_sex.close
	set rs_sex=nothing
end if
end function
%>
