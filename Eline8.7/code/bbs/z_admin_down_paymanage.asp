<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<html>
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<script language=javascript>
function check() {
	if(document.form1.username.value=="") {
		alert("�û���Ϊ��");
		return false;
	}
}
</script>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%dim str
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	if Request("action") = "savemod" then
		call savemod()
	elseif Request("action") = "saveconfig" then
		call saveconfig()
	elseif Request("action") = "alladd" then
		call alladd()
	else
		call payinfo()
	end if
end if
conndown.close
set conndown=nothing

sub payinfo()
	dim sql1,rs1
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select * from showpage"
	rs.open sql,conndown,1,1%>
	<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
		<form action=?action=savemod method=post>
		<tr>
			<th valign=middle height=23 align=left width="100%">�����û�����ϵͳģʽ�趨</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="75%">˵���������û�����ϵͳ������ģʽ<br><font color="#FF0000">�Զ����ϵͳ</font>���û���ǰ̨���Կ�������������������������������룬�����ֱ�Ӽ�Ϊ�������û���<br><font color="#FF0000">�˹����ϵͳ</font>���û���ǰ̨���Կ���������������������������Եݽ����룬������ύ������Ա�����������׼���Ϊ�����û�</td>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=center width="75%">�Զ���Ӹ����û�<input type=radio name="paymod" value="1" <%if rs("paymod")=1 then%>checked<%end if%>>���������˹���Ӹ����û�<input type=radio name="paymod" value="2" <%if rs("paymod")=2 then%>checked<%end if%>>������<input type=submit name="submit" value="�� ��"></td>
		</tr>
		</form>
	</table>
	<%rs.close%>
	<br>
	<%sql="select * from [payconfig]" 
	set rs=conndown.execute(sql)%>
	<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
		<form name="form" method="post" action="?action=saveconfig"> 
		<tr>
			<th align=left colspan=2 height=1 align="left">�����û�������������</th> 
		</tr>
		<tr> 
			<td align=right width="40%" height="25" class=forumrow>���븶���û������</td> 
			<td align=left class=forumrow><INPUT name=userwealth size="20" value="<%=rs("wealth")%>"></td> 
		</tr> 
		<tr> 
			<td align=right width="40%" height="25" class=forumrow>���븶���û����辭�飺</td> 
			<td align=left class=forumrow><INPUT name=userep size="20" value="<%=rs("userep")%>"></td> 
		</tr> 
		<tr> 
			<td align=right width="40%" height="25" class=forumrow>���븶���û�����������</td> 
			<td align=left class=forumrow><INPUT name=usercp size="20" value="<%=rs("usercp")%>"></td> 
		</tr> 
		<tr> 
			<td align=right width="40%" height="25" class=forumrow>���븶���û�����������</td> 
			<td align=left class=forumrow><INPUT name=userpower size="20" value="<%=rs("power")%>"></td> 
		</tr> 
		<tr> 
			<td align=right width="40%" height="25" class=forumrow>���븶���û�������֣�</td> 
			<td align=left class=forumrow><INPUT name=userscore size="20" value="<%=rs("userscore")%>"></td> 
		</tr> 
		<tr> 
			<td align=right width="40%" height="25" class=forumrow>���븶���û�������������</td> 
			<td align=left class=forumrow><INPUT name=userart size="20" value="<%=rs("article")%>"></td> 
		</tr> 
		<tr> 
			<td align=center colspan=2 height="27" class=forumrow><INPUT type=submit value=�޸� name=submit></td> 
		</tr> 
		</form> 
	</table>
	<%rs.close%>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
		<tr>
			<th valign=middle height=23 align=left colspan=2>�����û�����</th>
		</tr>
		<tr>
			<td class=forumrow width="40%"><b>���Ӹ����û�</b><br>����ȷ�ϴ��û��Ǳ���̳���û���</td>
			<form method="post" action="z_admin_down_addpayuser.asp?menu=2" name="form1" onsubmit="javascript:return check();">
			<td class=forumrow>����Ƕ���û����ö��ŷָ���<br><input maxLength=20 name=username size=20>&nbsp;&nbsp; <input type="submit" name="Submit" value="ȷ��"></td>
			</form>
		</tr>
		<tr>
			<td class=forumrow width="40%"><b>������ӣ�</b><br>�����û�����û���Ϊ�����û���</td>
			<form method="post" action="?action=alladd" name="form3">
			<td class=forumrow><select name=usertitle size=1><%
				sql="select * from usertitle order by usertitleid desc"
				set rs=conn.execute(sql)
				do while not rs.eof
					response.write "<option value="&rs("usertitle")&""
					response.write ">"&rs("usertitle")&"</option>"
					rs.movenext
				loop
				rs.close
			%></select>&nbsp;&nbsp;<input type="submit" name="Submit" value="���"></td>
			</form>	
		</tr>
		<tr>
			<td width="40%" class=forumrow><b>�����û�����</b></td>
			<form name="form2" method="get" action="z_admin_down_payuser.asp">
			<td valign=middle class=forumrow><INPUT type=submit value='�� �� �� ��' name=submit></td>
			</form>
		</tr> 
		<tr>
			<td width="40%" class=forumrow><b>���������û�����</b></td>
			<form name="form4" method="get" action="z_admin_down_payapplyuser.asp">
			<td valign=middle class=forumrow><INPUT type=submit value='�� �� �� ��' name=submit></td>
			</form>
		</tr> 
	</table>
	<br>
<%end sub

sub savemod()
	if request("paymod")="" then
		response.write "��������ݲ���Ϊ�գ�"
	else
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from showpage"
		rs.open sql,conndown,1,3
		rs("paymod")=request("paymod")
		rs.update
		response.write"�޸ĳɹ���"
	end if
	response.end
end sub

sub saveconfig() 
	dim money,userep,usercp,userpower,userart,userscore
	if not isInteger(request.form("userwealth")) or not isInteger(request.form("userart")) or not isInteger(request.form("userep")) or not isInteger(request.form("usercp")) or not isInteger(request.form("userpower")) or not isInteger(request.form("userscore")) then 
		response.write"��������ݱ���Ϊ������" 
		response.end 
	else 
		if request.form("userwealth")="" then 
			money=0 
		else 
			money=int(request.form("userwealth")) 
		end if 
		if request.form("userep")="" then 
			userep=0 
		else 
			userep=int(request.form("userep")) 
		end if 
		if request.form("usercp")="" then 
			usercp=0 
		else 
			usercp=int(request.form("usercp")) 
		end if 
		if request.form("userpower")="" then 
			userpower=0 
		else 
			userpower=int(request.form("userpower")) 
		end if 
		if request.form("userart")="" then 
			userart=0 
		else 
			userart=int(request.form("userart")) 
		end if 
		if request.form("userscore")="" then
			userscore=0 
		else 
			userscore=int(request.form("userscore")) 
		end if 
	end if 
	set rs= server.createobject ("adodb.recordset") 
	sql = "select * from [payconfig]" 
	rs.Open sql,conndown,1,3 
	if not(rs.eof and rs.bof) then 
		rs("article") = userart 
		rs("power") = userpower 
		rs("wealth") = money 
		rs("userep") = userep 
		rs("usercp") = usercp 
		rs("userscore") = userscore
		rs.Update 
	end if 
	rs.close 
	set rs=nothing 
	response.write "�޸ĳɹ�!"
	response.end
end sub 

sub alladd()
	dim usertitle,rs1,sql1,rss,rs2,sql2,mess
	usertitle=request("usertitle")
	Set rs1 = Server.CreateObject("ADODB.Recordset")
	sql1="select * from [user]"
	rs1.open sql1,conndown,1,3
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from [user] where userclass='"&usertitle&"'"
	rs.open sql,conn,1,1
	if not(rs.eof and rs.bof) then
		do while not rs.eof
			set rss=conndown.execute("select * from [user] where username='"&rs("username")&"'")
			if rss.eof and rss.bof then
				rs1.addnew
				rs1("username")=rs("username")
				rs1("regtime")=date()
				rs1("allpoint")=0
				rs1("lockuser")=1
				rs1("apply")=1
				rs1.update
				set rs2=server.createobject("adodb.recordset") 
				sql2="select * from [message]"
				rs2.open sql2,conn,1,3
				rs2.addnew
				rs2("sender")=forum_info(0)&"��������"
				rs2("incept")=rs("username")
				rs2("title")="���ĸ��������û��˺��Ѿ���ͨ��" 
				rs2("content")="���ĸ��������û��˺��Ѿ���ͨ�������Ʊ����������룡"
				rs2("flag")=0
				rs2("issend")=1
				rs2.update
				rs2.close
				set rs2=nothing
			end if
			rss.close
			rs.movenext
		loop
	end if
	rs.close
	set rs=nothing
	rs1.close
	set rs1=nothing
	set rs2=server.createobject("adodb.recordset")
	sql2="select * from events"
	rs2.open sql2,conndown,1,3
	mess=""&membername&"�����С�"&usertitle&"����Ϊ�����û�"	
	rs2.addnew
	rs2("event")=mess
	rs2("addtime")=date()
	rs2.update
	rs2.close
	set rs2=nothing
	response.write "��ӳɹ���"
	response.end
end sub%>