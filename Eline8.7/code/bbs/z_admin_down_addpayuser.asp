<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file=z_down_conn.asp-->
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target="footer">
<script language=javascript>
function check() {
	if(document.form1.username.value=="") {
		alert("�û���Ϊ��");
		return false;
	}
}
</script>
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
<%dim menu
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	menu=request("menu")
	if menu=1 or menu="" then
		call main()
	elseif menu=2 then
		call addpayuser()
	end if
end if

sub main()%>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form method="post" action="?menu=2" name="form1"  onsubmit="javascript:return check();">
		<tr>
			<th colspan="2">�����µĸ����û�</th>
		</tr>
		<tr>
			<td class=forumrow colspan="2">ע����� ��ȷ�ϴ��û��Ǳ���̳���û�������Ƕ���û����ö��ŷָ���</td>
		</tr>
		<tr>
			<td class=forumrow width="30%">�û�����</td>
			<td class=forumrow width="70%"><input maxLength=20 name=username size=20></td>
		</tr>
		<tr>
			<td class=forumrow colspan=2 align=center><input type="submit" name="Submit" value="ȷ��"></td>
		</tr>
		</form>	
	</table>
<%end sub

sub addpayuser()
	dim user,rs_f,mess
	if request("username")="" then	
		response.write "��������д�����û��˰�"
		exit sub
	else
		user=CheckStr(request("username"))
		user=split(user,",")
	end if
	for i=0 to ubound(user)
		set rs_f=conndown.execute("select * from [user] where username='"&user(i)&"'")
		if not(rs_f.eof and rs_f.bof) then
			Response.Write "�û�"&user(i)&"�Ѿ��Ǹ����û���<br>"
		else
			rs_f.close
			set rs_f=conn.execute("select * from [user] where username='"&user(i)&"'")
			if rs_f.eof and rs_f.bof then
				Response.Write "�û�"&user(i)&"������̳�û�<br>"
			else
				sql="select * from [user]"
				Set rs= Server.CreateObject("ADODB.Recordset")
				rs.open sql,conndown,1,3
				rs.addnew
				rs("username")=user(i)
				rs("allpoint")=0
				rs("regtime")=date()
				rs("lockuser")=1
				rs("apply")=1
				rs.update
				rs.close
				sql="select * from events"
				rs.open sql,conndown,1,3
				mess=""&membername&"����˸����û���"&user(i)&"��"	
				rs.addnew
				rs("event")=mess
				rs("addtime")=date()
				rs.update
				rs.close
				sql="select * from [message]"      
				rs.open sql,conn,1,3      
				rs.addnew      
				rs("sender")=forum_info(0)&"��������"      
				rs("incept")=user(i)
				rs("title")="���ĸ��������û��˺��Ѿ���ͨ��"      
				rs("content")="���ĸ��������û��˺��Ѿ���ͨ�������Ʊ����������룡"      
				rs("flag")=0      
				rs("issend")=1      
				rs.update      
				rs.close
				response.write "�µĸ����û�"&user(i)&"��ӳɹ���<br>"
			end if
		end if
	next
end sub%>