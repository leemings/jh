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
<%dim username,right_class
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	if request("submit")<>"" then
		username=Request("username")
		right_class=CInt(Request("right_class"))
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.Open "Select * from [user] where username='"&username&"'",conn,1,1
		if  rs.EOF and rs.bof then
			Response.Write "���û�������̳�û�"
			rs.Close
			Set rs=Nothing
		else
			rs.close
			rs.Open "Select * from Admin where username='"&username&"'",conndown,1,1
			if rs.EOF and rs.bof then
				rs.Close
				set rs=Nothing
				sql="insert into Admin (username,Flag,addtime) values ('"&username&"',"&right_class&",date())"
				conndown.Execute sql
				conndown.Close
				set conndown=Nothing
				Response.Redirect "z_admin_down_adminuser.asp"
			else
				Response.Write "���û����Ѿ�����"
				rs.Close
				Set rs=Nothing
			end if
		end if
	else%>
		<table border="0" width="95%" cellspacing="1" align=center>
			<tr>
				<td><font color="#FF0000">��������Ա</font>���Խ���&nbsp;������ϴ����&nbsp;&nbsp;&nbsp;&nbsp;������������&nbsp;&nbsp;&nbsp;���޸�ɾ��&nbsp;&nbsp;<br><font color="#FF0000">�������Ա</font>���Խ���&nbsp;������ϴ����&nbsp;&nbsp;&nbsp;&nbsp;������������&nbsp;&nbsp;&nbsp;&nbsp;<br><font color="#FF0000">��ͨ����Ա</font>���Խ���&nbsp;������������</td>
			</tr>
		</table>
		<table width="95%" border="0" cellspacing="1" cellpadding="3" class=tableborder align=center>
		  <form method="post" action="?" name="form1" onsubmit="javascript:return check();">
			<tr> 
				<th height="25" align="center">��������Ա</th>
			</tr>
			<tr> 
				<td height="30" align="center" class=forumrow>�� �� ��<input type="text" name="username">&nbsp; <font color="#FF0000">(��ȷ�����û�Ϊ��̳��Ա)</font></td>
			</tr>
			<tr> 
				<td height="30" align="center" class=forumrow>Ȩ������</td>
			</tr>
			<tr>
				<td height="30" align="center" class=forumrow><input type="radio" name="right_class" value="4" checked>��ͨ����Ա&nbsp;&nbsp;<input type="radio" name="right_class" value="3">�������Ա&nbsp;&nbsp;<input type="radio" name="right_class" value="2">��������Ա</td>
			</tr>
			<tr>
				<td height="30" align="center" class=forumrow><input type="submit" name="Submit" value="ȷ��"></td>
			</tr>
			</form>
		</table>
	<%end if%>
<%end if
conndown.Close
Set conndown=Nothing%>
</body>
