<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	dim str
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><B>˵��</B>��<BR>1�����������ݿ�ʱ��ȷ��maillist.mdb������̳���ذ���ToolsĿ¼�У��Ѿ�������̳��Ŀ¼�¡�<BR>2��ʹ�õ������ı��Ĺ�����Ҫ�������˱���֧��FSO������FSO���ѯ΢�����վ��<a href=http://www.aspsky.net>http://www.aspsky.net</a><BR>3�������ʼ��б�ǳ��ķѷ�������Դ���뾡���ڱ��ػ������粻��æ��ʱ��ִ��<br></font></td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<form name="maildbout" method="post" action="admin_mailout.asp?action=maildb">
<tr> 
<th width="100%" colspan=2 align=center>�ʼ��б��������������ݿ�</th>
</tr>
        <tr>
    <td>�����ʼ��б����ݿ⣺
      <input type="text" name="maildb" value="maillist.mdb">
      <input type="submit" name="Submit" value="��ʼ">      
    </td>
  </tr>
  </form>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<form name="mailtxtout" method="post" action="admin_mailout.asp?action=mailtxt">
<tr> 
<th width="100%" colspan=2 align=center>�ʼ��б������������ı���ע�⣺ʹ�øù��ܷ������˱���֧��FSO��</th>
</tr>
  <tr>
    <td>�����ʼ��б� �� ���� 
      <input type="text" name="mailtxt" value="maillist.txt">
      <input type="submit" name="Submit2" value="��ʼ">
    </td>
  </tr>
  </form>
</table>
<%
dim tempcount
set rs=conn.execute("select count(*) from [user] where useremail like '%@%'")
tempcount=rs(0)
set rs=server.createobject("adodb.recordset")
sql="select top "&tempcount&" useremail from [user] where useremail like '%@%'"
set rs=conn.execute(sql)

select case Request("action")
case "maildb"
	call mailoutdb()
case "mailtxt"
	call mailouttxt()
end select
end sub

sub mailoutdb
'response.write "������̳�ʼ���������By Quest"
dim tconn,tconnstr,trs,tsql,tdb,tempcount

tdb=request("maildb")
Set tconn = Server.CreateObject("ADODB.Connection")
tconnstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(tdb)
tconn.Open tconnstr

do while not rs.eof
set trs=tconn.execute("insert into [user](useremail) values ('"&rs(0)&"')")
rs.movenext
loop
set trs=tconn.execute("select count(*) from [user]")
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<form name="maildbout" method="post" action="admin_mailout.asp?action=maildb">
<tr> 
<th width="100%" colspan=2 align=left><%response.write "�����ɹ��������� "&trs(0)&" ���û�Email��ַ�����ݿ� "&tdb&" (<a href="&tdb&"><font color=ffffff>����������ػر���</font></a>)"%></th>
</tr>
</table>
<%
rs.close
set rs=nothing
tConn.close
Set tconn = Nothing
end sub
sub mailouttxt
dim ttxt,file,filepath,writefile

ttxt=request("mailtxt")
Set file = CreateObject("Scripting.FileSystemObject")
Application.lock
filepath=Server.MapPath(""&ttxt&"")
Set Writefile = file.CreateTextFile(filepath,true)
do while not rs.eof
Writefile.WriteLine rs(0)
rs.movenext
loop
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<form name="maildbout" method="post" action="admin_mailout.asp?action=maildb">
<tr> 
<th width="100%" colspan=2 align=left><%response.write "�������ı�"&ttxt&"�ɹ���(<a href="&ttxt&"><font color=ffffff>����������ػر���</font></a>)"%></th>
</tr>
</table>
<%
rs.close
set rs=nothing
Writefile.close
Application.unlock
end sub
%>