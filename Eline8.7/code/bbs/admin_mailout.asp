<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	dim str
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
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
<td height="23" colspan="2" ><B>说明</B>：<BR>1、导出到数据库时请确认maillist.mdb（在论坛下载包的Tools目录中）已经至于论坛根目录下。<BR>2、使用导出到文本的功能需要服务器端必须支持FSO，关于FSO请查询微软的网站或<a href=http://www.aspsky.net>http://www.aspsky.net</a><BR>3、导出邮件列表非常耗费服务器资源，请尽量在本地或在网络不繁忙的时候执行<br></font></td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<form name="maildbout" method="post" action="admin_mailout.asp?action=maildb">
<tr> 
<th width="100%" colspan=2 align=center>邮件列表批量导出到数据库</th>
</tr>
        <tr>
    <td>导出邮件列表到数据库：
      <input type="text" name="maildb" value="maillist.mdb">
      <input type="submit" name="Submit" value="开始">      
    </td>
  </tr>
  </form>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<form name="mailtxtout" method="post" action="admin_mailout.asp?action=mailtxt">
<tr> 
<th width="100%" colspan=2 align=center>邮件列表批量导出到文本（注意：使用该功能服务器端必须支持FSO）</th>
</tr>
  <tr>
    <td>导出邮件列表到 文 本： 
      <input type="text" name="mailtxt" value="maillist.txt">
      <input type="submit" name="Submit2" value="开始">
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
'response.write "动网论坛邮件批量导出By Quest"
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
<th width="100%" colspan=2 align=left><%response.write "操作成功，共导出 "&trs(0)&" 个用户Email地址到数据库 "&tdb&" (<a href="&tdb&"><font color=ffffff>点击这里下载回本地</font></a>)"%></th>
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
<th width="100%" colspan=2 align=left><%response.write "导出到文本"&ttxt&"成功，(<a href="&ttxt&"><font color=ffffff>点击这里下载回本地</font></a>)"%></th>
</tr>
</table>
<%
rs.close
set rs=nothing
Writefile.close
Application.unlock
end sub
%>