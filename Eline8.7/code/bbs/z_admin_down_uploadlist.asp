<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<!--#include file="z_down_function.asp"-->
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
<%
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	call main()
end if
conn.close
set conn=nothing

sub main()%>
	<table cellpadding=3 cellspacing=1 border=0 width=95% align=center>
		<tr>
			<td><b>����ϴ��ļ�</b>�������ܱ��������֧��FSOȨ�޷���ʹ�ã�FSOʹ�ð��������΢����վ���ļ�Ŀ¼Ϊ<%=downpath%>���������������֧��FSO���ֶ�����</td>
		</tr>
		<tr> 
			<td><font color=<%=forum_body(8)%>>ע�⣺�����ɾ�����ļ�����ô���ݿ��й��ڴ��ļ�����������ϢҲ�ᱻɾ������С��ʹ�ã�</font></td>
		</tr>
		<tr> 
			<td><b>��ǰ��� <%=downpath%> �ļ�Ŀ¼�������ļ�</b>��</td>
		</tr>
		<tr> 
			<td><%
				dim objFSO
				dim uploadfolder
				dim uploadfiles
				dim upname
				dim upfilename
				Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
				if request("filename")<>"" then
					objFSO.DeleteFile(Server.MapPath(downpath&request("filename")))
				end if
				set rss=conndown.execute("select * from download where filename ='"&request("filename")&"' or filename1 = '"&request("filename")&"' or filename2= '"&request("filename")&"'")
				if not (rss.bof and rss.eof) then 
					sql="delete from download where filename = '"&request("filename")&"' or filename1 = '"&request("filename")&"' or filename2 = '"&request("filename")&"'"
					conndown.execute sql
				end if
				Set uploadFolder=objFSO.GetFolder(Server.MapPath(downpath))
				Set uploadFiles=uploadFolder.Files
				For Each Upname In uploadFiles
					upfilename=downpath&upname.name
					response.write "<a href="""&upfilename&""">"&upfilename&"</a> <a href='?path="&downpath&"&filename="&upname.name&"'>ɾ��</a><br>"
				next
				set uploadFolder=nothing
				set uploadFiles=nothing
			%></td>
		</tr>
	</table>
<%end sub%>







