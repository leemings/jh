<%@LANGUAGE="VBSCRIPT"%>
<%
        option explicit
	dim conn,startime,db,sql,rs,connstr,i,m,MaxList,sjjh_name
	dim dbpath,dbbbs,connbbs,sqlbbs,rsbbs
   	set conn=server.createobject("adodb.connection")
	DBPath = Server.MapPath("51e_music.asp")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath

if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(Session("sjjh_name"))&" ")=0 then
   	Response.Write "<script language=JavaScript>{alert('��ʾ��������������Ҳſ��Բ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
        end if

	'������̳���ݿ�·�������� 
	dbbbs=Application("sjjh_usermdb")		'��̳���ݿ�·���������޸�Ϊ������̳�����ݿ�·��
	Set connbbs = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbbbs
 	connbbs.Open connstr

function CloseDatabase
	Conn.close
	connbbs.close
	set connbbs=nothing
	Set conn = Nothing
End Function

%>