<%@LANGUAGE="VBSCRIPT"%>
<%
	option explicit
	dim startime,conn,connstr,db,sql,rs
	dim dbbbs,connbbs,sqlbbs,rsbbs

	startime=timer()
	'更改数据库名字
	db="Eline_stock.asa"		'股票数据库路径，一般不需要修改
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
 	conn.Open connstr
	
        if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(Session("sjjh_name"))&" ")=0 then
   	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
        end if

	'更改论坛数据库路径和名字 
	dbbbs=Application("sjjh_usermdb")		'论坛数据库路径，必须修改为您的论坛的数据库路径
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