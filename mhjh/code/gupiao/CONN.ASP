<%
	option explicit
	dim startime,conn,connstr,db,sql,rs
	dim dbbbs,connbbs,sqlbbs,rsbbs

	startime=timer()
	'更改数据库名字
	db="stock1.asp"		'股票数据库路径，一般不需要修改
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
 	conn.Open connstr
	
if session("yx8_mhjh_inchat")="" then
Response.write "<script language='javascript'>alert ('你不能进入,请先进入聊天室再来。谢谢合作'); window.close();</script>"
Response.End 
        end if

	'更改论坛数据库路径和名字 
	dbbbs=Application("yx8_mhjh_connstr")		'论坛数据库路径，必须修改为您的论坛的数据库路径
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