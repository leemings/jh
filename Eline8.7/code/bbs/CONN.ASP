<%@LANGUAGE="VBSCRIPT"%>
<%
option explicit
dim startime,endtime,conn,connstr,db
startime=timer()
'更改数据库名字
db="data/eline_bbs_6.3.0.asp"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
conn.Open connstr

function CloseDatabase
Conn.close
Set conn = Nothing
End Function
%>
