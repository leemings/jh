<%dim bbs_conn
dim connstrbbs_h
dim dbbbs_h
dim bbs_sql,bbs_rs
 '改为你的论坛数据库名
dbbbs_h="data/eline_bbs_6.3.0.asp"
Set bbs_conn = Server.CreateObject("ADODB.Connection")
connstrbbs_h="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbbbs_h)
bbs_conn.Open connstrbbs_h
%>