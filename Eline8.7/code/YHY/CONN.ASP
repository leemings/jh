<%
dim conn,DBPath
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("Eline_yhy.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath



''''''''''''''''''''''''''''''''
'请误修改以下内容 已免不能升级
'江湖妓院 大刀版
'首页www.1615.net
'vesion="broadsword V1.0" 
''''''''''''''''''''''''''''''''

%>