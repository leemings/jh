<%
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from map order by posr,posc",conn
for i=1 to 24
msg=msg&"<tr align=center><td bgcolor=ffff00>r"&i&"</td>"
for j=0 to 29
id=rst("id")
posr=rst("posr")+1
posc=rst("posc")+1
displace=rst("displace")
affair=rst("affair")
msg=msg&"<td><a href='chgmap.asp?id="&id&"' title="&chr(34)&"R:"&posr&"；C:"&posc&"；可移动："&displace&"；事件："&affair&chr(34)&"><img src='../image/land/land_r"&posr&"_c"&posc&".gif' border=0></a></td>"
rst.MoveNext
next
msg=msg&"</tr>"
next
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel='stylesheet' href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif'>
<p align=center>地图系统</p>
<hr>
<table border=2 width=58% align=center>
<tr bgcolor=#FFFF00 align=center><td>r\c</td><td>c1</td><td>c2</td><td>c3</td><td>c4</td><td>c5</td><td>c6</td><td>c7</td><td>c8</td><td>c9</td><td>c10</td><td>c11</td><td>c12</td><td>c13</td><td>c14</td><td>c15</td><td>c16</td><td>c17</td><td>c18</td><td>c19</td><td>c20</td><td>c21</td><td>c22</td><td>c23</td><td>c24</td><td>c25</td><td>c26</td><td>c27</td><td>c28</td><td>c29</td><td>c30</td></tr>
<%=msg%>
</table>
</body>
