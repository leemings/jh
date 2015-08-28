<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=Session("yx8_mhjh_userchatroomsn")
mycorp=session("yx8_mhjh_usercorp")
username=session("yx8_mhjh_username")
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 招式,等级 from 攻击 where 姓名='"&username&"' order by 等级 desc"
rst.Open sqlstr,conn
msgarr=";"
do while not (rst.EOF or rst.BOF)
attackname=rst("招式")
msgarr=msgarr&attackname&";"
msg=msg+"<tr><td><font color=#FFFF00>"&attackname&"</font></td><td>"&rst("等级")&"级</td></tr>"
rst.movenext
loop
rst.Close
sqlstr="select 招式 from 招式 where 门派='"&mycorp&"'"
rst.Open sqlstr,conn
do while not (rst.EOF or rst.BOF)
attackname=rst("招式")
if instr(msgarr,";"&attackname&";")=0 then msg=msg+"<tr><td ><font color=FF0000>"&attackname&"</font></td><td>0级</td></tr>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
if msg="" then msg="<tr valign=middle height='80%'><td colspan>对不起了，你没有修炼过任何的武功<br>请到聊神幻化修炼适合你的武功!!!<br><p align=center><img border=0 src=../image/duanyu.gif ></p></td></tr>"
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 6</title>
<link rel="stylesheet" href="../style2.css">
</head>
<body>
<table width="100%" border="0" cellpadding="0" cellspacing="2" bordercolor="#FF9933">
<%=msg%>
</table></body>

</html>
