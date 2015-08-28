<!--#include file="data.asp"--><%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
sqlstr="select 名称,等级 from mj where 所有者='"&username&"' and 类型='鬼' order by 攻击 desc"
rs.open sqlstr,connb,1,2
msgarr=";"
msgarr=";鬼斧神功;"
do while not (rs.EOF or rs.BOF)
attackname=rs("名称")
if instr(msgarr,";"&attackname&";")<>0 then msg=msg+"<a href='fh2.asp' onmouseover=""window.status='气转乾坤';return true;"" onmouseout=""window.status='';return true;"" title='气转乾坤'><font color=FF0000>"&attackname&"</font></a> 第"&rs("等级")&"重"
rs.MoveNext
loop
rs.Close
set rs=nothing
connb.Close
set connb=nothing
%>
<html>
<head>
<title>凤凰涅盘</title>
<link rel=stylesheet href='../../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#000000" text="#FFFFFF" topmargin="0">
<div align="center">
  <center>
  <table border="0" width="100%" cellspacing="0" cellpadding="0" height="391">
    <tr>
      <td width="50%" rowspan="2" height="391"><img border="0" src="0022.jpg" width="370" height="278"></td>
    </center>
    <td width="50%" height="262">
      <p align="right"><img border="0" src="0011.jpg"></td>
  </tr>
  <center>
  <tr>
    <td width="50%" height="129">
      <p align="center">　</p>
      <p align="center"><a href="fh2.asp"><font color="#00FF00">凤凰涅盘，火中重新</font></a></p>
      <p align="center">　<%=msg%></p>
      <p align="center">（鬼魂可以在这里修行，四十九重后最终转世投胎为人）</p>
    </td>
  </tr>
  </table>
  </center>
</div>
</body>
</html>

