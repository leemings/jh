<!--#include file="data.asp"--><%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
sqlstr="select 名称,等级,状态 from mj where 所有者='"&username&"' and 类型='轩' order by 等级 desc"
rs.open sqlstr,connb,1,2
if rs.EOF or rs.BOF then
msg="你还没有任何要塞武器"
end if
do while not (rs.EOF or rs.BOF)
attackname=rs("名称")
zt=rs("状态")
if zt=True then lx="攻击"
if zt=False then lx="防御"
ll="等级"
msg=msg+"<a href='xl2.asp?mg="&attackname&"' onmouseover=""window.status='修炼要塞武器';return true;"" onmouseout=""window.status='';return true;"" title='修炼要塞武器'><font color=FF0000>"&attackname&"</font></a> "&rs("等级")&"级 "&lx&"型 "
rs.MoveNext
loop
rs.Close
set rs=nothing
connb.Close
set connb=nothing
%>
<html>
<head>
<title>修炼要塞武器</title>
<link rel=stylesheet href='../../style.css'>
</head>
<body oncontextmenu="self.event.returnValue=false" bgcolor="#000000" text="#FFFFFF" topmargin="0">
<p align="center"><img border="0" src="wu.gif"></p>
<div align="center">
  <center>
  <table border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="100%"><%=msg%></td>
    </tr>
  </table>
  </center>
</div>
<p align="center"><a href="xx.asp"><font color="#FF0000">创建要塞武器</font></a></p>
</body>
</html>



























