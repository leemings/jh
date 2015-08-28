<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
if Application("yx8_mhjh_fellow")= True then
sqlstr="select tc.cname,tc.cnumber from card tc left join 用户 tu on tc.username=tu.姓名 where tc.username='"&username&"' and tc.cnumber>0 and tu.会员=true"
else
sqlstr="select cname,cnumber from card where username='"&username&"' and cnumber>0"
end if
rst.Open sqlstr,conn
msg="<table width=95% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=1 bordercolordark=FFFFFF><tr></tr>"
do while not (rst.BOF or rst.EOF)
msg=msg&"<tr><td><a href="&chr(34)&"javascript:parent.talkfrm.settalk('//道具','"&rst("cname")&"');"&chr(34)&" target='talkfrm' onmouseover="&chr(34)&"window.status='使用道具';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"&rst("cname")&"</a></td><td width=15  bgcolor=ffff00>"&rst("cnumber")&"</td></tr>"
rst.MoveNext
loop
msg=msg&"</table>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><link rel="stylesheet" href="css3.css"></head>
<body   oncontextmenu=self.event.returnValue=false leftmargin="4" marginwidth="4" background='bg1.gif'>
<div align=center><br>
<hr noshade size="1" color=red>
<font color=0000FF>使用道具</font><br>
<%=msg%>
</div>
<br>&nbsp;&nbsp;&nbsp;<a href="#" onClick="window.open('../hy/Cardshop.asp','reg','width=580,height=250,scrollbars=yes')"><font color="#F71004">去道具店看看</font></a> 
</body>
