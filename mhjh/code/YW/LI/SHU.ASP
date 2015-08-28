<!--#include file="jiu.asp"--> 
<%
bgcolor=Application("yx8_mhjh_backgroundcolor")
nickname=Session("yx8_mhjh_username")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select 银两 from 用户 where 姓名='" & nickname & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	  Response.Redirect "../../error.asp?id=016"
conn.close
else
yl=rs("银两")
id=request("id")
sql="SELECT * FROM 名妓 where id="&id&""
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
mess="你有没有搞错呀，本院哪有这个姑娘呀，作好人也不能这么作啊！"
conn.close
connt.close
else
mingji=rs("姓名")
if yl<1000000 then
mess="钱不够不能赎身"
conn.close
else
sql="update 用户 set 道德=道德+1000,银两=银两-1000000 where 姓名='"&nickname&"' "
conn.execute sql
sql1="delete * from 名妓 where 姓名='"&mingji&"'"
connt.execute sql1
mess="看来您真是一位有同情心的大侠，花费100万两为"&mingji&"赎身,涨道德1000"
rs.close
conn.close
connt.close
end if
end if
end if%>
<html>

<head>
<title>怡红院</title>
<LINK href="../../style4.css" rel=stylesheet>
</head>

<BODY oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>'  leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<br><br><br><br><br>
<p align=center><%=mess%></p>
<p align=center> <a href="xiaojie.asp">返回</a></P>
</body>
</html>