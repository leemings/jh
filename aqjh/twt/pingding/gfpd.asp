<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<html>
<head>
<title>官府评定</title>
</head>
<body bgcolor=#800000 background="../bgcheetah.gif" text="#000000">
<BR>
<p align="center"><font size="6" face="隶书">管理员评定</font></p>
<table border=1 align=center width=700 cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" style=font-size:9pt>
<tr bgcolor="#3399FF">
    <td width=80 height=40 align=center><font color=white>姓　名</td>
    <td width=40 align=center><font color=white>管理等级</td>
    <td width=40 align=center><font color=white>战斗等级</td>
    <td width=80 align=center><font color=white>最后登陆日期</td>
    <td width=80 align=center><font color=white>日均在线</td>
    <td width=80 align=center><font color=white>消失天数</td>
    <td width=40 align=center><font color=white>IP变更次数</font></td>
    <td width=260 align=center><font color=white>评定参考</td>
</tr>
<%
date2=split(date(),"-")
rs.open "SELECT * FROM 用户 where 门派='官府' and 姓名<>'风的季节' order by grade DESC",conn
do while not rs.eof and not rs.bof
ipcname=rs("姓名")
xsts=int(DateDiff("d",rs("lasttime"),now()))
rjzx=int(rs("mvalue")/date2(2)/10)

if rjzx<30 or xsts>=5 then
 pdck="<font color=red><b>该管理员严重失职，建议撤职</b></font>"
elseif rjzx>=30 and rjzx<60 then
 pdck="<b>该管理员表现很差，请尽快改正</b>"
elseif rjzx>=60 and rjzx<180 and xsts<5 then
 pdck="<b>该管理员表现一般，请继续努力</b>"
elseif rjzx>=180 and xsts<3 then
 pdck="<font color=green><b>该管理员认真负责，建议奖励</b></font>"
elseif ipc(ipcname)>=5 then
 pdck="<font color=red><b>该管理员ip变动频繁，请去论坛说明原因</b></font>"
else
 pdck="<b>该管理员表现一般，请继续努力</b>"
end if

%>
<tr bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';">
    <td align=center width=80><%=rs("姓名")%></td>
    <td align=center width=40><%=rs("grade")%></td>
    <td align=center width=40><%=rs("等级")%></td>
    <td align=center width=80><%=rs("lasttime")%></td>
    <td align=center width=80><%=rjzx%> 分钟</td>
    <td align=center width=80><%=xsts%></td>
    <td align=center width=40><a href=gfip.asp?name=<%=rs("姓名")%>><%=ipc(ipcname)%></a></td>
    <td align=center width=260><%=pdck%></td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing


function ipc(name)


   pddb="pd.mdb"
   Set pdconn = Server.CreateObject("ADODB.Connection")
   Set rspd=Server.CreateObject("ADODB.RecordSet")
   pdconnstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(pddb)
   pdconn.open pdconnstr
   sqlpd="select * from p where a='"&name&"'"
   rspd.open sqlpd,pdconn,1,3
   if not rspd.eof then
	ipcs=split(rspd("b"),"|")
	ipc=ubound(ipcs)-1
   else
	ipc="NO"
   end if
   rspd.close
   set rspd=nothing
   pdconn.close
   set pdconn=nothing


end function
%>
</table>
</body>
</html>
