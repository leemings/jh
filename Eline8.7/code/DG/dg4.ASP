<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 体力,操作时间 from 用户 where 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=454"
	response.end
end if
sj=DateDiff("n",rs("操作时间"),now())
if sj<30 then
ss=30-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('刚刚打过工，还在休息中，请等："&ss&"分后再来！');location.href = 'javascript:history.back()';}</script>"
	Response.End 
end if
if rs("体力")<200 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的体力不够200，不可以在这里工作！');location.href = 'javascript:history.back()';}</script>"
	Response.End 
end if
conn.execute "update 用户 set 体力=体力-200,银两=银两+8000,操作时间=now() where 姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
<title>一线网络♀wWw.happyjh.com♀</title></head>
<body bgproperties="fixed" bgcolor="#000000" vlink="#000000">
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
<table height="260">
<tr>
          <td height="37"> <font color="#000000"><strong>世纪江湖打工:</strong></font> 
        <tr>
<td height="182" valign="top">辛辛苦苦作了一天保镖，你获得银两8000！体力下降200！<Br><Br>
          </td>
</tr>
<td align=center height="29">&nbsp;
<div align="right">
<input type=button value="返 回" onclick="location.href='dg.asp'">
</div>
</td></tr>
</table>
</td></tr>
</table>
</body>
</html>
