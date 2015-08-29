<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sqlstr="select * from 用户 where 姓名='"&sjjh_name&"'"
rs.open sqlstr,conn
yin=rs("银两")
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
<head><title></title><LINK rel=stylesheet href='css.css'></head>
<bgsound src="horse.mid" loop="-1">
<body oncontextmenu=self.event.returnValue=false bgcolor="#CCCCCC" topmargin="5" >
<form method="post" action="compete.asp"  target=compfrm>
<table  border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width='95%' align=center>
<tr>
<td width="27">0号</td>
<td width="34">
<input  maxlength=4 size=4  name="horse0" >
</td>
<td width="24">1号 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse1" >
</td>
<td width="24">2号 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse2" >
</td>
<td width="24">3号 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse3" >
</td>
<td width="24">4号</td>
<td width="34">
<input  maxlength=4 size=4  name="horse4" >
</td>
<td width="24">5号 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse5" >
</td>
<td width="24">6号 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse6" >
</td>
<td width="24">7号</td>
<td width="34">
<input  maxlength=4 size=4  name="horse7" >
</td>
<td width="24">8号 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse8" >
</td>
<td width="24">9号</td>
<td width="34">
<input  maxlength=4 size=4  name="horse9" >
</td>
<td width="32">10号</td>
<td width="34">
<input  maxlength=4 size=4  name="horse10" >
</td>
<td width="32">11号 </td>
<td width="7">
<input  maxlength=4 size=4  name="horse11" >
</td>
</tr>
<tr>
<td colspan=10 height="16">
<div align="center"><font size="-1"><%=Application("sjjh_chatroomname")%>赛马场</div>
</td>
<td colspan=4 align=right valign="top" height="16">
<div align="center"><font size="-1">
<input type=submit value="下 注" name="submit">
<input type=button onClick=javascript:top.window.close(); value='关 闭' name="button">
</font></div>
</td>
<td colspan=10 height="16">
<div align="center"><span style='font-size:9pt'>用户：<%=sjjh_name%>
银两：<%=yin%></span></div>
</td>
</tr>
</table>
</FORM>
</body>
