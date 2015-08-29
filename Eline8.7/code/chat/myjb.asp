<%@ LANGUAGE=VBScript codepage ="936" %> 
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Weekday(date())=1 and Hour(time())>=20 and Hour(time())<21 then
	Response.Write "<script language=JavaScript>{alert('提示：现在为竞标时间，您不可以操作!');window.close();}</script>"
	Response.End 
end if
wpname=LCase(trim(Request.QueryString("wpname")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM b where a='"&wpname&"' and (b='药品' or b='鲜花')  order  by b,o",conn,2,2
if rs("n")="" or isnull(rs("n")) or rs("n")="无" then
	Response.Write "<script language=JavaScript>{alert('提示：["&wpname&"]现在还没有经营者!');window.close();}</script>"
	Response.End 
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
%>
<html>
<script>
window.moveTo(100,30);
</script>
<head>
<title>物品操作</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#000000" background="bg.gif">
<form method=POST action="myjbok.asp" name="form">
<table width="380" border="1" bordercolor="#000000" cellspacing="3" cellpadding="3" height="291" align="center">
  <tr> 
    <td bgcolor="#3399CC" height="2" colspan="4"> 
      <div align="center"><b><font color="#FFFF00">竞标物品[<font color=blue><%=wpname%></font>](</font></b><font color="#FFFFFF" size="-1">暂时只支持药品及鲜花</font><b><font color="#FFFF00">)</font></b></div>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#3399CC" height="23" width="21%">经营者：</td>
    <td bgcolor="#3399CC" height="23" width="25%"><b><font color="#FFFF00"><font color=blue><%=rs("n")%></font></font></b>
        <input type=hidden name=wpname value="<%=rs("a")%>" size="8" maxlength="10">
    </td>
    <td bgcolor="#3399CC" height="23" width="24%">投资数:</td>
    <td bgcolor="#3399CC" height="23" width="30%"><b><font color="#FFFF00"><font color=blue><%=rs("q")%></font></font></b>两</td>
  </tr>
  <tr> 
    <td bgcolor="#3399CC" width="21%" height="26">物价：</td>
    <td bgcolor="#3399CC" width="25%" height="26"> 
      <input type=text name=money <%if rs("n")<>sjjh_name and sjjh_grade<10 then%> readonly <%end if%> value="<%=rs("h")%>" size="8" maxlength="10">
    </td>
    <td bgcolor="#3399CC" width="24%" height="26">收入：</td>
    <td bgcolor="#3399CC" width="30%" height="26"><b><font color="#FFFF00"><font color=blue><%=rs("o")%></font></font></b>两</td>
  </tr>
  <tr> 
    <td bgcolor="#3399CC" colspan="4" height="19"> 
      <div align="center">经营者公告</div>
    </td>
  </tr>
  <tr align="left" valign="top"> 
    <td bgcolor="#3399CC" colspan="4" height="118">
      <div align="center"><b><font color="#FFFF00"><font color=blue> 
        <textarea name="ggsm" <%if rs("n")<>sjjh_name and sjjh_grade<10 then%> readonly <%end if%> cols="45" rows="5"><%=rs("p")%></textarea>
        <br>
        </font></font></b><font color="blue" size="-1">作为经营者，在这里可以修改物价及公告！</font></div>
    </td>
  </tr>
  <%if rs("n")=sjjh_name then%>
  <tr> 
    <td bgcolor="#3399CC" colspan="4"> 
      <div align="center">
        <input type=submit value=确定修改 name="submit" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid">
      </div>
    </td>
  </tr>
 <%end if%> 
 <%
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing%>
</table>
</form>
</body>
</html>
