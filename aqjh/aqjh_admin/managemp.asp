<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
id=Request.QueryString("ID")
sql="Select * from p where a='"&Id&"'"
rs=conn.execute(sql)
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','管理记录','修改门派资料：["&id&"]')"
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<div align="center">门派内容修改</div>
<form action="updatemp.asp?subject=<%=rs("a")%>" method=POST>
  <ul>
    <table border="0" cellspacing="1" cellpadding="4" bgcolor="#B8AF86" cellspacing=0 cellpadding=0 align="center" width="500">
      <tr> 
        <td bgcolor="ffffff" width="180">门派</td>
        <td width="308" bgcolor="f2f2ea"> 
          <input name="mp" size=40 maxlength=50 value="<%=RS("a")%>" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
          <input name="id" type="hidden" size=40 maxlength=50 value="<%=RS("id")%>">
        </td>
      </tr>
      <tr> 
        <td bgcolor="ffffff" width="180"><font class="c" size="2">掌门<font color="#cc0000">&nbsp;</font></font></td>
        <td width="308" bgcolor="f2f2ea"> 
          <input name="zm" value="<%=RS("b")%>" size=40 maxlength=50 class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
      </tr>
      <tr> 
        <td bgcolor="ffffff" width="180"><font class="c" size="2">简介</font></td>
        <td width="308" bgcolor="f2f2ea"> 
          <input name="sm" value="<%=RS("d")%>" size=40 maxlength=50 class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
      </tr>
      <tr> 
        <td width="180" bgcolor="ffffff"><font class="c" size="2">限制说明<font color="#cc0000"></font></font></td>
        <td width="308" bgcolor="f2f2ea"><font size="3" class="c" color="#000000"> 
          <input type="text" name="xzsm" size="40" value="<%=RS("e")%>" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
          </font> </td>
      </tr>
      <tr> 
        <td valign=top width="180" height="32" bgcolor="ffffff"><font size="2">表达式</font></td>
        <td width="308" height="32" bgcolor="f2f2ea"> 
          <input name="bds" value="<%=RS("g")%>" size=40 maxlength=200 class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
      </tr>
      <tr> 
        <td valign=top width="180" bgcolor="ffffff"><font size="2">门派基金</font></td>
        <td width="308" bgcolor="f2f2ea"> 
          <input name="jj" value="<%=RS("h")%>" size=40 maxlength=200 class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
      </tr>
    </table>
    <div align="center"> 
      <p><font size="3" class="c" color="#000000"> <br>
        </font></p>
      <p><font size="3" class="c" color="#000000"><br>
        <input type="HIDDEN" name="action" value="RegSubmit">
        <input type="SUBMIT" name="Submit" value="更新">
        </font></p>
    </div>
  </ul>
</form>
</body>
</html>
<%
set rs=nothing
%>