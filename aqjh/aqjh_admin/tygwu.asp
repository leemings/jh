<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>

<title>怪物管理</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 宋体}
.c {  font-family: 宋体; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body background="0064.jpg">
<div align="center"> 
  <p><br>
    <font color="#009900"><b><font color="#FF0000" size="+6">[怪物管理]</font></b></font><font
color="#FFFFFF"><br>
    </font> </p>
  </div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="700"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27" width="54"> 
      <div align="center"><font color="#FF6600">ID</font></div>
    </td>
    <td height="27" width="94"> 
      <div align="center"><font color="#FF6600">怪物名</font></div>
    </td>
    <td height="27" width="123"> 
      <div align="center"><font color="#FF6600">图片说明</font></div>
    </td>
    <td height="27" width="71" bgcolor="#99CCFF"> 
      <div align="center"><font color="#FF6600">攻击</font></div>
    </td>
    <td height="27" width="61"> 
      <div align="center"><font color="#FF6600">防御</font></div>
    </td>
    <td height="27" width="66"> 
      <div align="center"><font color="#FF6600">体力</font></div>
    </td>
    <td height="27" width="56"> 
      <div align="center"><font color="#FF6600">kill</font></div>
    </td>
    <td height="27" width="56"> 
      <div align="center"><font color="#FF6600">操作</font></div>
    </td>
    <td height="27" width="141"> 
      <div align="center"><font color="#FF6600">执行</font></div>
    </td>
  </tr>
  <%Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
rs.open "SELECT * from 怪物区 ",conn
s=0
do while not rs.eof and not rs.bof
s=s+1
%>
  <form method=POST action='tygwuok.asp?a=m&id=<%=rs("id")%>'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="54"> 
        <div align="center"><font color="#FF6600">ID:</font> <font color="#FF6600"><%=rs("ID")%></font></div>
        <div align="center"></div>
      </td>
      <td height="2" width="94"> 
        <div align="center"> 
          <input type="text" name="a" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="12"
value="<%=rs("怪物")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="123"> 
        <div align="center">  
          <input type="text" name="b" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="8"
value="<%=rs("状态")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="71"> 
        <div align="center"> 
          <input type="text" name="c" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
value="<%=rs("攻击")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="61"> 
        <div align="center"> 
          <input type="text" name="d" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
value="<%=rs("防御")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="66"> 
        <div align="center"> 
          <input type="text" name="e" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
value="<%=rs("体力")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="56"> 
        <div align="center"> 
          <input type="text" name="f" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
value="<%=rs("kill")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="56"> 
        <div align="center"> 
          <select name="select" style="border: 1px solid; background-color:#EEFFEE;font-size: 9pt; border-color:
#000000 solid" >
            <option value="aa" selected>乘以2</option>
            <option value="bb">乘以4</option>
            <option value="cc">乘以6</option>
            <option value="dd">乘以8</option>
            <option>------</option>
            <option value="ee">除以2</option>
            <option value="ff">除以4</option>
            <option value="gg">除以6</option>
            <option value="hh">除以8</option>
          </select>
        </div>
      </td>
      <td height="2" width="141"> 
        <div align="center"> 
          <input type="submit" value="修改"
name="submit">
          &nbsp;&nbsp; 
          <input type="submit" value="删除"
name="submit" >
        </div>
      </td>
    </tr>
  </form>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</body>

</html>
</body>
</html>
