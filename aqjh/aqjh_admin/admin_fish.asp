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
<link rel="stylesheet" href="css/css.css" type=text/css>
</head>

<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<div align="center"> 
  <p><br>
    <font color="#009900"><b><font color="#FF0000" size="+6">[钓鱼场地管理]</font></b></font><font
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
      <div align="center"><font color="#FF6600">地 点</font></div>
    </td>
    <td height="27" width="123"> 
      <div align="center"><font color="#FF6600">必备渔具</font></div>
    </td>
    <td height="27" width="71" bgcolor="#99CCFF"> 
      <div align="center"><font color="#FF6600">可选渔饵</font></div>
    </td>
    <td height="27" width="71" bgcolor="#99CCFF"> 
      <div align="center"><font color="#FF6600">渔饵选择</font></div>
    </td>
    <td height="27" width="61"> 
      <div align="center"><font color="#FF6600">地点说明</font></div>
    </td>
    <td height="27" width="66"> 
      <div align="center"><font color="#FF6600">钓鱼等级</font></div>
    </td>
    <td height="27" width="66"> 
      <div align="center"><font color="#FF6600">操  作</font></div>
    </td>
  </tr>
  <%Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
rs.open "Select * From aqjh_fish",conn,0,1
if not rs.eof then
%>
  <%do while not rs.eof%>
  <form method=POST action='admin_fishok.asp?id=<%=rs("fish_id")%>'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="54"> 
        <div align="center"><font color="#FF6600">ID:</font> <font color="#FF6600"><%=rs("fish_id")%></font></div>
        <div align="center"></div>
      </td>
      <td height="2" width="94"> 
        <div align="center"> 
          <input type="text" name="address" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="12"
value="<%=rs("fish_address")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="123"> 
        <div align="center">  
          <input type="text" name="tool" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="8"
value="<%=rs("fish_tool")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="71"> 
        <div align="center"> 
          <input type="text" name="all_bait" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
value="<%=rs("fish_all_bait")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="61"> 
        <div align="center"> 
          <input type="text" name="bait" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
value="<%=rs("fish_bait")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="66"> 
        <div align="center"> 
          <input type="text" name="main" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
value="<%=rs("fish_main")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="56"> 
        <div align="center"> 
          <input type="text" name="dj" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
value="<%=rs("dy_dj")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="66"> 
          <input type="submit" value="修改"
name="submit">&nbsp;<input type="submit" value="删除"
name="submit">
      </td>
    </tr>
  </form>
<%rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
%>
</table>
</body>
</html>
