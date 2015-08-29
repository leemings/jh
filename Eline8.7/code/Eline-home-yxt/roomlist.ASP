<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
%>
<html>
<head>
<title>房间列表♀wWw.51eline.com♀</title>
<style></style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="setup.css">
</head>

<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align="center"><font color="#0000FF">『E线江湖』房间列表</font><br>
  <br>
  显示：0则表示允许，1则表示禁止，只有主站长可以修改资料以上资料,<br>
  修改完成需服务器重启或将虚拟目录文件：global.asa改名再改回来才会生效！ </p>
<table width="90%" border="1" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000"
bgcolor="#006699">
  <tr> 
    <td colspan="8" height="11" bgcolor="#336633"> 
      <p align="center">现 有 房 间 列 表</p>
    </td>
    <td height="11" bgcolor="#336633"> 
      <div align="center"><font size="2" color="#000000"> <a href="addroom.asp"><font color="#FFFFFF">新增房间</font></a></font></div>
    </td>
  </tr>
  <tr> 
    <td width="104" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">房间名</font></div>
    </td>
    <td width="64" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">人数</font></div>
    </td>
    <td width="68" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">发招</font></div>
    </td>
    <td width="74" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">事件</font></div>
    </td>
    <td width="56" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">保护</font></div>
    </td>
    <td width="60" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">卡片</font></div>
    </td>
    <td width="72" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">赌博</font></div>
    </td>
    <td width="84" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">限制</font></div>
    </td>
    <%if sjjh_name=Application("sjjh_user") then %>
    <td width="89" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">操作</font> </div>
    </td>
    <%end if%>
  </tr>
  <%
sql="SELECT * FROM r"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
  <tr> 
    <td width="104" height="23"> 
      <div align="center"><font color="#FFFFFF" size="2"><%=rs("a")%> </font></div>
    </td>
    <td width="64" height="23"> 
      <div align="center"><font color="#FFFFFF" size="2"><%=rs("b")%></font></div>
    </td>
    <td width="68" height="23"> 
      <div align="center"><font color="#FFFFFF" size="2"><%=rs("f")%></font></div>
    </td>
    <td width="74" height="23"> 
      <div align="center"><font color="#FFFFFF" size="2"><%=rs("g")%></font></div>
    </td>
    <td width="56" height="23"> 
      <div align="center"><font color="#FFFFFF" size="2"><%=rs("h")%></font></div>
    </td>
    <td width="60" height="23"> 
      <div align="center"><font color="#FFFFFF" size="2"><%=rs("i")%></font></div>
    </td>
    <td width="72" height="23"> 
      <div align="center"><font color="#FFFFFF" size="2"><%=rs("j")%></font></div>
    </td>
    <td width="84" height="23"> 
      <div align="center"><font color="#FFFFFF" size="2"><%=rs("c")%></font></div>
    </td>
    <%if sjjh_name=Application("sjjh_user") then %>
    <td width="89" height="23"> 
      <div align="center"> <a href="modiroom.asp?wupinid=<%=rs("id")%>"><font color="#FFFFFF" size="-1">修改</font></a> 
        <%if rs("id")<>1 then %>
        <a href="delroom.asp?id=<%=rs("id")%>"><font color="#FFFFFF" size="-1">删除</font></a> 
        <%end if%>
      </div>
    </td>
    <%end if%>
  </tr>
  <tr> 
    <td colspan="9" height="23"><font size="-1" color="#FFFF00">今日话题:</font><font color="#FFFF00" size="2"><%=rs("k")%></font></td>
  </tr>
  <%if rs("c")<>0 then %>
  <tr> 
    <td colspan="9" height="16"> 
      <table width="100%" border="1" cellspacing="1" bgcolor="#006699" bordercolor="#000000">
        <tr> 
          <td width="48%" height="4"> 
            <div align="center"><font size="2" color="#000000">限制说明</font></div>
          </td>
          <td width="52%" height="4"> 
            <div align="center"><font color="#CCCCFF"></font> <font size="2" color="#000000">限制表达式</font></div>
          </td>
        </tr>
        <tr> 
          <td width="48%" height="4"> 
            <div align="center"><font color="#0000FF" size="2"><%=rs("d")%></font></div>
          </td>
          <td width="52%" height="4"> 
            <div align="center"><font color="#0000FF" size="2"><%=rs("e")%></font></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <%
 end if
rs.movenext
loop
conn.close
%>
</table>
</body>
</html>
