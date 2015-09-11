<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<html>
<head>
<title>房间列表</title>
<style></style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<LINK href="css/css.css" type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><font color="#0000FF">房间列表<br>
  <br>
  显示：0则表示允许，1则表示禁止，只有主站长可以修改资料以上资料,<br>
  修改完成需服务器重启或将虚拟目录文件：global.asa改名再改回来才会生效！ <br><font color=red>注意：聊天大厅是npc的发放房间、恩怨情仇是动武限制判断的房间，因此这两个请不要有任何改动。</p>
<table width="90%" border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" align="center" bgcolor="#006699">
  <tr> 
    <td colspan="8" height="11" bgcolor="#336633"> 
      <p align="center"><font color="#FFFFFF">现 有 房 间 列 表</p>
    </td>
    <td height="11" bgcolor="#336633"> 
      <div align="center"><font color="#000000"> <a href="addroom.asp"><font color="#FFFFFF">新增房间</a></div>
    </td>
  </tr>
  <tr> 
    <td width="104" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">房间名</div>
    </td>
    <td width="64" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">人数</div>
    </td>
    <td width="68" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">发招</div>
    </td>
    <td width="74" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">事件</div>
    </td>
    <td width="56" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">保护</div>
    </td>
    <td width="60" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">卡片</div>
    </td>
    <td width="72" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">赌博</div>
    </td>
    <td width="84" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">限制</div>
    </td>
    <td width="89" height="4" bgcolor="#C4DEFF"> 
      <div align="center"><font color="#000000">操作 </div>
    </td>
  </tr>
  <%
sql="SELECT * FROM r"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
  <tr> 
    <td width="104" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("a")%> </div>
    </td>
    <td width="64" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("b")%></div>
    </td>
    <td width="68" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("f")%></div>
    </td>
    <td width="74" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("g")%></div>
    </td>
    <td width="56" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("h")%></div>
    </td>
    <td width="60" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("i")%></div>
    </td>
    <td width="72" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("j")%></div>
    </td>
    <td width="84" height="23"> 
      <div align="center"><font color="#FFFFFF" ><%=rs("c")%></div>
    </td>
    <td width="89" height="23"> 
      <div align="center"> <a href="modiroom.asp?wupinid=<%=rs("id")%>"><font color="#FFFFFF">修改</a> 
        <%if rs("id")<>1 then %>
        <a href="delroom.asp?id=<%=rs("id")%>"><font color="#FFFFFF">删除</a> 
        <%end if%>
      </div>
    </td>
  </tr>
  <tr> 
    <td colspan="9" height="23"><font color="#FFFF00">今日话题:<font color="#FFFF00" ><%=rs("k")%></td>
  </tr>
  <%if rs("c")<>0 then %>
  <tr> 
    <td colspan="9" height="16"> 
      <table width="100%" border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" bgcolor="#006699">
        <tr> 
          <td width="48%" height="4"> 
            <div align="center"><font color="#000000">限制说明</div>
          </td>
          <td width="52%" height="4"> 
            <div align="center"><font color="#CCCCFF"> <font color="#000000">限制表达式</div>
          </td>
        </tr>
        <tr> 
          <td width="48%" height="4"> 
            <div align="center"><font color="#0000FF" ><%=rs("d")%></div>
          </td>
          <td width="52%" height="4"> 
            <div align="center"><font color="#0000FF" ><%=rs("e")%></div>
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
<p align=center><a href=room_reset.asp?rst=1>重启房间列表</p>
</body></html>