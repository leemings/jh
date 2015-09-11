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
<!--#include file="config.asp"-->
<%
tiaojian=Request.Form("tiaojian")
show=trim(Request.Form("show"))
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>用户资料查看程序</title>
<LINK href=css/css.css type=text/css rel=stylesheet>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"> <font color="#CC0000"><a href="javascript:this.location.reload()">刷新</a> 
  <br>
  这一些是满足条件的人！点姓名进行修改！<br>
  <font color="#FF0000"><b><%=tiaojian%></b> <br>
 <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
	if show<>"" then
		rs.open "SELECT * FROM 用户 where "& tiaojian &" order by "&show,conn
	else
		rs.open "SELECT * FROM 用户 where "& tiaojian &" order by lasttime",conn
	end if
%>
<table width="485" border="1" style="border-collapse: collapse" bordercolor="#B8AF86" cellspacing="0" cellpadding="0" bgcolor="f2f2ea" cellspacing=0 cellpadding=0 height="13">
        <tr> 
          <td width="28" height="10"> 
            <div align="center">ID</div>
          </td>
          <td width="47" height="10"> 
            <div align="center">姓名</div>
          </td>
          <td width="25" height="10"> 
            <div align="center">性别</div>
          </td>
          <td width="63" height="10"> 
            <div align="center">门派</div>
          </td>
          <td width="54" height="10"> 
            <div align="center">身份</div>
          </td>
          <td width="82" height="10"> 
            <div align="center">最后登陆时间</div>
          </td>
          <td width="58" height="10"> 
            <div align="center">江湖等级</div>
          </td>
          <%if show<>"" then%>
          <td width="73" height="10"> 
            <div align="center"><%=show%></div>
          </td>
          <%end if%>
          <td width="35" height="10"> 
            <div align="center">登录</div>
          </td>
        </tr>
        <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
        <tr> 
          <td width="28" height="30"> 
            <div align="center"><%=rs("ID")%></div>
          </td>
          <td width="47" height="30"> 
            <div align="center"><a href="SHOWUSER.asp?username=<%=rs("姓名")%>"><font color="#FF9900"><%=rs("姓名")%></a></div>
          </td>
          <td width="25" height="30"> 
            <div align="center"><%=rs("性别")%></div>
          </td>
          <td width="63" height="30"> 
            <div align="center"><%=rs("门派")%></div>
          </td>
          <td width="54" height="30"> 
            <div align="center"><%=rs("身份")%></div>
          </td>
          <td width="82" height="30"> 
            <div align="center"><%=rs("lasttime")%></div>
          </td>
          <td width="58" height="30"> 
            <div align="center"><%=rs("等级")%></div>
          </td>
          <%if show<>"" then%>
          <td width="73" height="30"> 
            <div align="center"><%=rs(show)%></div>
          </td>
          <%end if%>
          <td width="35" height="30"> 
            <div align="center"><%=rs("times")%></div>
          </td>
        </tr>
        <%
rs.movenext
loop
conn.close
%>
      </table>
<div align="center"><font color="#000000">条件人数:<b><font color="#0000FF"><%=(jl)%></b><font color="#000000">人 
</div></body></html>