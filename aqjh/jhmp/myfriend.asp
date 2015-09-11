<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>拉人记录</title>
<style type="text/css">
<!--
p            { line-height: 20px; font-size: 9pt }
table        { font-size: 9pt }
a:link       { color: #FF9900; text-decoration: none }
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" vlink="#0000FF" topmargin="0"
leftmargin="0" background="../bg.gif" alink="#0000FF">
<p align="center"> <font
color="#FF6600"><b><font size="+1">江湖拉人记录</font></b></font><font color="#CC0000" face="幼圆"><a href="javascript:this.location.reload()"><br>
  <font color="#0000FF">刷新</font></a></font> <br>
  感谢你朋友这些人是你拉到我们江湖的！<br>
  只有你拉的人存点够1000点才可以在这里显示出来！<br>
  <br>
  <font color="#0000FF">简介：拉人可以增加你自己的点数，当你所拉到我们江湖的人存点大于1000时,<br>
  你每个月都可以从他身上扒皮一定点数（是计算机自动按5%计算，并不影响所拉人的泡点），在月底的<br>
  时候是最多的。如果这一个月过去了，你将扒不到皮的，只有等一下个月了，此值不累加！所以请大家多拉人吧！</font> <br>

<table width="621" border="1" cellpadding="0" cellspacing="0" height="13" align="center" bordercolor="#000000">
  <tr> 
    <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM 用户 where allvalue>1000 and 介绍人='"& aqjh_name &"' order by -mvalue",conn
%>
    <td width="28" height="10"> 
      <div align="center"><font color="#000000">ID</font></div>
    </td>
    <td width="67" height="10"> 
      <div align="center"><font color="#000000">姓名</font></div>
    </td>
    <td width="28" height="10"> 
      <div align="center"><font color="#000000">性别</font></div>
    </td>
    <td width="57" height="10"> 
      <div align="center"><font color="#000000">门派</font></div>
    </td>
    <td width="63" height="10"> 
      <div align="center"><font color="#000000">身份</font></div>
    </td>
    <td width="125" height="10"> 
      <div align="center">最后登陆</div>
    </td>
    <td width="52" height="10"> 
      <div align="center"><font color="#FF0000">月泡点</font></div>
    </td>
    <td width="46" height="10"> 
      <div align="center"><font color="#000000">管理级</font></div>
    </td>
    <td width="37" height="10"> 
      <div align="center"><font color="#000000">战斗级</font></div>
    </td>
    <td width="42" height="10"> 
      <div align="center">扒皮点</div>
    </td>
    <td width="52" height="10"> 
      <div align="center"><font color="#000000">扒皮</font></div>
    </td>
  </tr>
  <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
  <tr> 
    <td width="28" height="30"> 
      <div align="center"><font color="#000000"><%=rs("ID")%></font></div>
    </td>
    <td width="67" height="30"> 
      <div align="center"><font color="#0000FF"><%=rs("姓名")%></font></div>
    </td>
    <td width="28" height="30"> 
      <div align="center"><font color="#000000"><%=rs("性别")%></font></div>
    </td>
    <td width="57" height="30"> 
      <div align="center"><font color="#000000"><%=rs("门派")%></font></div>
    </td>
    <td width="63" height="30"> 
      <div align="center"><font color="#000000"><%=rs("身份")%></font></div>
    </td>
    <td width="125" height="30"> 
      <div align="center"><font color="#000000"><%=rs("lasttime")%></font></div>
    </td>
    <td width="52" height="30"> 
      <div align="center"><font color="#FF0000"><%=rs("mvalue")%></font></div>
    </td>
    <td width="46" height="30"> 
      <div align="center"><font color="#000000"><%=rs("grade")%></font></div>
    </td>
    <td width="37" height="30"> 
      <div align="center"><font color="#000000"><%=rs("等级")%></font></div>
    </td>
    <td width="42" height="30"> 
      <div align="center"><font color="#FF0000"><%=int(rs("mvalue")*0.05)%></font></div>
    </td>
    <td width="52" height="30"> 
      <div align="center"><font color="#000000"> 
        <%
prevtime=CDate(rs("lasttime"))
if DateDiff("m",prevtime,sj)=0 and rs("保留2")<>"扒皮"&Month(date()) and rs("mvalue")>20 then%>
        <a href="babi.asp?id=<%=rs("id")%>"><font color="#0000FF">扒皮</font></a> 
<%else%>
	      不能操作 
<%end if%>
        </font></div>
    </td>
  </tr>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<p align="center"> <font color="#000000"><br>
拉人总数:<b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">人</font></font> 
<div align="center"></div>
</body>
</html>