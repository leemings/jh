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
%>
<html>
<head>
<title>会员列表</title>
<LINK href=css/css.css type=text/css rel=stylesheet>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"> <font color="#CC0000"><a href="javascript:this.location.reload()">刷新</a> 
  <br>
感谢这些朋友对我们江湖的大力支持！加入会员点<a href="../chat/hy.asp" target="_blank">这里</a><br>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM 用户 where 状态='正常' and 会员等级>0 order by 会员日期 ",conn
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
            <div align="center">会员级</div>
          </td>
          <td width="88" height="10"> 
            <div align="center">会员结束日期</div>
          </td>
          <%if show<>"" then%>
          <td width="73" height="10"> 
            <div align="center">江湖等级</div>
          </td>
          <%end if%>
          <td width="35" height="10"> 
            <div align="center">登录</div>
          </td>
        </tr>
<%
dengji1=0
dengji2=0
dengji3=0
dengji4=0
dengji5=0
do while not rs.bof and not rs.eof
Select Case rs("会员等级")
Case 1
	dengji1=dengji1+1
case 2
	dengji2=dengji2+1
case 3
	dengji3=dengji3+1
case 4
	dengji4=dengji4+1
case 5
	dengji5=dengji5+1
case 6
	dengji6=dengji6+1
case 7
	dengji7=dengji7+1
case 8
	dengji8=dengji8+1
end select
%>
        <tr height="20"> 
          <td width="28"> 
            <div align="center"><%=rs("ID")%></div>
          </td>
          <td width="47"> 
            <div align="center"><a href="SHOWUSER.asp?username=<%=rs("姓名")%>"><font color="#FF9900"><%=rs("姓名")%></a></div>
          </td>
          <td width="25"> 
            <div align="center"><%=rs("性别")%></div>
          </td>
          <td width="63"> 
            <div align="center"><%=rs("门派")%></div>
          </td>
          <td width="54""> 
            <div align="center"><%=rs("身份")%></div>
          </td>
          <td width="82"> 
            <div align="center"><%=rs("会员等级")%></div>
          </td>
          <td width="88"> 
            <div align="center"><%=rs("会员日期")%></div>
          </td>
          <%if show<>"" then%>
          <td width="73" height="30"> 
            <div align="center"><%=rs("等级")%></div>
          </td>
          <%end if%>
          <td> 
            <div align="center"><%=rs("times")%></div>
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
<div align="center">一级会员：<%=dengji1%> 二级会员:<%=dengji2%> 三级会员：<%=dengji3%> 四级会员：<%=dengji4%> 五级会员：<%=dengji5%> 六级会员:<%=dengji6%> 七级会员：<%=dengji7%> 八级会员：<%=dengji8%><br>
会员总数：<%=(dengji1+dengji2+dengji3+dengji4+dengji5+dengji6+dengji7+dengji8)%>人</div></body></html>
