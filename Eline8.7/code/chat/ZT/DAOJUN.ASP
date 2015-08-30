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
rs.open "SELECT 攻击,防御,w3 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3
mygj=rs("攻击")
myfy=rs("防御")
%>
<html>
<head>
<title>♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="img/style.css">
</head>
<body  topmargin="0" leftmargin="0" bgcolor="#CCCCCC" oncontextmenu=self.event.returnValue=false>
<table width="256" border="0" cellspacing="0" cellpadding="0" height="224" align="left">
  <tr> 
    <td width="256" rowspan="3" valign="top"> <div align="center">
        <p align="center"><img src="img/dao.gif" width="40" height="15">装备物品一览<img src="img/dao1.gif" width="40" height="15"> 
        <font color="#CC0000" face="幼圆"><a href="javascript:this.location.reload()">刷新</a></font></div>   
      <div align="center" style="width: 256; height: 67"> 
        <table width="255" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
          <tr> 
            <td width="45"> <div align="center">物品名</div></td>
            <td width="28"> <div align="center">类型</div></td>
            <td width="24"> <div align="center">数量</div></td>
            <td width="32"> <div align="center">防御</div></td>
            <td width="31"> <div align="center">攻击</div></td>
            <td width="25"> <div align="center">耐久</div></td>
            <td width="24"> <div align="center">特性</div></td>
            <td width="28"> <div align="center">操作</div></td>
          </tr>
          <%
if not(isnull(rs("w3"))) then
data1=split(rs("w3"),";")
data2=UBound(data1)
for y=0 to data2-1
	data3=split(data1(y),"|")
	rs.close
rs.open "select * FROM b WHERE a='" & data3(0) &"'",conn,3,3
If not(Rs.Bof and Rs.Eof) then
%>
          <tr align="center"> 
            <td width="45" height="11"><img src='../../HCJS/jhjs/images/<%=rs("i")%>' alt='<%=rs("c")%>'></td>
            <td width="28" height="11"><%=rs("b")%></td>
            <td width="24" height="11"> <%=data3(1)%> </td>
            <td width="32" height="11"><%=rs("g")%></td>
            <td width="31" height="11"><%=rs("f")%></td>
            <td width="25"><%=rs("j")%></td>
            <td width="24" height="11"><%=rs("k")%></td>
            <td width="28" height="11"> <a href="wupin2.asp?name=<%=data3(0)%>"><font color="#0000FF">装备</font></a> 
            </td>
          </tr>
          <%
end if
erase data3
next
erase data1
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
        </table>
        注意：点装备安装此类物品,鼠标放在图片上面有说明！</div></td>
  </tr>
</table>
</body>
</html>
