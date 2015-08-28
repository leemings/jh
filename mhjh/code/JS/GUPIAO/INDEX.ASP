<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
%>

<head>
<title>股票交易中心</title>
<script language=javascript>function selectstock(stock){parent.stockfrm.location.replace('showstock.asp?stock='+stock);parent.infofrm.location.replace('info.asp?stock='+stock);}setTimeout('location.reload();',180000);</script>
<link rel="stylesheet" href="new.CSS" type="text/css">
</head>
<body link="#0000FF" vlink="#0000FF" alink="#0000FF" bgcolor="#523E3A">
<table border="0" width="100%" height="41" cellpadding="0" cellspacing="0">
  <tr>
    <td width="20%" height="16">
      <p align="center"><font color="#FFFFFF" face="隶书" size="6">爱火神话</font></td>
    <td width="1%" height="16">
      <p align="center"><img border="0" src="../../image/laugh2.gif"></td>
    <td width="20%" height="16">
      <p align="center"><font color="#FFFFFF" face="隶书" size="6">再战江湖</font></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr> 
<td> 
<table width="100%" border="0" cellspacing="0" bgcolor="#000000" cellpadding="0">
<tr bgcolor="whitesmoke"> 
<td width="10%" height="25"> 
<div align="center">股票代号</div>
</td>
<td width="20%" height="25"> 
<div align="center">股票名称</div>
</td>
<td width="20%" height="25"> 
<div align="center">流通股总量</div>
</td>
<td width="10%" height="25"> 
<div align="center">发行价</div>
</td>
<td width="10%" height="25"> 
<div align="center">当前价格</div>
</td>
<td width="10%" height="25"> 
<div align="center">历史高价</div>
</td>
<td width="10%" height="25"> 
<div align="center">历史低价</div>
</td>
<td width="10%" height="25"> 
<div align="center">升降</div>
</td>
</tr>
</table>
<%
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
rs.Open "select * from 股票",conn,2,2
do while not(rs.EOF or rs.BOF)
increa =(rs("现价")/rs("最近成交价")-1)*100
if increa>0 then 
increa ="<font color=0000FF>"&formatnumber(increa,2,-1)&"</font>"
elseif increa<0 then
increa ="<font color=FF0000>"&formatnumber(increa,2,-1)&"</font>"
else
increa ="<font color=000000>"&formatnumber(increa,2,-1)&"</font>"
end if 
%>
      <table width="100%" border="0" cellpadding="2" cellspacing="1">
        <tr bgcolor="#339966" onmouseout="this.bgColor='#339966';"onmouseover="this.bgColor='#DFEFFF';"> 
<td width="10%"><div align="center"><%=rs("id")%> </div> 
</td>
<td width="20%"> 
            <div align="center"> <a href="#" onClick="selectstock('<%=rs("股票名称")%>');"><%=rs("股票名称")%></a></div>
</td>
<td width="20%"><div align="center"><%=rs("流通股总量")%> 
</div>
</td>
<td width="10%"><div align="center"><%=rs("发行价")%> 
</div>
</td>
<td width="10%"><div align="center"><%=rs("现价")%> 
</div>
</td>
<td width="10%"><div align="center"><%=rs("流通股总量")%> 
</div>
</td>
<td width="10%"><div align="center"><%=rs("历史低价")%> 
</div>
</td>
<td width="10%"><div align="center"><%=increa%> 
</div>
</td>
</tr>
</table>
<%
rs.MoveNext
loop
%>
</td>
</tr>
</table>
<% 
rs.close
conn.close
set rs=nothing
set conn=nothing
%>