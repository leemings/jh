<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../const.asp"-->
<html>
<head>
<title>我的物品</title>
<SCRIPT LANGUAGE=javascript>
if(window.name!='wupin'){this.location.href='index.asp'}
</SCRIPT>
<script language="JavaScript">
function s(mywpname,mywpsl,myzd){
	parent.bxg.document.bxgform.filedname.value=myzd;
	parent.bxg.document.bxgform.wpname.value=mywpname;
	parent.bxg.document.bxgform.sl.value=mywpsl;
	parent.bxg.document.bxgform.sl.focus();
}
</script>
<style>
td{font-size:9pt;}
body {  font-size: 12px; color: #FFFFFF}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#000000" text="#FFFFFF" bgproperties="fixed" leftmargin="0" link="#00CCFF" vlink="#00CCFF" alink="#00CCFF">
<p align="center"><b><font size="4">我随身携带的物品</font></b><br>
  <br>
  <a href="#" onclick="javascript:parent.wupin.document.location.reload();">点击刷新物品</a></p>
<table border="1" cellspacing="0" cellpadding="0" width="145" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" >
  <tr>
    <td width="80" height="21" > 
      <div align="center">物 品 名</div>
    </td>
    <td width="30" height="21" >
      <div align="center">数量</div>
    </td>
  </tr>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w1,w2,w4,w5,w6,w7,w8 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
'药品类
Response.Write "<td colspan='3' height='15' align='center'><b>药 品 类</b></td>"
if not(isnull(rs("w1"))) then
	call show(rs("w1"),"w1")
end if
'毒药类
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15' align='center'><b>毒 药 类</b></td>"
if not(isnull(rs("w2"))) then
	call show(rs("w2"),"w2")
end if
'暗器类
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15' align='center'><b>暗 器 类</b></td>"
if not(isnull(rs("w4"))) then
	call show(rs("w4"),"w4")
end if
'卡片类
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15' align='center'><b>卡 片 类</b></td>"
if not(isnull(rs("w5"))) then
	call show(rs("w5"),"w5")
end if
'物品类
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15' align='center'><b>物 品 类</b></td>"
if not(isnull(rs("w6"))) then
	call show(rs("w6"),"w6")
end if
'鲜花类
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15' align='center'><b>鲜 花 类</b></td>"
if not(isnull(rs("w7"))) then
	call show(rs("w7"),"w7")
end if
'配药类
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15' align='center'><b>配 药 类</b></td>"
if not(isnull(rs("w8"))) then
	call show(rs("w8"),"w8")
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
'取物品数量子程序
sub show(data,lb)
	data1=split(data,";")
	data2=UBound(data1)
	for y=0 to data2-1
		data3=split(data1(y),"|")%>
  <tr onmouseout="this.bgColor='#000000';"onmouseover="this.bgColor='#D00000';"> 
    <td width="80" align="center"><a href="javascript:s('<%=data3(0)%>',<%=data3(1)%>,'<%=lb%>')"><%=data3(0)%></a></td>
    <td width="30" align="center"><%=data3(1)%></td>
  </tr>
	  <%erase data3
	next
	erase data1
end sub%>
</table>
</body>
</html>