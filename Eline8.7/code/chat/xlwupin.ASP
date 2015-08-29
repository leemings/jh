<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>在线拍卖♀wWw.51eline.com♀</title>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}</script>
<style>
a:hover{color:#FF9900;CURSOR:url('45.ani');}
body{
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff
}
td{font-size:9pt;}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0" topmargin="2" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<p style="font-size:12pt; color:red" align="center"> <font color="#FFFFFF"><b>在 
  线 拍 卖</b><br>
  </font><font color="#FFFFFF"><span style='font-size:9pt'><font size="3" color="#FF0000">【</font><font size="3"><%=sjjh_name%><font color="#FF3333">】<br>
  </font></font></span></font></p>
<div align="center"><font color="#efefef" size="-1">左键银两，右键金币!<br>拍卖$价钱|币制|类别|物品名|数量|说明</font></div>
<table border="1" cellspacing="0" cellpadding="0" width="145" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#b7d4f1" align="center" >
  <tr> 
    <td width="80" > 
      <div align="center">物 品 名</div>
    </td>
    <td width="30" > 
      <div align="center">数量</div>
    </td>
    <td width="30" > 
      <div align="center">自用</div>
    </td>
  </tr>
  <tr bgcolor="#85C2E0" onMouseOut="this.bgColor='#85C2E0';"onMouseOver="this.bgColor='#DFEFFF';"> 
  </tr>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w1,w2,w4,w5,w6,w7,w8,w9,w10 from 用户 where 姓名='"&sjjh_name&"'",conn
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>药 品 类</font></b></div></td>"
if not(isnull(rs("w1"))) then
	call show(rs("w1"),"使用","w1")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>毒 药 类</font></b></div></td>"
if not(isnull(rs("w2"))) then
	call show(rs("w2"),"下毒","w2")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>暗 器 类</font></b></div></td>"
if not(isnull(rs("w4"))) then
	call show(rs("w4"),"投掷","w4")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>卡 片 类</font></b></div></td>"
if not(isnull(rs("w5"))) then
	call show(rs("w5"),"用卡","w5")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>物 品 类</font></b></div></td>"
if not(isnull(rs("w6"))) then
	call show(rs("w6"),"物品","w6")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>鲜 花 类</font></b></div></td>"
if not(isnull(rs("w7"))) then
	call show(rs("w7"),"送花","w7")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>配 药 类</font></b></div></td>"
if not(isnull(rs("w8"))) then
	call show(rs("w8"),"配药","w8")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>魔 宝 类</font></b></div></td>"
if not(isnull(rs("w9"))) then
	call show(rs("w9"),"发射","w9")
end if

Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>秘 笈 类</font></b></div></td>"
if not(isnull(rs("w10"))) then
	call show(rs("w10"),"执行","w10")
end if



rs.close
set rs=nothing
conn.close
set conn=nothing
sub show(data,lx,lb)
data1=split(data,";")
data2=UBound(data1)
for y=0 to data2-1
	data3=split(data1(y),"|")
%>
  <tr bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="80"> 
      <div align="center"><%=data3(0)%> </div>
    </td>
    <td width="30"> 
      <div align="center"><%=data3(1)%> </div>
    </td>
    <td width="30"> 
	    	<div align="center"> <a href="javascript:s('/拍卖$ 1000|银两|<%=lb%>|<%=data3(0)%>|1|说明！')" oncontextmenu="javascript:s('/拍卖$ 10|金币|<%=lb%>|<%=data3(0)%>|1|说明！')"><font color="#FF0000">拍卖</font></a> 
	 </div>
    </td>
  </tr>
  <%
erase data3
next
erase data1
end sub%>
</table>
</body>
</html>