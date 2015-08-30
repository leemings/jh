<%@ LANGUAGE=VBScript codepage ="936" %>

<%Response.Expires=0
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
<title>夺宝专用♀一线网络→wWw.happyjh.com♀</title>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();
if (lijigongji==true) {parent.f2.document.af.subsay.click()};
}</script>
<style>
body{
CURSOR: url('44.ani');
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
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#006699" bgproperties="fixed" leftmargin="0" topmargin="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table border="1" width="100%" cellspacing="0" cellpadding="0" bgcolor="#0099CC" height="10" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
        <td width="100%" height="28"> <div align="center"><font color="#FF0000" size="2"><strong>夺宝专用</strong></font></div></td>
      </tr>
    </table>
<div align="center"><input type="checkbox" name="liji" id="liji" value="on"><font color="#Ffffff" size="2">立即使用</font></div>
<table border="1" cellspacing="0" cellpadding="0" width="145" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#b7d4f1" align="center" >
  <tr> 
    <td width="80" > 
      <div align="center">物品名</div>
    </td>
    <td width="30" > 
      <div align="center">数量</div>
    </td>
    <td width="30" > 
      <div align="center">使用</div>
    </td>
  </tr>
  <tr bgcolor="#85C2E0" onMouseOut="this.bgColor='#85C2E0';"onMouseOver="this.bgColor='#DFEFFF';"> 
  </tr>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w1,w2,w4,w5,w6,w7,w8,w9,w10 from 用户 where 姓名='"&sjjh_name&"'",conn
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>毒 药 类</font></b></div></td>"
if not(isnull(rs("w2"))) then
	call show(rs("w2"),"夺宝下毒","w2")
end if
Response.Write "<tr></tr><tr></tr>"
Response.Write "<td colspan='3' height='15'><div align='center'><b><font color='#0000FF'>暗 器 类</font></b></div></td>"
if not(isnull(rs("w4"))) then
	call show(rs("w4"),"夺宝投掷","w4")
end if
Response.Write "<tr></tr><tr></tr>"

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
          <div align="center">
          <a href="javascript:s('/丢弃$ <%=lb%>|<%=data3(0)%>|1')" oncontextmenu="javascript:s('/丢弃$ <%=lb%>|<%=data3(0)%>|<%=data3(1)%>')"><font color=blue><%=data3(0)%></font></a>
           </div>
    </td>
    <td width="30"> 
      <div align="center"><%=data3(1)%> </div>
    </td>
    <td width="30"> 
		<%if lx="用卡" or lx="配药" then%>
		    <div align="center"> <a href="javascript:s('/<%=lx%>$ <%=data3(0)%>')" oncontextmenu="javascript:s('/赠送$ <%=lb%>|<%=data3(0)%>|1')"><font color="#FF0000"><%=lx%></font></a> 
		<%elseif lx="物品" then%>
	    	<div align="center"> <a href="javascript:s('/赠送$ <%=lb%>|<%=data3(0)%>|1')" oncontextmenu="javascript:s('/赠送$ <%=lb%>|<%=data3(0)%>|1')"><font color="#FF0000">赠送</font></a> 
		<%else%>
	    	<div align="center"> <a href="javascript:s('/<%=lx%>$ <%=data3(0)%>&1')" oncontextmenu="javascript:s('/赠送$ <%=lb%>|<%=data3(0)%>|1')"><font color="#FF0000"><%=lx%></font></a> 
		<%end if%>
      </div>
    </td>
  </tr>
  <%
erase data3
next
erase data1
end sub%>
</table><br>
  <div align="center"> <a href=javascript:history.go(-1)><font color=#00FF00 size=-1>返回上级</font></a></div>

</body>
</html>








