<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%
id=request("id")
if instr(id,"官")<>0 then
		Response.Write "<script language=JavaScript>{alert('严重警告，不要搞乱');window.close();}</script>"
		Response.End
end if
my=request("my")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css><!--
.text {  font-size: 12px; color: #FFFFFF; text-decoration: none}
-->
</style>
<link rel="stylesheet" href="../chat/READONLY/Style.css">
<title>门派聚议厅♀wWw.happyjh.com♀</title>
</head>
<body leftmargin="0" topmargin="5" marginwidth="0" marginheight="0" bgcolor="#FFFFFF" background="../chat/lvbgcolor.gif" oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false>
<div align="center">
<b><font size="+2" color="#008080">门派聚议厅</font></b><br>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT 门派 FROM 用户 where 姓名='"&sjjh_name&"'",conn

gongsi=rs("门派")

rs.close
set rs=nothing
conn.close
set conn=nothing
%>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 where 门派='"&gongsi&"' and 身份<>'堂主' and 身份<>'护法' and 身份<>'长老' and 身份<>'掌门' order by 门派基金 DESC",conn
%>
</div>

<table border="1" width="80%" cellspacing="0" cellpadding="3" height="0" align="center" bordercolordark="#FFFFFF" bordercolorlight="#000000" bgcolor="#FFFFFF">
  <tr align="center"> 
    <td height="20" colspan="6"> 
      <table border="1" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFCCCC" align="center">
        <tr> 
          <td width="25%" valign="middle" align="center"><a href="bentop4.asp"><font color="#0000ff">本门弟子排行</font></a></td>
          <td width="25%" valign="middle" align="center"><a href="bentop1.asp"><font color="#ff0000">本门堂主排行</font></a></td>
          <td width="25%" valign="middle" align="center"><a href="bentop2.asp"><font color="#0000FF">本门护法排行</font></a></td>
          <td width="25%" valign="middle" align="center"><a href="bentop3.asp"><font color="#0000FF">本门长老排行</font></a></td>
        </tr>
      
	 </table>
<tr align="center"> 
    <td height="16" width="18%" bgcolor="#669900"><font color="#FFFFFF">姓名</font></td>
    <td height="16" width="16%" bgcolor="#669900"><font color="#FFFFFF">内力</font></td>
    <td height="16" width="16%" bgcolor="#669900"><font color="#FFFFFF">体力</font></td>
    <td height="16" width="16%" bgcolor="#669900"><font color="#FFFFFF">积分</font></td>
    <td height="16" width="16%" bgcolor="#669900"><font color="#FFFFFF">战斗等级</font></td>
    <td height="16" width="18%" bgcolor="#669900"><font color="#FFFFFF">门派基金</font></td>

<%do while not rs.bof and not rs.eof%>
  <tr align="center"> 
    <td height="19"><font size="2"><%=rs("姓名")%></font></td>
    <td height="19"><font size="2"><%=rs("内力")%></font></td>
    <td height="19"><font size="2"><%=rs("体力")%></font></td>
    <td height="19"><font size="2"><%=rs("mvalue")%></font></td>
    <td height="19"><font size="2"><%=rs("等级")%></font></td>
    <td height="19"><font size="2"><%=rs("门派基金")%></font></td>
<tr>

	<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>

<table border="1" width="80%" cellspacing="0" cellpadding="3" height="0" align="center" bordercolordark="#FFFFFF" bordercolorlight="#000000" bgcolor="#FFFFFF">
  <tr bgcolor="#669900"> 
    <td colspan="3"><b><font color="#FFFFFF">门派身份申请</font></b></td>
  </tr>
  <tr> 
    <td rowspan="2" width="19%" class="text" bgcolor="#008080"><b><font size="2" color="#FFFF00">长老(8个)</font></b></td>
    <td colspan="2" class="text" bgcolor="#008080"><font color="#FFFFFF"><b>申请条件</b>：<b> 
      <font color="#FF9900"><font color="#DD2222"> </font></font></b></font><b> 
      <font color="#FFFF00"> 
      战斗等级</font></b><font color="#FFFFFF">25级以上，</font><b><font color="#FFFF00">门派基金</font></b><font color="#FFFFFF">30000000点以上</font></td>
  </tr>
  <tr> 
    <td width="68%" class="text" bgcolor="#008080"><b><font color="#FFFFFF">拥有权利</font></b><font color="#FFFFFF">：<b> 
      </b>本派刑法，本派罚款，掌门令，奖励弟子<b>，</b>教训弟子，招收弟子</font></td> 
    <td width="13%" class="text" bgcolor="#008080"><b><font color="#FFFFFF"><a href="zongcai.asp" target="_blank"></a></font><a href="zongcai.asp" target="_blank"><font color="#FF0000">我要申请</font></a><font color="#FFFFFF"></font></b></td>
  </tr>
  <tr> 
    <td rowspan="2" width="19%" class="text" bgcolor="#008080"><b><font size="2" color="#FFFF00">护法(10个)</font></b></td>
    <td colspan="2" class="text" bgcolor="#008080"><font color="#FFFFFF"><b>申请条件</b>： 
      </font> <b><font color="#FFFF00">战斗等级</font></b><font color="#FFFFFF">20以上，</font><b><font color="#FFFF00">门派基金</font></b><font color="#FFFFFF">20000000点以上</font></td>
  </tr>
  <tr> 
    <td width="68%" class="text" bgcolor="#008080"> 
      <p><b><font color="#FFFFFF">拥有权利</font></b><font color="#FFFFFF">： 本派刑法，本派罚款，奖励弟子</font></p> 
    </td>
    <td width="13%" class="text" bgcolor="#008080"><b><font color="#FFFFFF"><a href="mishu.asp" target="_blank"></a></font><a href="mishu.asp" target="_blank"><font color="#FF0000">我要申请</font></a><font color="#FFFFFF"></font></b></td>
  </tr>
  <tr> 
    <td width="19%" rowspan="2" class="text" bgcolor="#008080"><b><font size="2" color="#FFFF00">堂主(16个)</font></b></td>
    <td colspan="2" class="text" bgcolor="#008080"><font color="#FFFFFF"><b>申请条件</b>： 
      </font> <b><font color="#FFFF00">战斗等级</font></b><font color="#FFFFFF">15级以上，</font><b><font color="#FFFF00">门派基金</font></b><font color="#FFFFFF">15000000点以上</font></td>
  </tr>
  <tr> 
    <td width="68%" class="text" bgcolor="#008080"><b><font color="#FFFFFF">拥有权利</font></b><font color="#FFFFFF">： 
      本派刑法，本派罚款</font></td>
    <td width="13%" class="text" bgcolor="#008080"> 
      <div align="left"><b><font color="#FFFFFF"><a href="jingli.asp" target="_blank"></a></font><a href="jingli.asp" target="_blank"><font color="#FF0000">我要申请</font></a><font color="#FFFFFF"></font></b></div>
    </td>
  </tr>
</table>
    <p align="center"><font color="#3366FF">『快乐江湖』&#8482;</font></p> 


