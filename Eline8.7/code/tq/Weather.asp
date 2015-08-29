<%
'***********************************************************************************

' 文件名.........: Weather.asp
' 作者...........: flanker
' 说明...........: 自动获取天气预报
' 注意...........: 
' 版权...........: Copyright (c) 2003, www.su27.net.
' 修改记录.......: 时间         人员     备注
'                  ---------    -------  -------------------------------------------
'                  2003-01-21   flanker   创建文件

'***********************************************************************************
%>
<!--#include file="lib/pub_Config.asp"-->
<%
shen   = Trim(Request("shen"))
Action = Trim(Request("Action"))
Select Case(Action)
	Case("")
	ShowList()                                 '显示列表
	Case("1")
	ShowDetail()                               '显示详细
end select

'-----------------------------显示列表----------------------------
Function ShowList()
Response.Write "<title>全国天气预报</title>"
set ShowRs = conn.execute("Select * from Part1 order by id")
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function changeshen(){
location.href = "Weather.asp?shen="+document.all("shen").value;
}
function openpop(){
window.open('Weather.asp?action=1&shi='+document.all("shi").value,'_blank','width=400,height=400,toolbar=no,scrollbars=no');
}
//-->
</SCRIPT>
<body bgcolor="#006699" leftmargin="5" topmargin="0" oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false>
<div align="center">
  <center>
    <table width="100%" border="1" cellpadding="2" cellspacing="2" bordercolor="#111111" bgcolor="#FFFFFF" id="AutoNumber1" style="border-collapse: collapse">
      <tr>
      <td width="100%">
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
        <tr>
              <td width="100%" align="center" height="47"> <img border="0" src="images/weather.gif" width="140" height="47"></td>
        </tr>
        <tr>
              <td width="100%" align="center" height="35" colspan=2>
<select id="shen" size="1" name="shen" onchange="javascript:changeshen();" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 12px; BACKGROUND: #666666; BORDER-LEFT: #000000 1px solid; COLOR: #f6e700; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 16px;font-size:12px">
		  <%
		  if shen <> "" then Response.Write "<option value='"&shen&"' selected>"&GetShenname(shen)&"</option>"
		  do while not ShowRs.eof
			Response.Write "<option value='"&ShowRs("id")&"'>"&ShowRs("name")&"</option>"
			showRs.movenext
		  loop
		  Set ShowRs = nothing
		  %>
          </select></td>
        </tr>
        <tr>
              <td width="100%" align="center" height="34">
<select size="1" id="shi" name="shi" style="BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 12px; BACKGROUND: #666666; BORDER-LEFT: #000000 1px solid; COLOR: #f6e700; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 16px;font-size:12px">
		  <%
		  if shen = "" then Shen = 1
		  Set showRs = conn.execute("Select * from part2 where part1="&shen&" order by id")
		  if ShowRs.eof then Response.Write "<option value='0'>-----------</option>"
		  do while not ShowRs.eof
			  Response.Write "<option value='"&ShowRs("value")&"'>"&ShowRs("name")&"</option>"
			  showRs.movenext
		  loop
		  Set ShowRs = nothing
		  %>
          </select>&nbsp;<input type="button" value="查看" name="B1" style="BACKGROUND-COLOR: #009933; BORDER-BOTTOM: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; CURSOR: hand; FONT-FAMILY: '宋体'; COLOR: #ffffff; FONT-SIZE: 9pt; HEIGHT: 20px" onclick="javascript:openpop()"></td>
        </tr>
		<tr><td><div align="center"><img src="images/Z16.gif"></div></td></tr>
     <tr>
              <td height="20">
<div align="center"><strong><font color="#009933">全国天气预报</font></strong><font color="#009933"></font></div></td></tr>
	  </table>
      </td>
    </tr>
  </table>
  </center>
</div></body>
<%
end Function

'---------------------------显示详细-------------------------------
Function ShowDetail()
	shi = Trim(Request("shi"))
	if Shi = "" or shi = "0" then
		Response.Write "缺少必须参数！"
		Response.end
	end if
	Response.Write "<title>"&GetShiName(shi)&"</title>"
	if left(shi,2) = "TW" then
		myurl	= "http://cn.weather.yahoo.com/TWXX/"&Shi&"/index_c.html"
	else
		myurl	= "http://cn.weather.yahoo.com/CHXX/"&Shi&"/index_c.html"
	end if
	Content	= getHTTPPage(myurl)
	if Content	= "" or len(Content) < 500 then ShowError()
	Temp1	= instr(Content,"今天天气")
	Temp3	= instr(Content,"<td valign=top width=100")
	if Temp1 = 0 and Temp3 = 0 then showError()
	if Temp1 > 0 then	
		Temp2	= instr(Temp1,Content,"<td valign=top align=center>")
		if Temp1 =0 or Temp1 < 0 or Temp2 = 0 or Temp2 < 0  or Temp2 < Temp1 then showError()
		Content	= mid(Content,Temp1,Temp2-Temp1)
	else
		Temp2	= instr(Temp3,Content,"<td valign=top nowrap")
		if Temp3 =0 or Temp3 < 0 or Temp2 = 0 or Temp2 < 0  or Temp2 < Temp3 then showError()
		Content	= Replace(mid(Content,Temp3,Temp2-Temp3),"<td valign=top nowrap","")
	end if
	Response.Write content
	Response.Write "<br><div align=center><input type=button value='关 闭' class=stbtm onclick='javascript:window.close();'></div>"
end Function
%>