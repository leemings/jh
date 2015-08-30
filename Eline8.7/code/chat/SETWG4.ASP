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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")

pai=sjjh_name

%>
<html>
<head>
<title><%=pai%>轩辕招式♀wWw.happyjh.com♀</title>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();
if (liji==true) {parent.f2.document.af.subsay.click()};
}
</script>
<style type="text/css">
td           { font-family: 宋体; font-size: 12px }
body         { font-family: 宋体; font-size: 12px;
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699" leftmargin="0" topmargin="0">
<div align="center"><font color="#FFFFFF"><span style='font-size:9pt'><font size="-1">〖</font></span><font size="-1"><%=pai%>武功<span style='font-size:9pt'>〗</span></font></font><font size="3"><br>
</font><br>
  <table cellpadding="2" cellspacing="0" border="1" bgcolor="#0099CC" align="center" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr valign="middle"> 
      <td width="43%"> <div align="center"><font color="#330033" size="2">招式</font></div></td>
      <td width="29%"> <div align="center"><font size="2" color="#330033">武功</font></div></td>
      <td width="28%"> <div align="center"><font size="2" color="#330033">内力</font></div></td>
    </tr>
    <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM n where b='" & pai & "' order by e",conn
do while not rs.eof and not rs.bof
dj=int(sqr((rs("e")/100)))+1
%>
    <tr valign="middle"> 
      <td width="43%"> <div align="center"> <font color="#330033" size="2"> <a href="javascript:s('/轩辕$ <%=rs("a")%>')" title="此招试为：<%=dj%>重!"><%=rs("a")%></a></font></div></td>
      <td width="29%"> <div align="center"><font color="#330033" size="2"><%=rs("c")%></font></div></td>
      <td width="28%"> <div align="center"><font color="#330033" size="2"><%=rs("d")%></font></div></td>
    </tr>
	<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing%>
        </table>
</td>
    </tr>
    
  </table>
  <table width="100%" height="20" border="1" align="center" cellpadding="3" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#0099df">
    <tr> 
            <td><div align="center"><a href="xx/SETWG.ASP" target="_blank"><font color="#990000" size="2">添加|修改|修炼</font></a></div></td>
          </tr>
        </table>
</div>
<p class=p1 align="center"><font color="#FFFFFF" size="2">杀伤力等于<br>
所用内力+攻击力<br>
<input type="checkbox" name="liji" id="liji" value="on">
立即攻击 </font></p>
</body>
</html>
