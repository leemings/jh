<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if Weekday(date())<>6 or (Hour(time())<21 or Hour(time())>=22) then
Response.Write "<script Language=Javascript>alert('提示：门派大战只可以在周五21-22点进行！并且只有堂主以上才有资格，请提前和掌门准备好！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
rs.open "SELECT 门派 FROM 用户 where 姓名='" & aqjh_name &"'",conn
pai=rs("门派")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title><%=pai%>武功设置</title>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();
if (liji==true) {parent.f2.document.af.subsay.click()};
}
</script>
<style>
body{
CURSOR: url('3.ani');
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699" leftmargin="0" topmargin="0">
<table border="1" width="100%" cellspacing="0" cellpadding="0" bgcolor="#66CCCC" height="10" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
        <td width="100%" height="28"> <div align="center"><font color="#FF0000" size="2"><strong>门派大战</strong></font></div></td>
      </tr>
    </table>
  
<table cellpadding="2" cellspacing="0" border="1" bgcolor="#33CCCC" align="center" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr valign="middle"> 
      <td width="43%"> 
        <div align="center"><font color="#330033" size="2">招式</font></div>
      </td>
      <td width="29%"> 
        <div align="center"><font size="2" color="#330033">武功</font></div>
      </td>
      <td width="28%"> 
        <div align="center"><font size="2" color="#330033">内力</font></div>
      </td>
    </tr>
    <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM y where b='" & pai & "' order by e",conn
do while not rs.eof and not rs.bof
dj=int(sqr((rs("e")/100)))+1
%>
    <tr valign="middle"> 
      <td width="43%"> 
        <div align="center"> <font color="#330033" size="2"> <a href="javascript:s('/门派大战$ <%=rs("a")%>')" title="此招试为：<%=dj%>重!"><%=rs("a")%></a></font></div>
      </td>
      <td width="29%"> 
        <div align="center"><font color="#330033" size="2"><%=rs("c")%></font></div>
      </td>
      <td width="28%"> 
        <div align="center"><font color="#330033" size="2"><%=rs("d")%></font></div>
      </td>
    </tr>
    <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing%>
  </table>
  
<table border="1" width="100%" cellspacing="0" cellpadding="0" bgcolor="#33CCCC" height="10" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
        <td width="100%" height="28"><div align="center"><img src="go.gif" width="138"></div></td>
      </tr>
    </table>
<p class=p1 align="center"><font color="#FFFFFF" size="2">杀伤力等于<br>
所用内力+攻击力<br>
<input type="checkbox" name="liji" value="on">
立即攻击 </font></p>
</body>
</html>
