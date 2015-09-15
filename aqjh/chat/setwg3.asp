<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
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
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" background="<%=chatimage%>" bgproperties="fixed">
<center><font color=ffffff>〖<%=pai%>武功〗</font>
  <table cellpadding="3" cellspacing="0" border="1" bgcolor="#CCCCCC" align="center" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr valign="middle"> 
      <td width="43%"> 
        <div align="center"><font color="#330033">招式</font></div>
      </td>
      <td width="29%"> 
        <div align="center"><font color="#330033">武功</font></div>
      </td>
      <td width="28%"> 
        <div align="center"><font color="#330033">内力</font></div>
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
        <div align="center"> <font color="#330033"> <a href="javascript:s('/发招$ <%=rs("a")%>')" title="此招试为：<%=dj%>重!"><%=rs("a")%></a></font></div>
      </td>
      <td width="29%"> 
        <div align="center"><font color="#330033"><%=rs("c")%></font></div>
      </td>
      <td width="28%"> 
        <div align="center"><font color="#330033"><%=rs("d")%></font></div>
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
</center>
<p class=p1 align="center"><font color="#FFFFFF">杀伤力等于<br>
所用内力+攻击力<br>
<input type="checkbox" name="liji" value="on">
立即攻击 </font></p>
</body>
</html>