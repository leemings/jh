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
if Weekday(date())<>6 or (Hour(time())<21 or Hour(time())>=22) then
Response.Write "<script Language=Javascript>alert('��ʾ�����ɴ�սֻ����������21-22����У�����ֻ���������ϲ����ʸ�����ǰ������׼���ã�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
rs.open "SELECT ���� FROM �û� where ����='" & sjjh_name &"'",conn
pai=rs("����")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title><%=pai%>�书���á�wWw.happyjh.com��</title>
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
        <td width="100%" height="28"> <div align="center"><font color="#FF0000" size="2"><strong>���ɴ�ս</strong></font></div></td>
      </tr>
    </table>
  
<table cellpadding="2" cellspacing="0" border="1" bgcolor="#33CCCC" align="center" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr valign="middle"> 
      <td width="43%"> 
        <div align="center"><font color="#330033" size="2">��ʽ</font></div>
      </td>
      <td width="29%"> 
        <div align="center"><font size="2" color="#330033">�书</font></div>
      </td>
      <td width="28%"> 
        <div align="center"><font size="2" color="#330033">����</font></div>
      </td>
    </tr>
    <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM y where b='" & pai & "' order by e",conn
do while not rs.eof and not rs.bof
dj=int(sqr((rs("e")/100)))+1
%>
    <tr valign="middle"> 
      <td width="43%"> 
        <div align="center"> <font color="#330033" size="2"> <a href="javascript:s('/���ɴ�ս$ <%=rs("a")%>')" title="������Ϊ��<%=dj%>��!"><%=rs("a")%></a></font></div>
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
<p class=p1 align="center"><font color="#FFFFFF" size="2">ɱ��������<br>
��������+������<br>
<input type="checkbox" name="liji" value="on">
�������� </font></p>
</body>
</html>
