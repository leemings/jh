<!--#include file="data.asp"-->
<!--#include file="dalie.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
nickname=session("yx8_mhjh_username")
chatroomsn=session("yx8_mhjh_userchatroomsn")
sql="select * from �û� where ����='"&nickname&"'"
set rs=conn.execute(sql)
if rs("����")<>"�ٸ�"  then
mess="��û�з������Ȩ�ޡ�"
else
sql1="update ���� set dalie='�ϻ�'"
connt.execute sql1
mess="���Գɹ�"
		dim newtalkarr(600) 
        talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username
		newtalkarr(595)="���" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>����Ϣ��</font><b><font color=red>ͻȻһֻҰ��<img src=laohu.gif>���뽭�������ˣ�������ǿ�ȥ���԰���</font></b>" 
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
end if
%>
<html>

<head>
<title>����</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
-->
</style>

</head>

<body bgcolor="#CCCCCC">
<p>&nbsp;</p>
<table width="98%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="17" bgcolor="#996633" align="center"><font color="#FFFFFF">��Ϣ</font></td>
  </tr>
  <tr> 
    <td bgcolor="#999966" align="center" height="157"><font>  <%=mess%></font></td>
  </tr>
</table>
<p>&nbsp;</p>

</body> 
</html> 
