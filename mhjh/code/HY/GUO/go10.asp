<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=session("yx8_mhjh_username")
if Session("usepro")<>"5" then
Session("usepro")=""
response.redirect "index.asp"
end if
%>
<% 
if name="" then %>
<script language=vbscript>
MsgBox "�Բ����㻹û�е�¼���߳�ʱ"
location.href = "<a href="javascript:self.close()">"
</script>
<%end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="update �û� set ����=����+10000000,����=����+10000,����=����+500000,����=����+1000000 where ����='"&name&"'"
conn.Execute(sql)
Session("usepro")=""
conn.Close    
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
		newtalkarr(594)=name 
		newtalkarr(595)="���" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>������ն����</font><font color=blue>"&name&"����ǧ�����,���ڴ�����5��,������һ���������ֵ�����,������˾޶��!</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
%>
<% 
if datediff("s",session("refresh"),now())<50000 then 
response.write "�ܾ�����ˢ��!С��ɾ���ʺţ�" 
response.end 
else 
 session("refresh")=now()  

end if 
%>
<html>

<head>
<title>������ȡ���˶�����ܵı���</title>
<LINK href="../../style.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body background="bg.gif" leftmargin="0" oncontextmenu=self.event.returnValue=false >
<div align="center">
  <table border="0" width="600">
    <tr> 
      <td width="100%"> <p align="left" style="line-height: 200%; margin: 20"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
          <font color="#FFFF00">���ڵ��������ܵı����ˣ����ǵ����������𿪱��ص����ǣ�ȴ���ֲ�û��ʲô�����޵е������ؼ���ֻ��ɽ����ʯ�������ɷ����д������һ���֣�</font></b></td>
    </tr>
    <tr> 
      <td width="100%"><img border="0" src="finish.gif"></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td width="100%" valign="top"> 
        <p align="center" style="line-height: 200%; margin: 20"><b><font color="#000000">&nbsp;&nbsp;<font color="#FFFF00"> 
          ԭ�������������δ���һ�ܵ�ԭ������������������ţ�ԭ�����ǵ����������ž��������޵е�</font></font></b><font color="#FFFF00"><b>�ؾ�ѽ������ɽ�����ҵ�</b></font><b><font color="#FFFFFF">1000000</font><font color="#FFFF00">�����ӣ���ʯ���϶�������漣����������������</font><font color="#FFFFFF">50000</font><font color="#FFFF00">�㣬�����ͷ���������<font color="#000000">��</font></font><font color="#FFFFFF">10000</font><font color="#FFFF00">�㣡��ϲ������ϲ��������</font></b></p>
        </td>
    </tr>
  </table> 
</div> 
</html> 
