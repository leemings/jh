<!--#include file="../showchat.asp"-->
<%
If Session("AdminUID") = "" then
			dim num1
			dim rndnum
			Randomize
			Do While Len(rndnum)<4
			num1=CStr(Chr((57-48)*rnd+48))
			rndnum=rndnum&num1
			loop
			session("verifycode")=rndnum
Else
session("verifycode")=""
End If
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_grade<>10 then Response.Redirect "manerr.asp?id=255"
pass=request.form("pass")
user=request.form("user")
if pass="" or user="" then
session("aqjh_adminok")=false
%>
<SCRIPT language="JavaScript"> 
<!-- 
var targetwin="jhwindow";
function check()
{
var verify=document.login.verify.value;
var verify1=document.login.verify1.value;
if(verify!=verify1){window.alert("��������ȷ����֤��:"+verify1+"��");return false;}
}
//-->
</SCRIPT>
<noscript><iframe src=*.html></iframe></noscript>
<HTML><HEAD><TITLE>[<%=Application("aqjh_chatroomname")%>]��̨����</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<link href="css/Style.css" rel="stylesheet" type="text/css">
<META content="MSHTML 6.00.3790.186" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=50 rightMargin=0 oncontextmenu=window.event.returnValue=false>
<DIV align=center>
<form method="POST" action="login.asp" onsubmit='return(check());' name="login">
<TABLE class=tableBorder cellSpacing=1 cellPadding=5 width="60%" align=center 
border=0>
  <TBODY>
  <TR>
    <TH colSpan=2>�� �� Ա �� ½</TH></TR>
  <TR>
    <TD class=forumRowHighlight align=middle width="30%">�û�����</TD>
    <TD class=forumRow width="70%"><INPUT class=tableBorder maxLength=10 
      size=25 name=user value=''></TD></TR>
  <TR>
    <TD class=forumRowHighlight align=middle>��&nbsp;&nbsp;�룺</TD>
    <TD class=forumRow><INPUT class=tableBorder type=password maxLength=30 
      size=25 name=pass></TD></TR>
  <TR>
    <TD class=forumRowHighlight align=middle>�����룺</TD>
    <TD class=forumRow><INPUT class=tableBorder maxLength=20 size=4 
      value='<%=session("verifycode")%>' name=verify><input type="hidden" name="verify1" size="4" style="font-size: 12px" value="<%=session("verifycode")%>"> �����븽����:<font color=red><%=session("verifycode")%></font></TD></TR>
  <TR>
    <TD class=forumRowHighlight align=middle colSpan=2><INPUT class=kuang type=submit value="��½" name=Submit>&nbsp;&nbsp;<INPUT class=kuang onclick=window.location.reload() type=button value=ˢ�� name=refresh>&nbsp; 
<INPUT class=kuang onclick="javascript:window.close();" type=button value=�ر� name=Submit1></TD></TR></TBODY></TABLE></FORM>
<P align=center>[��Ȩ���У�<A 
href="http://www.happyjh.com/">���ֽ�����</A> �ͷ�QQ��865240608]</P></DIV></BODY></HTML>
<%elseif pass=Application("aqjh_adminkey") and user=Application("aqjh_adminuser") then
session("aqjh_adminok")=true
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','�ɹ���¼��')"
conn.close
set conn=nothing
if aqjh_name<>Application("aqjh_user") then
'��ʾ��¼��Ϣ
says="<font color=red>��ϵͳ��ʾ��["&aqjh_name&"]</font><font color=green>��¼�˺�̨����ϵͳ���Ա���Ҽල��[��ǰʱ��:"&now()&"]"
call showchat(says)
'��ʾ��¼��Ϣ
end if
Response.Redirect "admin.asp"
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& Request.ServerVariables("REMOTE_ADDR") &"','�����¼','��¼ʧ�ܡ�')"
conn.close
set conn=nothing
if aqjh_name<>Application("aqjh_user") then
'��ʾ��¼��Ϣ
says="<font color=red>��ϵͳ���ء�["&aqjh_name&"]</font><font color=green>��Ƿ���¼��̨���ѱ�ϵͳ���أ�[��ǰʱ��:"&now()&"]"
call showchat(says)
'��ʾ��¼��Ϣ
end if
Response.Redirect "../exit.asp"
end if%>
