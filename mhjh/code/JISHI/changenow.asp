<%
Response.Expires=-1
bgcolor=Application("yx8_mhjh_backgroundcolor")
cmid=Request.QueryString("id")
cmprice=Request.Form("price")
cmprice=clng(cmprice)
if cmprice<1 then cmprice=1
if not (isnumeric(cmid) and  isnumeric(cmprice))then Response.Redirect "../error.asp?id=024"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="select * from ��Ʒ where id="&cmid&" and ������='"&username&"' and ����=True "
rst.open sqlstr,conn
if  rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=030"
jg=rst("�۸�")
rst.Close
set rst=nothing
if cmprice>jg then Response.Redirect "../error.asp?id=032"
sqlstr="update ��Ʒ set �۸�="&cmprice&" where id="&cmid&" and ������='"&username&"'"
conn.Execute sqlstr
conn.Close
set conn=nothing
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
		newtalkarr(599)="<font color=red>������ͨ����</font><font color=blue>"&username&"�����Լ����۵�һ��������Ʒ�۸��ʺϣ���΢������һ�£����"&cmprice&"��Ӧ�ò���˰ɣ�����Ҫ�Ŀ�ȥ�����г�����</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr

%>
<head><title>���ɼ���</title><LINK href="../style4.css" rel=stylesheet><script language=javascript>setTimeout("location.replace('chgprice.asp');",3000);</script></head>
<body bgcolor='<%=bgcolor%>'  oncontextmenu=self.event.returnValue=false>
<table width=100% border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF bgcolor=FFB442>
<tr align=center>
<td><A href="market.asp" onmouseover="window.strtus='����������ͨ����Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
<td><A href="sale.asp" onmouseover="window.strtus='��۳����Լ�ӵ�е���Ʒ';return true;" onmouseout="window.status='';return true">����</a></td>
<td bgcolor=#FFCC00><A href="chgprice.asp" onmouseover="window.strtus='����Ѽ�����Ʒ�ļ۸�';return true;" onmouseout="window.status='';return true">�ļ�</a></td>
</tr></table>
<p align=center><b><font color="#CC0000" size="4" face="��Բ"><%=Application("yx8_mhjh_systemname")%></font><font color="#FFFF00" size="4" face="��Բ">���ɼ���</font></b></p>
<div align="center">���ǻᰴ�µı�۳��۴���Ʒ������������ǽ����Ͻ�Ǯ���㣬ˡ������֪ͨ�ˣ��������10% </div>
<p align=center><a href="javascript:location.replace('chgprice.asp');" onmouseover="window.status='����';return true;" onmouseout="window.status='';return true;">����</a></p>
</body>
