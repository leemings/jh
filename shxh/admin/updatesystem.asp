<%
Response.Buffer=true
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
if not(session("Ba_jxqy_usercorp")="�ٸ�" and Session("Ba_jxqy_usergrade")=Application("Ba_jxqy_allright")) then Response.Redirect "../error.asp?id=500"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open application("Ba_jxqy_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
rst.Open "ϵͳ����",conn,1,2
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
i=1
do while not rst.EOF 
	msg=msg&"<tr><td width='30%'>"&rst("����")&"</td><td>"&Request.Form(i)&"</td></tr>"
	rst("����ֵ")=Request.Form(i)
	rst.Update
	rst.MoveNext
	i=i+1
loop
rst.Requery
do while Not rst.EOF
		Select Case rst("����")
			Case "ϵͳ����"
				Application("Ba_jxqy_systemname")=rst("����ֵ")
			Case "�û���"
				Application("Ba_jxqy_userright")=rst("����ֵ")
			Case "��Ȩ����"
				Application("Ba_jxqy_copyrightassert")=rst("����ֵ")
			Case "���к�"
				Application("Ba_jxqy_seriesnumber")=rst("����ֵ")
			Case "����ɫ"
				Application("Ba_jxqy_backgroundcolor")=rst("����ֵ")	
			Case "����ͼƬ"
				Application("Ba_jxqy_backgroundimage")=rst("����ֵ")	
			Case "���к�"
				Application("Ba_jxqy_backgroundcolor")=rst("����ֵ")	
			Case "��������"
				Application("Ba_jxqy_visitor")=rst("����ֵ")
			Case "�������ʱ��"
				rst("����ֵ")=cstr(now())
				rst.Update
				Application("Ba_jxqy_opendata")=cstr(now())
			Case "��������"
				Ba_grade=split(rst("����ֵ"),";")
				for i=0 to ubound(Ba_grade)
					application("Ba_jxqy_grade"&i)=clng(Ba_grade(i))
				next
			Case "������ʾ"
				Application("Ba_jxqy_guestjoin")=rst("����ֵ")
			Case "�˳���ʾ"
				Application("Ba_jxqy_guestleft")=rst("����ֵ")
			Case "������ʾ"
				Application("Ba_jxqy_guestisnotconnection")=rst("����ֵ")
			Case "����Ȩ��"
				Application("Ba_jxqy_kickguestright")=clng(rst("����ֵ"))	
			Case "����Ȩ��"
				Application("Ba_jxqy_arrestright")=clng(rst("����ֵ"))
			Case "����Ȩ��"
				Application("Ba_jxqy_gaolright")=clng(rst("����ֵ"))
			Case "����Ȩ��"
				application("Ba_jxqy_lockipright")=clng(rst("����ֵ"))		
			Case "ը��Ȩ��"
				Application("Ba_jxqy_bombright")=clng(rst("����ֵ"))
			Case "����Ȩ��"
				Application("Ba_jxqy_unlockipright")=clng(rst("����ֵ"))
			Case "����Ȩ��"
				Application("Ba_jxqy_exaltgraderight")=clng(rst("����ֵ"))
			Case "����Ȩ��"
				Application("Ba_jxqy_declinegraderight")=clng(rst("����ֵ"))
			Case "վ��Ȩ��"
				Application("Ba_jxqy_allright")=clng(rst("����ֵ"))				
			Case "��������"
				Application("Ba_jxqy_illegidimacyname")=rst("����ֵ")
			Case"��������ʱ��"
				Application("Ba_jxqy_nosaytime")=clng(rst("����ֵ"))
			Case "ϵͳ��Կ"
				Application("Ba_jxqy_passwordkey")=rst("����ֵ")
			case "����"	
				Application("Ba_jxqy_advertisemen")=rst("����ֵ")
			case "����߶�"
				Application("Ba_jxqy_advertisemenheight")=clng(rst("����ֵ"))
			case "�ݵ�����"
				Application("Ba_jxqy_paycent")=clng(rst("����ֵ"))
			case "��������"
				Application("Ba_jxqy_maxonline")=clng(rst("����ֵ"))
			Case "���˻���"
				Application("Ba_jxqy_newuser")=clng(rst("����ֵ"))
			case "��Ա���ܿ���"
				Application("Ba_jxqy_fellow")=cbool(rst("����ֵ"))
			Case "��󹥻�"
				Application("Ba_jxqy_Maxattack")=clng(rst("����ֵ"))
			Case "�ĳ����ʱ��(��)"
				Application("Ba_jxqy_bettime")=clng(rst("����ֵ"))			
		end select
		rst.MoveNext
	loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel="stylesheet" href="../chatroom/css.css">
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=0 background="<%=bgimage%>" bgcolor="<%=bgcolor%>">
<table width=100% border=3><tr><td>ϵͳ����</td><td>������ֵ</td></tr>
<%=msg%>
</table>
<p align=center>ϵͳ���³ɹ�<br>
<input type="button" name="ok" value="��ȷ ����" onclick=javascript:history.go(-1)></p>
</body></html>