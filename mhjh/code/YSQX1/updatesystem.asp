<%
Response.Buffer=true
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"�ٸ�" then Response.Redirect "../exit.asp"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open application("yx8_mhjh_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
rst.Open "x",conn,1,2
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
i=1
do while not rst.EOF
msg=msg&"<tr><td width='30%'>"&rst("a")&"</td><td>"&Request.Form(i)&"</td></tr>"
rst("b")=Request.Form(i)
rst.Update
rst.MoveNext
i=i+1
loop
rst.Requery
do while Not rst.EOF
Select Case rst("a")
			Case "�û���"
				Application("yx8_mhjh_systemname")=rst("b")
			Case "���к�"
				Application("yx8_mhjh_seriesnumber")=rst("b")
			Case "������"
				Application("yx8_mhjh_visitor")=rst("b")
			Case"������ʱ��"
                Application("yx8_mhjh_nosaytime")=rst("b")
			case "�汾��"
				Application("yx8_mhjh_banbenhao")=rst("b")
			case "�ݵ�����"
				Application("yx8_mhjh_paycent")=clng(rst("b"))
			Case "��󹥻�"
				Application("yx8_mhjh_Maxattack")=clng(rst("b"))
		   	Case "��ܿ���"
				Application("yx8_mhjh_chatroomsnkb")=clng(rst("b"))
			Case "վ��"			
			    Application("yx8_mhjh_admin")=rst("b")
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
<body oncontextmenu=self.event.returnValue=false topmargin=0 background="../chatroom/bg1.gif">
<table width=100% border=3><tr><td>ϵͳ����</td><td>������ֵ</td></tr>
<%=msg%>
</table>
<p align=center>ϵͳ���³ɹ�<br>
<input type="button" name="ok" value="��ȷ ����" onclick=javascript:history.go(-1)></p>
</body></html>











