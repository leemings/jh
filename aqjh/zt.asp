<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<%
sub gfsj(gfdata)
if not(isnull(gfdata)) and instr(gfdata,"|")<>0 then
	mydata=split(gfdata,"|")
	mygj=mygj+mydata(2)
	myfy=myfy+mydata(3)
	erase mydata
end if
end sub
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<!��IP>
<%
ip=request("ip")
if ip="" then
ip=Request.ServerVariables("REMOTE_ADDR")
end if
sip=split(ip,".")
if ubound(sip)<>3 then
ip=Request.ServerVariables("REMOTE_ADDR")
erase sip
sip=split(ip,".")
end if
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
rs.open "select * FROM z WHERE a<=" & num & " and b>=" & num,conn,1,1
if rs.eof or rs.bof then
country="����"
city="δ֪"
else
country=rs("c")
city=rs("d")
end if
if country="" then country="�й�"
if city="" then city="δ֪"
last="����:" & country & " ����:" & city
rs.close
%>
<!��IP>
<%
rs.open "Select * from �û� where ����='" & aqjh_name &"'",conn,3,3
gjsx=rs("�ȼ�")*aqjh_gjsx
fysx=rs("�ȼ�")*aqjh_fysx
if rs("��Ա")=true and clng(DateDiff("d",now(),rs("��Ա����")))>0 then
	pdstr="<font size=-1>[�ݵ��ƻ�Ա]"&rs("��Ա����")&"</font>"
else
	pdstr=""
end if
wgj=rs("�书��")
nlj=rs("������")
tlj=rs("������")
mygj=rs("����")
myfy=rs("����")
mywg=rs("�书")
call gfsj(rs("z1"))
call gfsj(rs("z2"))
call gfsj(rs("z3"))
call gfsj(rs("z4"))
call gfsj(rs("z5"))
call gfsj(rs("z6"))
if rs("����1")<>"����" and rs("����1")<>"" then
	zbdata=split(rs("����1"),"|")
	xs=0
	if UBound(zbdata)=5 then
		sj=datediff("n",zbdata(1),now())
		jgj=zbdata(2)
		jfy=zbdata(3)
		jwg=zbdata(4)
		if sj<=60 then
			mygj=mygj+zbdata(2)
			myfy=myfy+zbdata(3)
			mywg=mywg+zbdata(4)
			xs=1
		end if
	end if
	erase zbdata
end if
%>
<HTML><HEAD><TITLE>����״̬</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="ztimg/css.css" rel=stylesheet type=text/css>
</HEAD>
<BODY background=ztimg/back1.gif bgColor=#000000 leftMargin=0 text=#cccccc topMargin=0 marginheight="0" marginwidth="0" oncontextmenu=window.event.returnValue=false>
<DIV align=center>
<TABLE align=center bgColor=#000000 border=0 cellPadding=6 cellSpacing=1 width=600>
<TBODY>
<TR bgColor=#4f3e39>
<TD align=middle class=title colSpan=7>
<TABLE border=1 borderColor=#000000 cellPadding=0 cellSpacing=0 width="100%">
<TBODY>
<TR>
<TD bgColor=#003333 borderColor=#cccccc>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width="95%">
<TBODY>
<TR>
<TD><FONT color=#ffcc00><%=aqjh_name%></FONT>���� �����������ĸ���״̬<IMG align=absMiddle height=16 src="ztimg/icon16.gif" width=16>[<A href="../hcjs/jhzb/index.asp">����װ��</A>] [<a href="../CHANGAN/xiaowu.asp">����</a>]
<!��IP>
<form name="form2" method="get" action="">
��IP:<input name="ip" type="text" id="ip" size="15" maxlength="15" style="color: #666666; background-color: #ffffff; text-decoration: blink; border: 1px solid #009900" value="<%=ip%>">
<input type="submit" name="Submit" value="����" style="BACKGROUND-COLOR: #009933; BORDER-BOTTOM: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; CURSOR: hand; FONT-FAMILY: '����'; COLOR: #FFFFFF; FONT-SIZE: 9pt; HEIGHT: 18px">
���ԣ�<%=last%></form>
<!��IP>
</TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR>
<TR>
<TD align=middle bgColor=#644e49 rowSpan=3 width="45">
<IMG border=0 height=30 width=30 name=w_r6_c3 src="<%=rs("����ͷ��")%>">
</TD>
<TD bgColor=#644e49 colSpan=2 align=center>
<FONT color=#ffcc00><%=aqjh_name%></FONT>
</TD>
<TD align=middle bgColor=#51403c width=44>�ɣ�</TD>
<TD align=middle bgColor=#644e49 width="117"><FONT color=#ff0000><%=rs("id")%></FONT></TD>
<TD align=middle bgColor=#51403c width="30">�Ա�</TD>
<TD align=middle bgColor=#644e49 width=141><%=rs("�Ա�")%></TD>
</TR>

<TD align=middle bgColor=#51403c width="30">
<IMG height=1 src="ztimg/spacer.gif" width=30><BR>����</TD>
<TD align=middle bgColor=#644e49 width=117><%=rs("����")%></TD>
<TD align=middle bgColor=#51403c width="44">
<IMG height=1 src="ztimg/spacer.gif" width=30><BR>ְλ</TD>
<TD align=middle bgColor=#644e49 width=117><%=rs("ְλ")%></TD>
<TD align=middle bgColor=#51403c width="30">
<IMG height=1 src="ztimg/spacer.gif" width=30><BR>����</TD>
<TD align=middle bgColor=#644e49 width=141><%=rs("����")%></TD>
</TR>


<TR>
<TD align=middle bgColor=#51403c width="30">����</TD>
<TD align=middle bgColor=#644e49 width=117><FONT color=#ffcc00><IMG align=absMiddle height=16 src="ztimg/space.gif" width=18></FONT>
<FONT color=#cccccc><%=rs("����")%></FONT></TD>
<TD align=middle bgColor=#51403c width="44">���</TD>
<TD align=middle bgColor=#644e49 width=117><%=rs("���")%></TD>
<TD align=middle bgColor=#51403c width="30"><font color="#cccccc">����</font></TD>
<TD align=middle bgColor=#644e49 width=141><FONT color=#cccccc><%=aqjh_grade%></FONT></TD>
</TR>
<TR>


<TR bgColor=#433730>
<TD align=middle bgColor=#644e49 rowspan="9" valign="middle"><img src="ztimg/grzt.gif" width="45" height="192"></TD>
<TD align=middle bgColor=#51403c width="30">��Ա</TD>
<TD align=middle bgColor=#644e49 width="117"><%=rs("��Ա�ȼ�")%></TD>
<TD align=middle bgColor=#51403c width="44">����</TD>
<TD align=middle bgColor=#644e49 width="117"><%=rs("��Ա����")%></TD>
<TD align=middle bgColor=#51403c width="30">����</TD>
<TD align=middle bgColor=#644e49 width=141><%=rs("����")%></TD>
</TR>
<TR bgColor=#433730>
<TD align=middle bgColor=#51403c width="30">ְҵ</TD>
<TD align=middle bgColor=#644e49 width=117><%=rs("ְҵ")%></TD>
<TD align=middle bgColor=#51403c width="44">�ȼ� </TD>
<TD align=middle bgColor=#644e49 width=117><B><%=rs("�ȼ�")%></B></TD>
<TD align=middle bgColor=#51403c width="30">����</TD>
<%if rs("����")=true then
	bhlg="��������"
else
	bhlg="û�б���"
end if%>
<TD align=middle bgColor=#644e49 width=141><%=bhlg%></TD>
</TR>
<TR bgColor=#433730>
<TD align=middle bgColor=#51403c width="30">����</TD>
<TD align=middle bgColor=#644e49 width=117>
<%
if rs("�ȼ�")<20  then response.write("ƽ��")
if rs("�ȼ�")>=20 and rs("�ȼ�")<30 then response.write("ʿ��")
if rs("�ȼ�")>=30 and rs("�ȼ�")<50 then response.write("��ʿ")
if rs("�ȼ�")>=50 and rs("�ȼ�")<70 then response.write("��ʿ")
if rs("�ȼ�")>=70 and rs("�ȼ�")<95 then response.write("��ʿ")
if rs("�ȼ�")>=95 and rs("�ȼ�")<120 then response.write("ʥ��ʿ")
if rs("�ȼ�")>=120 and rs("�ȼ�")<155 then response.write("ʥ��ʿ")
if rs("�ȼ�")>=155 and rs("�ȼ�")<190 then response.write("ʥ��ʿ")
if rs("�ȼ�")>=190 and rs("�ȼ�")<235 then response.write("����սʿ")
if rs("�ȼ�")>=235 and rs("�ȼ�")<290 then response.write("����ս��")
if rs("�ȼ�")>=290 and rs("�ȼ�")<345 then response.write("�޼�ս��")
if rs("�ȼ�")>=345 and rs("�ȼ�")<410 then response.write("��������")
if rs("�ȼ�")>=410 and rs("�ȼ�")<475 then response.write("����֮��")
if rs("�ȼ�")>=475 and rs("�ȼ�")<550 then response.write("����ս��")
if rs("�ȼ�")>=550 then response.write("��������")%>
</TD>
<TD align=middle bgColor=#51403c width="44">����</TD>
<TD bgColor=#644e49 colspan="3">��飺<%=rs("������")%>�� <%if rs("��ż")<>"��" and rs("��ż")<>"" then%>����:<%=rs("��������")%>
        ���ʱ��:<%=DateDiff("d",rs("��������"),date())%>��<%end if%></TD>
</TR>
<TR bgColor=#433730>
<TD align=middle bgColor=#51403c width="30">���</TD>
<TD align=middle bgColor=#644e49 width=117><%=rs("���")%></TD>
<TD align=middle bgColor=#51403c width="44">��</TD>
<TD align=middle bgColor=#644e49 width=117><%=rs("��Ա��")%></TD>
<TD align=middle bgColor=#51403c width="30">����</TD>
<TD align=middle bgColor=#644e49 width=141><%=rs("����")%></TD>
</TR>
<TR bgColor=#433730>
<TD align=middle bgColor=#51403c width="30">����</TD>
<TD align=middle bgColor=#644e49 width=117>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width="120" height="13">
<TBODY>
<TR>
<%if rs("����")<=0 then
	abc=0
elseif rs("����")>=(rs("�ȼ�")*aqjh_tlsx+5260+tlj) then
	abc=120
else
	abc=int(120*(rs("����")/(rs("�ȼ�")*aqjh_tlsx+5260+tlj)))
end if%>
<TD width=120 background="ztimg/000.gif"><img src="ztimg/001.gif" width="<%=abc%>" height="11"></TD>
</TR>
<TR>
<TD width=121 align="center">
<font color=yellow><%=rs("����")%></font>
<%if aqjh_sx=1 then%>
/<font color=green><%=rs("�ȼ�")*aqjh_tlsx+5260+tlj%></font>
<%end if%>
</TD>
</TR>
</TBODY>
</TABLE>
</TD>
<TD align=middle bgColor=#51403c width="44">����</TD>
<TD align=middle bgColor=#644e49 width=117>
<TABLE border=0 cellPadding=0 cellSpacing=0 width="120" height="13">
<TBODY>
<%if rs("����")<=0 then
	abc=0
elseif rs("����")>=(rs("�ȼ�")*aqjh_nlsx+2000+nlj) then
	abc=120
else
	abc=int(120*(rs("����")/(rs("�ȼ�")*aqjh_nlsx+2000+nlj)))
end if%>
<TR>
<TD width=120 background="ztimg/000.gif"><img src="ztimg/001.gif" width="<%=abc%>" height="11"></TD>
</TR>
<TR>
<TD width=117 align="center">
<font color=yellow><%=rs("����")%></font>
<%if aqjh_sx=1 then%>
/<font color=green><%=rs("�ȼ�")*aqjh_nlsx+2000+nlj%></font>
<%end if%>
</TD>
</TR>
</TBODY>
</TABLE>
</TD>
<TD align=middle bgColor=#51403c width="30">�书</TD>
<TD align=middle bgColor=#644e49 width=141>
<TABLE border=0 cellPadding=0 cellSpacing=0 width="120" height="13">
<TBODY>
<%if mywg<=0 then
	abc=0
elseif mywg>=(rs("�ȼ�")*aqjh_wgsx+3800+wgj) then
	abc=120
else
	abc=int(120*(mywg/(rs("�ȼ�")*aqjh_wgsx+3800+wgj)))
end if%>
<TR>
<TD width=122 background="ztimg/000.gif"><img src="ztimg/001.gif" width="<%=abc%>" height="11"></TD>
</TR>
<TR>
<TD width=122 align="center">
<font color=yellow><%=mywg%></font>
<%if aqjh_sx=1 then%>
/<font color=green><%=rs("�ȼ�")*aqjh_wgsx+3800+wgj%></font>
<%end if%>
</TD>
</TR>
</TBODY>
</TABLE>
</TD>
</TR>
<TR bgColor=#51403c>
<TD align=middle width="45">����</TD>
<TD align=middle bgColor=#644e49><%=mygj%></TD>
<td align=middle>����</td>
<td align=middle bgColor=#644e49><%=myfy%></td>
<td>���</td>
<td align=middle bgColor=#644e49><%=rs("���")%></td>
</TR>
<TR bgColor=#51403c>
<TD align=middle width="45">���</TD>
<TD align=middle bgColor=#644e49><%=rs("allvalue")%></TD>
<td align=middle>����</td>
<td align=middle bgColor=#644e49>
<%if rs("grade")=3 and rs("���")="����" then
	response.write (rs("�ݶ�����")-int(rs("�ݶ�����")/500)*500)&"<font color=#FF0000>/500</font>"
else
	response.write (rs("�ݶ�����")-int(rs("�ݶ�����")/1000)*1000)&"<font color=#FF0000>/1000</font>"
end if%>
</td>
<td>����</td>
<td align=middle bgColor=#644e49>
<%if rs("grade")=3 and rs("���")="����" then
	response.write int(rs("�ݶ�����")/500)
else
	response.write int(rs("�ݶ�����")/1000)
end if%>
</td>
</TR>
<TR bgColor=#51403c>
<TD align=middle width="45">ɱ��</TD>
<TD align=middle bgColor=#644e49><%=rs("ɱ����")%></TD>
<td align=middle>����</td>
<td align=middle bgColor=#644e49><%=rs("��ɱ��")%></td>
<td>����</td>
<td align=middle bgColor=#644e49><%if rs("����")<>"��" and rs("����")<>"" then%><font color=red><%=rs("����")%></font><%else%><%=rs("����")%><%end if%></td>
</TR>
<TR bgColor=#51403c>
<TD align=middle width="45">����</TD>
<TD align=middle bgColor=#644e49><%=rs("����")%></TD>
<td align=middle>����</td>
<td align=middle bgColor=#644e49><%=rs("����")%></td>
<td>ʦ��</td>
<td align=middle bgColor=#644e49><%=rs("ʦ��")%></td>
</TR>

<%sj=DateDiff("n",rs("����ʱ��"),now())
if sj<=60 then%>
<TR bgColor=#433730>
<TD bgColor=#51403c colspan="7" height="28" align="center">
<font color="#00FFFF">��������ʱ�仹ʣ��<%=60-sj%>��</font>
</TD>
</TR>
<%end if%>
<TR bgColor=#000000>
<TD align=middle colspan="7" height="28"><FONT color=#ffffff>&copy; ��Ȩ���� 2004-2005 </FONT><A href="http://www.7758530.com/" target=_blank><FONT color=#ffffff>���齭����</FONT></A></TD>
</TR></TBODY></TABLE></DIV></BODY></HTML>