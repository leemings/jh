<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="jmconfig.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=session("yx8_mhjh_username")
sjjh_grade=session("yx8_mhjh_usergrade")
sjjh_jhdj=session("yx8_mhjh_usergrade")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"

yin=jmyin	'����
id=request("id")
if id="" then
	response.write "��������"
	response.end
end if
if not isnumeric(id) then
	response.write "<script language=JavaScript>{alert('����ID��ʹ������');window.close();}</script>"
	response.end
end if
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("zgzm.asp")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
rs.open "select * from zgjm where id="&id,conn
if rs.eof or rs.bof then
	jmsay="û���ҵ���ָ���������"
	lb="��"
else
	jmbt=rs("b")
	jmsay=rs("c")
	jmlb=rs("a")
	rs.close
	rs.open "select * from lb where id="&jmlb,conn
	if not (rs.eof or rs.bof) then
		lb=rs("lb")
	else
		lb="δ����"
	end if
end if
rs.close
conn.close
if lb<>"��" then
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
	rs.open "select ���� from �û� where ����='"&sjjh_name&"'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "�㲻�ǽ������ˡ�"
		response.end
	end if
	newyin=rs("����")-yin
	rs.close
	if newyin<0 then
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<script language=javascript>{alert('�ܹ����˽���Ҳ������ѵģ�һ����ȡ������"&yin&",��û��ô��Ǯ�����������εġ�');window.close();}</script>"
		response.end
	end if
	conn.execute "update �û� set ����="&newyin&" where ����='"&sjjh_name&"'"
	kl="<font color=0000FF>"&sjjh_name&"</font>���˸����Σ�<font color=0000FF>��"&jmbt&"��</font>,�����������ҵ���˵�е��ܹ�������"&yin&"������,���ڵõ��˴𰸣�<font color=red>"&jmsay&"��</font>"
	conn.close
end if
set rs=nothing
set conn=nothing
says="<font color=red><b>���ܹ����Ρ�</b></font>"&kl	
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
		newtalkarr(598)="008000" 
		newtalkarr(599)=""&says&""
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
%>
<html><head><title>�ܹ�����</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet></head>
<body background="bg.gif"><CENTER><TABLE cellSpacing=0 cellPadding=0 width=80% border=0><TBODY> 
<TR style="FONT-SIZE: 9pt"><TD colspan="2" bgColor=#D7EBFF></TD><TD width=8 height=16></TD></TR>
<TR><TD width=16 bgColor=#D7EBFF></TD><TD style="FONT-SIZE: 9pt; LINE-HEIGHT: 160%" bgColor=#D7EBFF><BLOCKQUOTE>
<P align=center class="3D"><B>�ܹ�����</B></P>
<BR>
&nbsp;<b>���</b><font color=red><%=lb%></font> <b>���ε��ˣ�</b><font color=red><%=jmbt%></font> 
<HR color=red SIZE=1><P style="word-break:break-all"><%=jmsay%></p>
<HR color=red SIZE=1><P align=center><input type=button value="�رմ���" onclick="window.close()" class="input">
<br><br>
            <font color="#0000FF"><%=application("sjjh_chatroomname")%></font>֮<font color="#FF0000">�ܹ�����</font></P>
        </BLOCKQUOTE>
        <P align=center>��Ϸ������<br><br><br>
        </P>
      </TD><TD align=middle width=8 bgColor=#808080></TD></TR>
<TR style="FONT-SIZE: 3pt" align=middle><TD width=16 height=8></TD><TD bgColor=#808080 height=8></TD>
<TD width=8 bgColor=#808080 height=8></TD></TR></TBODY></TABLE></CENTER></body></html>
