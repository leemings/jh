<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="jmconfig.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"

yin=jmyin	'����
jb=jmjb		'���
id=request("id")
if id="" then
	response.write "��������"
	response.end
end if
if not isnumeric(id) then
	response.write "<script language=JavaScript>{alert('����ID��ʹ������');window.close();}</script>"
	response.end
end if
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("e_zgzm.asp")
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
	conn.open application("sjjh_usermdb")
	rs.open "select ����,��� from �û� where ����='"&sjjh_name&"'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "�㲻�ǽ������ˡ�"
		response.end
	end if
	newyin=rs("����")-yin
	newjb=rs("���")-jb
	rs.close
	if newyin<0 or newjb<0 then
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<script language=javascript>{alert('�ܹ����˽���Ҳ������ѵģ�һ����ȡ������"&yin&",��ң�"&jb&"����û��ô��Ǯ�����������εġ�');window.close();}</script>"
		response.end
	end if
	conn.execute "update �û� set ����="&newyin&",���="&newjb&" where ����='"&sjjh_name&"'"
	conn.close
end if
set rs=nothing
set conn=nothing
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
        <P align=center>��E�߽�����&#8482; 2003-2004<br><br><br>
        </P>
      </TD><TD align=middle width=8 bgColor=#808080></TD></TR>
<TR style="FONT-SIZE: 3pt" align=middle><TD width=16 height=8></TD><TD bgColor=#808080 height=8></TD>
<TD width=8 bgColor=#808080 height=8></TD></TR></TBODY></TABLE></CENTER></body></html>
