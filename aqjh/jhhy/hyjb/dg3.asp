<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	response.write "�㲻�ǽ������ˣ����ܽ����������"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
hy=rs("��Ա�ȼ�")
pdhy=rs("��Ա")
if hy<>2 and pdhy=False then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('��ѽ��["&aqjh_name &"]���Ա����2��������������ң�');window.close();</script>"
 response.end
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=10-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');window.close();}</script>"
	Response.End
end if
if rs("����")<1000000 then
		Response.Write "<script language=javascript>alert('��Ǹ�������������100��');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("��")<80 then
		Response.Write "<script language=javascript>alert('��Ǹ����Ľ����Բ���80�㣡');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("�ȼ�")>400 then
		Response.Write "<script language=javascript>alert('��Ǹ����ȼ�̫���ˣ���ȥ��ĵط������ɣ�');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("����")<200000 then
		Response.Write "<script language=javascript>alert('��Ǹ�������������20��~��');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("����")<200000 then
		Response.Write "<script language=javascript>alert('��Ǹ�������������20��');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("ľ")<80 then
		Response.Write "<script language=javascript>alert('��Ǹ�����ľ���Բ���80�㣡~');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("��")<80 then
		Response.Write "<script language=javascript>alert('��Ǹ����������Բ���80�㣡~');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
	end if
if rs("����")<30 then
		Response.Write "<script language=javascript>alert('��Ǹ�������������30��');history.back();</script>"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End	
end if
rs.close
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
conn.execute "update �û� set ����=����-200000,����=����-200000,����=����-1000000,���=���+5,��=��-80,����=����-30,ľ=ľ-80,��=��-80,����ʱ��=now() where ����='"&aqjh_name&"'"
%>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>�����ƽ�ҡ�["&aqjh_name&"]</font><font color=blue>�ڻ�Ա���������˽��<font color=red>5</font>��,��ľ�����Լ���<font color=red>80</font>�㣬������������<font color=red>20</font>��㣬��������<font color=red>30</font>�㣬��������<font color=red>100</font>��!</font>"
call showchat(says)
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
</head>
<body bgproperties="fixed" bgcolor="#000000" vlink="#000000">
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
<table height="260">
<tr>
          <td height="37"> <strong><font color="#800080">���ֽ���--�������:</font></strong> 
        <tr>
          <td height="182" valign="top"><img src="../../chat/img/menoy.gif" width="50" height="40" border="0"><font color="#FF0000">�������������һ�죬���ý��5������������20������20������100������֮��ľ������80����������30�㣡</font><font color="#FF00FF"><Br>
            </font><Br>
          </td>
</tr>
<td align=center height="29">&nbsp;
<div align="right">
<input type=button value="�� ��" onclick="location.href='index.asp'">
</div>
</td></tr>
</table>
</td></tr>
</table>
</body>
</html>
