<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����ʱ�� from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=454"
	response.end
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<30 then
ss=30-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�ոմ������������Ϣ�У���ȣ�"&ss&"�ֺ�������');location.href = 'javascript:history.back()';}</script>"
	Response.End 
end if
if rs("����")<200 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�����������200�������������﹤����');location.href = 'javascript:history.back()';}</script>"
	Response.End 
end if
conn.execute "update �û� set ����=����-200,����=����+8000,����ʱ��=now() where ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>��˽�˱��ڡ�</font><font color=green>["&aqjh_name&"]�����������һ�챣�ڣ���������8000�������½�200</font>"
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
<title>���齭��</title></head>
<body bgproperties="fixed" bgcolor="#000000" vlink="#000000">
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
<table height="260">
<tr>
          <td height="37"> <font color="#000000"><strong>���齭����:</strong></font> 
        <tr>
<td height="182" valign="top">�����������һ�챣�ڣ���������8000�������½�200��<Br><Br>
          </td>
</tr>
<td align=center height="29">&nbsp;
<div align="right">
<input type=button value="�� ��" onclick="location.href='dg.asp'">
</div>
</td></tr>
</table>
</td></tr>
</table></body></html>