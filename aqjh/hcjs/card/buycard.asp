<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="cardset.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
id=lcase(trim(request("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"��")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���� from �û� where ����='"&aqjh_name&"'",conn,3,3
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㲻�ǽ������ˣ�����');}</script>"
	Response.End
end if
rs.close
rs.open "SELECT * FROM b where ID=" & id & " and h>0 and b='��Ƭ'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('û����Ҫ����Ļ�Ա��Ƭ��!');location.href = 'card.asp';}</script>"
	response.end
end if
card=rs("a")
hyyin=rs("h")
sm=rs("c")
rs.close
rs.open "select ��Ա��,w5 from �û� where  ����='"&aqjh_name&"'",conn,3,3
if hyyin>rs("��Ա��") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��Ļ�Ա�𿨲��㣡!');location.href = 'card.asp';}</script>"
	response.end
end if
mycard=add(rs("w5"),card,1)
sl=mywpsl(rs("w5"),card)+1
if sl>max then
   Response.Write "<script language=JavaScript>{alert('��Ŀ�Ƭ["&card&"]�Ѿ�����20�������һ���ٹ���');window.close();}</script>"
else
conn.execute "update �û� set w5='"&mycard&"',��Ա��=��Ա��-" &hyyin & " where  ����='"&aqjh_name&"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('��Ŀ�Ƭ["&card&"]����ɹ�,����"&sl&"����');location.href = 'card.asp';}</script>"
end if%>