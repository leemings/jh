<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../const.asp"-->
<!--#include file="../../chk.asp"-->
<!--#include file="../../mywp.asp"-->
<%
response.write "<body bgcolor=#000000>"
call chkpost()
wpid=Request("id")
wpsl=LCase(Trim(Request.Form("sl")))
if not isnumeric(wpsl) then
	Response.Write "<script language=javascript>{alert('��Ʒ������ʹ�ð��������д��');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
if not isnumeric(wpid) then
	Response.Write "<script language=javascript>{alert('�Ƿ����ݣ�');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
wpsl=abs(int(clng(wpsl)))
if wpsl=0 then
	Response.Write "<script language=javascript>{alert('����дҪȡ����Ʒ������');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from w where id="&wpid,conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>{alert('��Ĵ�������û������������');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
yhid=rs("a")
wpname=rs("b")
wpzs=rs("c")
filedname=rs("d")
if wpsl>wpzs then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>{alert('�������е�["&wpname&"]ֻ��["&wpzs&"]��������"&wpsl&"��');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
rs.close
rs.open "select id,"&filedname&" from �û� where ����='"&aqjh_name&"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>{alert('�㲻�ǽ������ˣ�');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
myid=rs("id")
if yhid<>myid then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>{alert('�ⲻ����Ķ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
mywp=trim(rs(filedname))
if mywp<>"" and not isnull(mywp) then
	data1=split(mywp,";")
	data2=UBound(data1)
	if data2>=25 then
		erase data1
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>{alert('�����������Ĵ�����Ʒ�ѳ���25�֣���Ҳ�Ų����ˡ�');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End
	end if
	erase data1
end if
if wpsl=wpzs then
	conn.execute "delete * from w where id="&wpid&" and a="&myid
else
	conn.execute "update w set c=c-"&wpsl&" where id="&wpid&" and a="&myid
end if
temper=add(mywp,wpname,wpsl)
conn.execute "update �û� set "&filedname&"='"&temper&"' where ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<Script Language=JavaScript>{parent.wupin.document.location.reload();}</Script>"
Response.Redirect "bxg.asp"
%>