<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.charset="gb2312"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
bkmn=Request.Form("money")
bkop=Request.Form("op")
id=Request.Form("money")
if id="" then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������д���');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if


if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"��")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then 
	Response.Write "<script language=JavaScript>{alert('��ʾ��ϵͳ��ֹ�ַ�����ȷ�����������������ȷ�ģ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����,ת�� FROM �û� WHERE ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ֹ����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
if round(rs("����"))<3000 then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����²���3000');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
if rs("ת��")<2 then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��ת������2�Σ�û�����Ȩ��');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
        rs.close

rs.open "select ���,���� FROM �û� WHERE ����='"&bkmn&"'",conn
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��"&bkmn&"��Ϣ������ȷ���Ƿ��д���');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if
if rs("���")=false  then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��"&bkmn&"û�д����״̬����ֹ���飡');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End	
end if

        rs.close

	set rs=nothing	
	conn.close
	set conn=nothing



select case bkop
case "�߳�"
	Response.Redirect "jlzudui.asp?id=ɾ����Ա&to1="&bkmn&""
case "����"
	Response.Redirect "jlzudui.asp?id=�����Ա&to1="&bkmn&""
end select
%>
</body>
</html>