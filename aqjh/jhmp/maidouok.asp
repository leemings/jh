<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
dds=request.form("dds")
if not isnumeric(dds) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&dds&"]������󣬶���������ʹ�����֣�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if instr(dds,"'")<>0 or instr(dds," ")<>0 or instr(dds,"or")<>0 or instr(dds,"OR")<>0 or instr(dds,"%20")<>0 or instr(dds,"=")<>0 or instr(dds,">")<>0 or instr(dds,"<")<>0 then Response.Redirect "../error.asp?id=440"
dds=int(abs(dds))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �ݶ����� from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
end if
pdds=rs("�ݶ�����")
rs.close
if pdds<1000 then
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('�����ڻ�û�ж��ӣ�ֻ�У�"&pdds&"�㣬���������ӣ�');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
dousl=int(pdds/1000)
if dds>dousl then
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('������ô�ඹ���𣿣�');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
d=dds*1000
s=pdds-dds*1000
conn.execute "update �û� set �ݶ�����="& s &",��Ա��=��Ա��+"& 2*dds &" where ����='"&aqjh_name&"'"
mess="���������ӣ�"&dds&"������Ա�����ӣ�"&dds*2&"Ԫ����ʣ:"&s&"�㣡"
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="../bg.gif">
<div align="center"><font color="#FF0000"><%=mess%></font> 
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p><input type=button value=�رմ��� onClick='window.close()' name="button"></p>
</div>
</body>
</html>