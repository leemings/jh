<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
sl=abs(int(Request.form("wpsl")))
if sl<1 or sl>50 then
	Response.Redirect "../../error.asp?id=72"
end if
action=lcase(trim(request.querystring("action")))
if InStr(action,"or")<>0 or InStr(action,"=")<>0 or InStr(action,"`")<>0 or InStr(action,"'")<>0 or InStr(action," ")<>0 or InStr(action,"��")<>0 or InStr(action,"'")<>0 or InStr(action,chr(34))<>0 or InStr(action,"\")<>0 or InStr(action,",")<>0 or InStr(action,"<")<>0 or InStr(action,">")<>0 then Response.Redirect "../../error.asp?id=54"
if action="" then 
		Response.Write "<script Language=Javascript>alert('��ʾ����������,ָ����Ʒ������');location.href = 'javascript:history.go(-1)'</script>"
		response.end
end if
name=aqjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ����ʱ�� FROM �û� WHERE ����='"&aqjh_name&"'",conn
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Բ���ϵͳ���ƣ����["&s&"����]�ٲ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if	
rs.close
rs.open "select * from w where c='ҩƷ' and a='" & action & "' and b=" & aqjh_id & " and i>="&sl,conn
	if rs.eof or rs.bof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('��ʾ�������Ʒ�������㣡');location.href = 'javascript:history.go(-1)'</script>"
			response.end
	end if
id=rs("ID")
nei=int(rs("e"))*sl
ti=int(rs("f"))*sl
conn.execute "update w set i=i-"& sl &" where id=" & id
conn.execute "delete * from w where i<=0"
conn.execute "update �û� set ����=����+" & nei & ", ����=����+" & ti & ",����ʱ��=now() where ����='"&aqjh_name&"'"
mess="�������ҩ��"& sl &"������������" & nei & "����������" & ti & "��"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=Application("aqjh_chatroomname")%></title></head>

<body background="back.gif" oncontextmenu=self.event.returnValue=false>
<table border="0" align="center" width="300" cellspacing="0" cellpadding="0">
<tr align="center">
<td width="300" height="30"><font
color="FF6600"><b>������ʾ</b></font></td>
</tr>
<tr align="center">
<td>
<table width="260">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
</td>
</tr>
</table>
</td>
</tr>
<tr align="center">
<td width="500" height="30"><a href="javascript:location.replace('eat.asp');" onmouseover="window.status='����';return true;" onmouseout="window.status='';return true;">����װ��һ��</a></td>
</tr>
</table>

</body>