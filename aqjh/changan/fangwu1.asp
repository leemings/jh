<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
id=request("id")
my=aqjh_name
if instr(action,"'") then response.end 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �ȼ�,��ż from �û� where ����='"&aqjh_name&"'",conn
'hy=rs("��Ա")
dj=rs("�ȼ�")
po=rs("��ż")
if dj<20 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('��Դ���ţ���ȼ�����20�����ϲſ��Թ���!');location.href = 'fangwu.asp';}</script>"

response.end
end if
if po="��" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�㻹û�����ţ���������!');location.href = 'fangwu.asp';}</script>"

response.end
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%><!--#include file="data1.asp"--><%
sql="select * from ���� where ����='" & my & "' or ����='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%><!--#include file="data1.asp"--><%
sql="select * from ���� where ID=" & id
rs=conn.execute(sql)
yin=rs("�ۼ�")
huzhu=rs("����")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���� from �û� where ����='"&aqjh_name&"'",conn
if rs("����")<=yin and my=huzhu and huzhu<>"��" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('���򲻳ɹ���ԭ�������������!');location.href = 'fangwu.asp';}</script>"

response.end
end if
     conn.execute "update �û� set ����=����-" & yin & " where ����='"&aqjh_name&"'"
          rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
 %><!--#include file="data1.asp"--><%
      sql="update ���� set ����='" & my & "',����='" & po & "' where ID=" & id
	rs=conn.execute(sql)
	conn.close
	Response.Redirect "fangwu.asp"

else %> 
<script language=vbscript>
            MsgBox "�������İ����Ѿ���������ˡ�"
            location.href = "fangwu.asp"
                    </script>
<%
   rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end if
%>
