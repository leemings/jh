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
boyid=cint(request("id"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if request("w")="del" then
if aqjh_grade<>10 then
 response.write "<script>alert('����Ȩɾ������������~��');history.back(-1);</script>"
 response.end
end if
sql="delete from gry where id="&boyid
conn.execute(sql)
response.redirect "lingyang.asp"
end if
rs.open "select * from �û� where ����='" &aqjh_name & "'",conn
if rs("����")<10000000 then
	Response.Write "<script language=JavaScript>{alert('��ʾ������û��1ǧ�򣬹¶�������һ���߰���');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("boy")<>"" then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㻹�к��Ӱ�����ʲô��');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
peiou=rs("��ż")
if peiou="" or peiou="��" then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㻹û����أ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
If DateDiff("d",rs("��������"),date())<30 Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����黹����һ���£�');location.href = 'javascript:history.go(-1)';</script>"	
	response.end
end if
rs.close
rs.open "select * from gry where id=" &boyid,conn
if rs.eof then
	Response.Write "<script language=JavaScript>{alert('��ʾ���˹¶������ڣ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
else
 boy=rs("boy")
end if
rs.close
zt=split(boy,"|")
if zt(1)="��" then boysex="images/boy.gif" else boysex="images/girl.gif"
conn.execute "update �û� set boy='"&boy&"',boysex='"&boysex&"',����=����-10000000,����=����+50 where ����='" & aqjh_name & "'"
conn.execute "update �û� set boy='"&boy&"',boysex='"&boysex&"' where ����='" & peiou & "'"
conn.execute "delete * from gry where id="&boyid
call showchat("<font color=red>������С����</font>"&aqjh_name&"����1ǧ���������ӹ¶���������һ��С������������50�㣡")
Response.Write "<script language=JavaScript>{alert('��ϲ������С���ɹ�����ȥι���ɣ�');location.href = 'javascript:history.go(-1)';}</script>"
%>