<%
Sub Msg (v)
 Response.Write "<script Language=JavaScript>alert('" & v & "');history.back();</script>"
 Response.End
End Sub
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_grade>5 then
%>
<script language=javascript>
     history.back()
     alert("��ʾ����Ϊ�ٸ����ܲ�����")
</script>
<%
response.end
end if
hytitle=request("hytitle")
hyname=request("hyname")
if hyname="" or hytitle="" then
   Msg "����Ϊ��"
end if
%>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
sql="select * from ��԰���ư� where ������='"&aqjh_name&"'"
rs.open sql,connt,1,1
if not rs.EOF then 
%>
<script language=javascript>
     history.back()
     alert("���棺���Ѿ��������˰���")
</script>
<%
response.end
else
    rs.close
    rs.open "select * from �û� where ����='"&aqjh_name&"'",conn,1,1
    if rs("����")<100000 then
%>
<script language=javascript>
     history.back()
     alert("��ʾ�����ڹٳ�̫�ڣ�û��10�����ǲ��еģ�")
</script>
<%
         response.end
    end if
    conn.Execute "update �û� set ����=����-100000 where ����='"&aqjh_name&"'"
    sql="INSERT INTO ��԰���ư� (������,��������,����˵��,��������) VALUES ('"&aqjh_name&"','"&hyname&"','"&hytitle&"','"&now()&"')"
    connt.Execute (sql)
%>
<script language=javascript>
     window.location.href ="hfyb.asp"
     alert("��ʾ����ϲ�㱨���ɹ����͵����н��ɣ�")
</script>
<%
end if
%>