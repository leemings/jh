<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
id=request("id")
my=sjjh_name
sex=sjjh_jhdj
if instr(action,"'") then response.end 
%>
<!--#include file="data1.asp"--><%
sql="select * from ���� where ID=" & id
rs=conn.execute(sql)
yin=rs("�ۼ�")
huzhu=rs("����")
blv=rs("����")
zt=rs("״̬")
if rs("����")<0 or rs("��������")<0 or rs("����")<0  then
%><script language=vbscript>
                         MsgBox "���Ѿ��ɹ�����."
                         location.href = "fangwu.asp"
                    </script><%
conn.execute "update ���� set ����='��',����='��',����=0,����=0,��������=0 where ID=" & id
conn.close
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="select * from �û� where ����='"&sjjh_name&"'"
set rs=conn.execute(sql)
if my=huzhu  and zt="����" then
      conn.execute "update �û� set ����=����+" & yin & "*1/5 where ����='"&sjjh_name&"'"
            %><!--#include file="data1.asp"--><%
      conn.execute "update ���� set ����='��',����='��',����=0,����=0,��������=0 where ID=" & id
set rs=nothing
conn.close
set conn=nothing
	Response.Redirect "fangwu.asp"
else
		%>
 <script language=vbscript>
 MsgBox "���ײ��ɹ���ԭ���㲻��������ӵ����˻�����ķ��ӳ��˵�״��!"
 location.href = "fangwu.asp"
 </script>
<%
		       rs.close
set rs=nothing
conn.close
set conn=nothing
		end if
   rs.close
   end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
