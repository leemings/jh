<%@ LANGUAGE=VBScript codepage ="936" %>
<%

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"


%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"


sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"

id=trim(request.form("id"))
qingren=trim(request.form("name"))
my=sjjh_name
%><!--#include file="dadata.asp"-->
<%
sql="select * from �û� where ����='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
response.write "�㲻�ǽ������ˣ����ܶ�������"
conn.close
response.end
else
sql="select * from �û� where ����='" & qingren & "' and ����<>'" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
conn.close
Response.Write "<script language=javascript>alert('������û�������������!');history.back();</script>"
response.end
else
qingren=rs("����")
sql="SELECT * FROM ��� where ID=" & id
Set Rs=connt.Execute(sql)
wu=rs("�����")
yin=rs("�ۼ�")
lx=rs("����")
nl=rs("����")
tl=rs("����")
jb=rs("���")
sl=rs("����")%>

<%
sql="select * from �û� where ����='" & my & "'"
rs=conn.execute(sql)
if yin<=rs("����") then
sql="update �û� set ����=����-" & yin & " where ����='" & my & "'"
rs=conn.execute(sql)%>
			
<%
sql="select * from ����б� where �����='" & wu & "' and ӵ����='" & my & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then
sql="insert into ����б�(�����,ӵ����,����,����,����,���,����,�ۼ�,�μ���,ʱ��) values ('"&wu&"','"&my&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",'"&qingren&"',now())"
rs=connt.execute(sql)
connt.close
set connt=nothing
if Instr(Application("sjjh_useronlinename"&session("nowinroom"))," "&qingren&" ")<>0 then

says="<font color=red>������Ϣ��</font><b><font color=green>"&my&"�ڽ�����Ƶ��"&wu&"��Ϊ"&qingren&"����<font color=red>��"&lx&"��</font>��ᣬ"&qingren&"��ȥѽ���Ǻ�..</font></b>"			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
end if
			
Response.Write "<script language=javascript>alert('��ϲ����Ϊ"&qingren&"���ľ����Ѿ�׼�����ˣ�');window.close();</script>"
response.end
else
connt.close
Response.Write "<script language=javascript>alert('���ܶ����磬ԭ�����Ѷ���ͬ�����ľ�ϯ��');history.back();</script>"
response.end
				
end if
else
connt.close
Response.Write "<script language=javascript>alert('���ܶ����磬ԭ���������������');history.back();</script>"
response.end
end if
rs.close
set rs=nothing
end if
end if
%>