<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Response.Buffer=true
Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
name=Session("aqjh_name")
nowinroom=session("nowinroom")
if name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='"&name&"'",conn,2,2
if rs("����")<>"ȫ����" then 
Response.Write "<script language=JavaScript>{alert('��ʾ�����Ѿ���ȫ����,�����ٽ����ˣ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("���")<200000 then 
Response.Write "<script language=JavaScript>{alert('��ʾ����Ľ�Ҳ���,������ȫ������Ҫ200000��ң�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if 

if rs("ת��")<9 then 
Response.Write "<script language=JavaScript>{alert('��ʾ����ĵȼ�����,��������������Ҫת��9�Σ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
conn.execute"update �û� set ����='ȫ����',���=���-200000,������=������+130,������=������+130,�书��=�书��+130 where ����='"&name&"'"
attack="<font color=red><b>�����ֽ���ʱ��������</b></font><font color=ff00ff>��ϲ<b><font color=red>��" & name & "��</font></b>ʱ������ȫ���˳ɹ�,�������������书���޸���130,<font color=red>ϣ������Ŭ��,���ս��������߲�,���ף�أ�</font></font>"
rs.close
set rs=nothing
conn.close
set conn=nothing
says=attack 
says=replace(says,chr(39),"\'") 
says=replace(says,chr(34),"\"&chr(34)) 
act="��Ϣ" 
towhoway=0 
towho="���" 
addwordcolor="660099" 
saycolor="008888" 
addsays="��" 
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
Response.Write "<script Language=Javascript>alert('��ϲ��������ȫ���˳ɹ�-��ø�����+130,ϣ������Ŭ����');location.href = 'index.asp';</script>"
Response.End 
rs.close
conn.close
set rs=nothing
set conn=nothing
%>

<%
Sub ErrALT(Str)
	Response.Write "<META http-equiv=Content-Type content=text/html;charset=gb2312><script>alert('" & Str &"');location.href='index.ASP';</script>"	
	Response.End
End Sub
%>