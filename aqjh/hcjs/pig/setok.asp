<%
Sub Msg (v)
 Response.Write "<script Language=JavaScript>alert('" & v & "');history.back();</script>"
 Response.End
End Sub
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_grade<10 or aqjh_name<>application("aqjh_user") then
	Response.Write "<b>[����ʧ��]</b><p>��û�и���ͶƱʱ���Ȩ�ޣ�"
	Response.End
end if
tp_dj=Trim(Request.Form("tp_dj"))
tp_start=Trim(Request.Form("tp_start"))
tp_end=Trim(Request.Form("tp_end"))
bm_start=Trim(Request.Form("bm_start"))
bm_end=Trim(Request.Form("bm_end"))
if IsNumeric(tp_dj)=false then
	Msg "[����ʧ��]ͶƱ����Ҫ�ȼ���ʽ����"
end if
if IsDate(bm_start)=false or len(bm_start)<>19 then
	Msg "[����ʧ��]��ʼ����ʱ���ʽ����"
end if
if IsDate(tp_start)=false or len(tp_start)<>19 then
	Msg "[����ʧ��]��ʼͶƱʱ���ʽ����"
	Response.end
end if
if IsDate(bm_end)=false or len(bm_end)<>19 then
	Msg "[����ʧ��]��ֹ����ʱ���ʽ����"
end if
if IsDate(tp_end)=false or len(tp_end)<>19 then
	Msg "[����ʧ��]��ֹͶƱʱ���ʽ����"
end if
if CDate(tp_end)<=CDate(tp_start) then
	Msg "[����ʧ��]��ֹͶƱʱ��������ڿ�ʼͶƱʱ��"
end if
if CDate(bm_end)<=CDate(bm_start) then
	Msg "[����ʧ��]��ֹ����ʱ��������ڿ�ʼ����ʱ��"
end if
if CDate(tp_start)<=CDate(bm_start) then
	Msg "[����ʧ��]��ʼͶƱʱ��������ڿ�ʼ����ʱ��"
end if
if CDate(tp_end)<=CDate(bm_end) then
	Msg "[����ʧ��]��ֹͶƱʱ�����������ֹ����ʱ��"
end if
%>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
sql="update ��Ȧ���� set ͶƱ��ʼ='" & tp_start & "',ͶƱ����='" & tp_end & "',������ʼ='"&bm_start&"',��������='"&bm_end&"',�ȼ�="&tp_dj
set rs=connt.execute(sql)
set rs=nothing
connt.close
set connt=nothing
Response.Write "<script Language=JavaScript>alert('���óɹ�');window.location.href ='hfyb.asp';</script>"
Response.End
%>