<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('ˢǮ����ϲ�����𣿵㰡��ˢ������');i=i+1;}top.location.href='../../exit.asp'}</script>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if aqjh_jhdj<15 then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻��15��������������������!');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,״̬,���� from �û� where ����='" & aqjh_name &"'",conn,2,2
if rs("����")<10000 then
    Response.Write "<script language=JavaScript>{alert('��ʾ�������������1����������!');window.close();}</script>"
	Response.End 
end if
if rs("״̬")="��" then
    Response.Write "<script language=JavaScript>{alert('��ʾ�����Ѿ����ˣ���ȥ�����!');window.close();}</script>"
	Response.End 
end if
if rs("����")<500 then
    Response.Write "<script language=JavaScript>{alert('��ʾ����ĵ��²���500��˭��������?');window.close();}</script>"
	Response.End 
end if
Session("diaoyu")=now()
Response.Redirect "diao.asp"
%>