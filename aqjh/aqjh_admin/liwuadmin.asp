<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../showchat.asp"-->

<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"

if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');</script>"
	Response.End
end if
r=request.querystring("username")
if r<>"" then

Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn1.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../liwu/sdlw.mdb")


select case r
	case 1 '�������������
		sql="delete * from �������� where id>0"
		says="<font color=red><marquee height=50 behavior=alternate loop=200 direction=up><marquee behav><ior=alternate><font size=6><b>��������Ϣ��"&aqjh_name&"վ����������,�ȼ�����55��������ȡ�������￩" & "</b></font></marquee></marquee>"
	Response.Write "<script Language=Javascript>alert('��ϲ,����ɹ�');</script>"

	case 2 '������������
		sql="delete * from �������� where id>0"
	says="<font color=red><marquee height=50 behavior=alternate loop=200 direction=up><marquee behav><ior=alternate><font size=6><b>��������Ϣ��"&aqjh_name&"վ����������������,�������������ȡ���￩" & "</b></font></marquee></marquee>"
	Response.Write "<script Language=Javascript>alert('��ϲ,����ɹ�');</script>"

	case 3 '������ĩ����
		sql="delete * from �������� where id>0"
	says="<font color=red><marquee height=50 behavior=alternate loop=200 direction=up><marquee behav><ior=alternate><font size=6><b>��������Ϣ��"&aqjh_name&"վ����������������,�������������ȡ���￩" & "</b></font></marquee></marquee>"
	Response.Write "<script Language=Javascript>alert('��ϲ,����ɹ�');</script>"

	case 4 '������ĩ����
		sql="delete * from ��ĩ���� where id>0"
		says="<font color=red><marquee height=50 behavior=alternate loop=200 direction=up><marquee behav><ior=alternate><font size=6><b>��������Ϣ��"&aqjh_name&"վ����������ĩ����,�����28��29��30�ſ�����ȡ���￩" & "</b></font></marquee></marquee>"
	Response.Write "<script Language=Javascript>alert('��ϲ,����ɹ�');</script>"

	case 5 '����վ�����ŵ�����
		sql="delete * from վ������ where id>0"
		says="<font color=red><marquee height=50 behavior=alternate loop=200 direction=up><marquee behav><ior=alternate><font size=6><b>��������Ϣ��"&aqjh_name&"վ������λ����׼���˷ḻ������,�Ͽ�ȥ��ȡվ�������" & "</b></font></marquee></marquee>"
			Response.Write "<script Language=Javascript>alert('��ϲ,���ųɹ�');</script>"
	
	case 6 '������������
		sql="delete * from �������� where id>0"
		says="<font color=red><marquee height=50 behavior=alternate loop=200 direction=up><marquee behav><ior=alternate><font size=6><b>��������Ϣ��"&aqjh_name&"վ����������������,������������ȡ���￩" & "</b></font></marquee></marquee>"
			Response.Write "<script Language=Javascript>alert('��ϲ,����ɹ�');</script>"
	
	case 7 '�����������
		sql="delete * from �������� where id>0"
		says="<font color=red><marquee height=50 behavior=alternate loop=200 direction=up><marquee behav><ior=alternate><font size=6><b>�����������ʦ�ڵ���,"&aqjh_name&"վ�������׼���˷ḻ������Ͽ�ȥ��ȡ�������￩!" & "</b></font></marquee></marquee>"
			Response.Write "<script Language=Javascript>alert('��ϲ,����ɹ�');</script>"


case else

	Response.Write "<script Language=Javascript>alert('�����ȷ�����ӽ���');</script>"
	Response.End

end select


set rs1=conn1.execute(sql)



set rs1=nothing
conn1.close
set conn1=nothing

call showchat(says)
end if



%>
<html>
<head>
<title>�������</title>
<LINK href=css/css.css type=text/css rel=stylesheet><body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body text="#000000" background="../chat/f3.gif">
<div align="center">
<table width="603" border="0" cellspacing="0" cellpadding="0" bordercolor="#0A246A" bgcolor="#808000">
  <tr>

    <td width="5%" align=center ><a href="?username=1"><font color="#800000">������������</font></a></td>
    <td width="5%" align=center ><a href="?username=2"><font color="#800000">������������</font></a></td>
    <td width="5%" align=center ><a href="?username=3"><font color="#800000">������ĩ����</font></a></td>
    <td width="5%" align=center ><a href="?username=4"><font color="#800000">������ĩ����</font></a></td>
    <td width="5%" align=center ><a href="?username=5"><font color="#800000">����վ������</font></a></td>
    <td width="5%" align=center ><a href="?username=6"><font color="#800000">������������</font></a></td>
    <td width="5%" align=center ><a href="?username=7"><font color="#800000">���Ž�������</font></a></td>


  </tr>
</table>
</div>
<p align=center>
��ע��:����������55��������ȡ,���塢���������յĶ����ڹ̶�������(������ʱ��,�����ķ�����ʱ�䲻׼����ϵ�������ṩ��),վ�������ǵ���֮��Ϳ����ô��ȥ��ȡ��,�е��˿��ܵ�ʱû����ȡ,���˼������ȡ,Ҳ������(������������ʱ������ĳ������վ��������,��ʵ����ǰû����ȡ��)</p>
<p align=center>
��</p>
<p align=center>
�˹���������������û�����Զ����������鷳��λվ����������</p>
<p align=center>
ע�⣺����������Բ��������������벻Ҫ�����涨��ʱ���������������������ֹ����������������ᵼ���ظ���ȡ</p>
<p align=center>
վ���뷢վ�������ʱ��͵����һ���������һ�����ʾ�ģ�û����ȡ��������Ҳ���Ǿ����Ͻ����ģ���������˵����ȡʲô��������������</p>
</body>
</html>