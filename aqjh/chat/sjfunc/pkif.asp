<%
'�Թ�����ж�
toname=Trim(Request.Form("towho"))
if Instr(Application("aqjh_guibin"),"|" & aqjh_name & "|")<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ǳ��������,���ò����κζ�Թ��');}</script>"
	Response.End
end if
if Instr(Application("aqjh_guibin"),"|" & toname & "|")<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է��Ǳ����������');}</script>"
	Response.End
end if
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
'�Է��е��ж�
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻���Է��У�');}</script>"
	Response.End
end if
'��pkʱ����ж�
f=Minute(time())
pktime=cint(chatinfo(10))
if f>pktime then
        Response.Write "<script language=JavaScript>{alert('��ʾ������PK���ʱ��ΪÿСʱǰ["&pktime&"]�֣�');}</script>"
	response.end
end if
'�Ծ�ս�����ж�
if chatinfo(0)="��ս����" and Hour(time())=18 then
        Response.Write "<script language=JavaScript>{alert('��ʾ����ս��PKʱ��δ����');}</script>"
	response.end
end if
%>