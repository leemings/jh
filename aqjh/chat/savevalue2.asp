<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../paodian.asp"-->
<!--#include file="../const3.asp"-->
<!--#include file="../chk.asp"-->
<!--#include file="sjfunc/func.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
nowinroom=session("nowinroom")
id=LCase(trim(request.querystring("id")))
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
ChkPost()
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if aqjh_disproxy=1 then Chkproxy()
%>
<script language="JavaScript"> 
if(window.top==window.self){var i=1;while (i<=50){window.alert("������ʲôѽ�����ң������ǲ��еģ�ȥ����ȥ�ɣ�����������50�Σ���");i=i+1;}top.location.href="../exit.asp"}
</script>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_yjdh=1 then Chkyjdh()
inroom=session("nowinroom")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������Ҳſ��Բ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
%>
<html>
<head>
<META http-equiv='content-type' content='text/html; charset=gb2312'>
<title>���澭��ֵ</title>
<link rel="stylesheet" href="READONLY/STYLE.CSS">
<style>
body{
CURSOR: url('boy.cur');
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff}
td{font-size:9pt;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed" topmargin="5" text="#FFFFFF">
<%
useronlinename=Application("aqjh_useronlinename"&inroom)
if aqjh_name="" or Session("aqjh_inthechat")<>"1" or Instr(useronlinename," "&aqjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
ip=Request.ServerVariables("REMOTE_ADDR")
If ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t
mycd=DateDiff("n",Session("aqjh_savetime"),now())
addvalue=mycd
addvalue1=mycd*aqjh_paofen
Session("aqjh_savetime")=now()
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� WHERE  ����='" & aqjh_name &"'",conn,1,3
mysl=clng(DateDiff("n",now(),rs("slsj")))
mysls="["&rs("sl")&"]"
olddj=rs("�ȼ�")
jsr=rs("������")
zhuans=rs("ת��")
if jsr<>"��"  then
jiang=1
else
jiang=0
end if
if olddj < 90 and rs("ת��")<1 then
aqjh_paofencd=(aqjh_paofencd*2.5)
else
aqjh_paofencd=aqjh_paofencd
end if
if mysl>0 then
	Select Case rs("sl")
	case "����"
		aqjh_paofencd=aqjh_paofencd*1.5
	case "����"
		aqjh_paofencd=aqjh_paofencd*2.5
	case "����"
		aqjh_paofencd=aqjh_paofencd*2
	case "����"
		aqjh_paofenyin=aqjh_paofenyin*5	
	case "˥��"
		aqjh_paofencd=int(aqjh_paofencd/2)
	case "����"
		aqjh_paofen=-(aqjh_paofen*10)
	case "����"
		aqjh_paofenyin=-(aqjh_paofenyin*5)
	end select
else
	rs("sl")="��"
end if

if rs("��Ա")=true and clng(DateDiff("d",now(),rs("��Ա����")))>0 then
	pd=paodian
	pdstr="�ݵ��ƻ�Ա"
else
	pd=1
	pdstr="���ݵ��Ա"
end if
sf=rs("ʦ��")
hydj=rs("��Ա�ȼ�")
zddj=rs("�ȼ�")
dj1=rs("�ȼ�")
sfwg=1
if sf<>"" and sf<>"��" then
	if Instr(LCase(Application("aqjh_useronlinename"&inroom))," "&LCase(sf)&" ")=0 then
		sfwg=1
	else
		sfwg=2
	end if
end if
jhhy=1
Select Case hydj
case 1
	jhhy=hy1
case 2
	jhhy=hy2
case 3
	jhhy=hy3
case 4
	jhhy=hy4
case 5
	jhhy=hy5
case 6
	jhhy=hy6
case 7
	jhhy=hy7
case 8
	jhhy=hy8
case else
	jhhy=1
end select
rs("����")=rs("����")+int(addvalue1*sfwg*jhhy)
rs("����")=rs("����")+int(addvalue1*jhhy*aqjh_paofenyin)
rs("�书")=rs("�书")+int(addvalue1/3*sfwg*jhhy)
rs("����")=rs("����")+int(addvalue1/3*jhhy)
rs("����")=rs("����")+int(addvalue1/3*jhhy)
rs("����")=rs("����")+int(addvalue1/3*jhhy)
rs("�Ṧ")=rs("�Ṧ")+int(addvalue1/3*jhhy)
rs("���")=rs("���")+int(addvalue1/16*jhhy)
rs("Ԫ��")=rs("Ԫ��")+int(addvalue1*sfwg*jhhy)
rs("��")=rs("��")+int(addvalue1/40)
rs("ľ")=rs("ľ")+int(addvalue1/40)
rs("ˮ")=rs("ˮ")+int(addvalue1/40)
rs("��")=rs("��")+int(addvalue1/40)
rs("��")=rs("��")+int(addvalue1/40)
bl=rs("��ż")
rs("allvalue")=rs("allvalue")+int(addvalue*jhhy*aqjh_paofencd)*pd
rs("�ݶ�����")=rs("�ݶ�����")+addvalue
prevtime=CDate(rs("lasttime"))
if DateDiff("m",prevtime,now())=0 then
	rs("mvalue")=rs("mvalue")+int(addvalue*jhhy*aqjh_paofencd)*pd
else
	rs("mvalue")=addvalue*pd
end if
rs("lasttime")=sj
rs("lastip")=ip
rs.Update
aqjh_value=rs("allvalue")
aqjh_mvalue=rs("mvalue")
dengji=int(sqr(aqjh_value/50))
rs("�ȼ�")=dengji
rs.Update
Session("aqjh_grade")=rs("grade")
Session("aqjh_jhdj")=rs("�ȼ�")
zddj=rs("�ȼ�")
%>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���,״̬,ת��,�Ա� FROM �û� WHERE ����='" & aqjh_name &"'",conn
if rs("״̬")="��" then
call boot(aqjh_name,"��") 
Response.Write "<script language=JavaScript>{alert('��������轣��µ��б�����!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("���")>=50000 then
conn.execute "update �û� set ״̬='��',�¼�ԭ��='���Լ�������|"&fn1&"' where ����='" & aqjh_name & "'"
call boot(aqjh_name,aqjh_name&"��"&aqjh_name&"�ķ�������") 
Response.Write "<script language=JavaScript>{alert('����Ѵ�50000�㱻�Լ���������!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

'�ж�����
if zddj>dj1 then
if jiang=1 then
if zhuans>1 then 
conn.execute "update �û� set ���=���+2,����=����+100  where ����='" & jsr & "' "
mess="<br><font color=blue><font color=red>["&aqjh_name&"]</font>�ﵽϵͳ��������$<font color=red>ת��2������</font>$,��˽�����<font color=red>["&jsr&"]</font>��<font color=red>["&aqjh_name&"]</font>ÿ��������ʱ��õ�ϵͳ����<font color=brown>���</font>2��,<font color=brown>����</font>100�㣡��</font>"
else
mess="<br><font color=blue><font color=red>["&aqjh_name&"]</font>ϣ���������������ֽ�����ϵͳ�н�����Ŷ!</font>"
end if
else 
mess="<br><font color=blue><font color=red>["&aqjh_name&"]</font>û�н����ˣ������������Լ��н�����Ŷ!</font>"
end if 
if rs("�Ա�")="��" then
says="<img src=img/up.gif><bgsound src=wav/system.wav loop=1><font color=green>��������Ϣ����ϲ</font>��˧��<font color=red>["& aqjh_name &"]</font><font color=green>������ֲ�и��Ŭ��,�������ˣ�<font color=red><b><img src=img/sjl.gif>"&dj1&"��<img src=img/sjl.gif>"&zddj&"</b></font>,<font color=blue>��ת��"&zhuans&"�Ρ�</font>,Ը��Խ��Խ˧��Խ��Խ����������������ŮŶ��˧�����Ŷ......</font>"&mess
else
says="<img src=img/up1.gif><bgsound src=wav/system.wav loop=1><font color=green>��������Ϣ����ϲ</font>����Ů<font color=red>["& aqjh_name &"]</font><font color=green>������ֲ�и��Ŭ��,�������ˣ�<font color=red><b><img src=img/sjl.gif>"&dj1&"��<img src=img/sjl.gif>"&zddj&"</b></font>,<font color=blue>��ת��"&zhuans&"�Ρ�</font>,ף����ŮԽ��ԽƯ�����ٺ٣���Ů����Ŷ......</font>"&mess
end if
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& inroom & ");<"&"/script>"
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
If olddj < 90 and rs("ת��")<1 Then
Response.Write "�ȼ�����90��,0��ת������,�ݵ㷭2.5��!<br>"
end if
%>
<table width="100%" border="1" cellspacing="0" cellpadding="5" align="center" height="267" bordercolor="#F7DEAC">
 <tr valign="top">
<td height="260" style="filter:dropshadow(color=black, offx=1,offy=1);">
<p align="center"> <%=aqjh_name%><br>�� ǰ ״ ̬</p>
<div align="center"><font color=yellow>
<%if sf<>"" and sf<>"��" and sfwg=2 then%> 
	ʦ��:<font color=red><%=sf%></font>�ڳ��书������������<br> 
<%end if%>
<%if hydj>0  then%>����<%=hydj%>����Ա<br>�ݵ�Ϊ<%=jhhy%>��<%else%>�ǻ�Ա�ݵ㲻����<%end if%> 
<br><font color=red><%if olddj < 90 and rs("ת��")<1 then%>�ȼ�С��90,0��ת����,�ݵ�Ϊ2.5��,����Ŷ<%end if%>
</font><br><%=pdstr%>
<%if mysl>0 then%><HR><%=mysls%><br>����<%=mysl%>����<HR><%end if%> 
</div></font>
<p align="left"> 
<%sjjf=(zddj+1)*(zddj+1)*50-aqjh_value%>
ʱ��:<%=mycd%>����<br> 
����:<%=aqjh_grade%> ��<br> 
ս��:<%=int(sqr((aqjh_value/50)))%> ��<br> 
�»�:<%=aqjh_mvalue%><br>
�ܻ�:<%=aqjh_value%><br>
����:<font color=red><%=sjjf%></font>��<br>
�ݵ�:<font color=red><%=int(addvalue*aqjh_paofencd*jhhy)*pd%></font>��<br>
<br>
�㹲����<br>
����:<%if sfwg=2 then%><font color=red><%=int(addvalue1*sfwg*jhhy)%></font>��<br><%else%><%=int(addvalue1*sfwg*jhhy)%>��<br><%end if%> 
����:<%=addvalue1*aqjh_paofenyin%>��<br> 
�书:<%if sfwg=2 then%><font color=red><%=int(addvalue1/2*sfwg*jhhy)%></font>��<br><%else%><%=int(addvalue1/2*sfwg*jhhy)%>��<br><%end if%> 
����:<%=int(addvalue1/3*jhhy)%>��<br> 
����:<%=int(addvalue1/3*jhhy)%>��<br> 
����:<%=int(addvalue1/3*jhhy)%>��<br> 
�Ṧ:<%=int(addvalue1/3*jhhy)%>��<br>
���:<%=int(addvalue1/16*jhhy)%>��<br> 
��:<font color=red><%=int(addvalue1/40)%></font>��<br>
ľ:<font color=red><%=int(addvalue1/40)%></font>��<br>
ˮ:<font color=red><%=int(addvalue1/40)%></font>��<br>
��:<font color=red><%=int(addvalue1/40)%></font>��<br>
��:<font color=red><%=int(addvalue1/40)%></font>��<br>
<br>
<%if bl<>"��" and bl<>"" then
peioujia=int(addvalue1/4)
conn.execute "update �û� set ����=����+" & addvalue1/2 & ",����=����+" & peioujia & ",�书=�书+" & peioujia & ",����=����+" & peioujia & ",����=����+"  & peioujia & ",�Ṧ=�Ṧ+"  & peioujia & ",����=����+" & peioujia & " where ����='" & bl & "'"%>
����<font color=red> <%=bl%></font><br> 
����:<%=int(addvalue1/4)%>��<br> 
����:<%=int(addvalue1/4)%>��<br> 
�书:<%=int(addvalue1/4)%>��<br> 
����:<%=int(addvalue1/4)%>��<br> 
����:<%=int(addvalue1/4)%>��<br> 
Ԫ��:<%=int(addvalue1/2*jhhy)%>��<br>
����:<%=int(addvalue1/4)%>��<br> 
�Ṧ:<%=int(addvalue1/4)%>��<br> 
<%end if
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</td></tr></table></html>