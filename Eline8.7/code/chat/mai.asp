<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������Ҳſ��Բ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if InStr(Application("sjjh_mai"),"|")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������ⲻ���Թ���');}</script>"
	Response.End
end if
zt=split(Application("sjjh_mai"),"|")
if ubound(zt)<>7 then 
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������ⲻ���Թ���');}</script>"
	Response.End
end if

'����0����Ǯ1������2�����3����Ʒ��4������5��ʱ��6
name=trim(zt(0))
yin=abs(int(clng(zt(1))))
bz=trim(zt(2))
lb=trim(zt(3))
zswupin=trim(zt(4))
wusl=abs(int(clng(zt(5))))
towho=zt(7)
Application.Lock
Application("sjjh_mai")=""
Application.UnLock
if sjjh_jhdj<=12 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������12�����ϲſ��Բ�����');}</script>"
	Response.End 
end if
if name=sjjh_name then
	Response.Write "<script language=JavaScript>{alert('��ʾ���Լ����Ķ���ҲҪ��');}</script>"
	Response.End 
end if
if towho<>"���" and towho<>sjjh_name then
	Response.Write "<script language=JavaScript>{alert('��ʾ������Ʒ������["&towho&"]�ģ�����Ȩ����');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select "&lb&" from �û� where ����='"&name&"'",conn,3,3
if  mywpsl(rs(lb),zswupin)<wusl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&name&"]����������("&wusl&")��,���򲻳ɹ���');}</script>"
	Response.End
end if
zstemp=abate(rs(lb),zswupin,wusl)
rs.close
rs.open "select "&bz&","&lb&" from �û� where ����='"&sjjh_name&"'",conn,3,3
if rs(bz)<yin then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����["&bz&"]����"&yin&"���򲻳ɹ���');}</script>"
	Response.End
end if
zstemp1=add(rs(lb),zswupin,wusl)
conn.execute "update �û� set "&lb&"='"&zstemp&"',"&bz&"="&bz&"+"&yin&" where  ����='"&name&"'"
conn.execute "update �û� set "&lb&"='"&zstemp1&"',"&bz&"="&bz&"-"&yin&" where  ����='"&sjjh_name&"'"
mai="<font color=red><b>��������Ϣ��</b></font>[##]��{%%}������Ʒ["&zswupin&"]"&wusl&"���ɹ�������"&bz&"��"&yin&"��"
rs.close
set rs=nothing
conn.close
set conn=nothing
zj="<a href=javascript:parent.sw('[" & sjjh_name & "]'); target=f2>"& sjjh_name & "</a>"
br="<a href=javascript:parent.sw('[" & name & "]'); target=f2>" & name & "</a>"
mai=Replace(mai,"##",zj,1,3,1)
mai=Replace(mai,"%%",br,1,3,1)
says=mai
says=replace(says,chr(39),"\'")
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
%>
