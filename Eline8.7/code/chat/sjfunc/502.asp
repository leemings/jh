<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�����wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>��ָ���书��<font color=" & saycolor & ">"+chiwu()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����chiwu
function chiwu()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select ����,ְҵ,����,�ȼ�,����,�书,�书��,����ʱ�� FROM �û� WHERE ����='" & sjjh_name &"'" ,conn,2,2
if rs("ְҵ")<>"��ʦ" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ְҵ������ʦ�����ܲ�����');}</script>"
	Response.End
end if

sj=DateDiff("s",rs("����ʱ��"),now())
if sj<(int(rnd*600)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*600)+1)-sj
	Response.Write "<script language=JavaScript>{alert('�������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
if rs("����")<5000  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ����5000�ſ��Բ�����');}</script>"
	Response.End
end if
if rs("�书")<rs("�ȼ�")*sjjh_wgsx+3800+rs("�书��") then
	conn.execute "update �û� set �书=�书+1000,����=����-200,����ʱ��=now() where ����='" & sjjh_name &"'"
	chiwu="��ʦ""##<bgsound src=wav/dz.wav loop=1>Ϊ�������Լ����书����ϧð���߻���ħ��Σ�ս�����ƶ�<font color=red>���ڳɹ��ָ��书</font>1000�㣬�����½�<font color=red>+3500</font>100�㣬����һɽҪ��һɽ�ߣ�"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("�书��")<=rs("�ȼ�")*1500 then
	conn.execute "update �û� set �书��=�书��+50,����=����-500,����ʱ��=now() where ����='" & sjjh_name &"'"
	chiwu="��ʦ""##<bgsound src=wav/dz.wav loop=1>Ϊ�������Լ����书����ϧð���߻���ħ��Σ�ս�����ƶ�<font color=red>���ڳɹ��ָ��书</font>50�㣬�����½�<font color=red>-500</font>�㣬����һɽҪ��һɽ�ߣ�"
else
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����������������ˣ������˼���ָ���书�ɣ�');}</script>"
	Response.End
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>







