<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_roominfo=split(Application("aqjh_room"),";")
useronlinename=Application("aqjh_useronlinename"&nowinroom)
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'�����뿪������Ѩ�ж�
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>��NPC������<bgsound src=wav/FX40A.wav><font color=" & saycolor & ">"+zynpc(mid(says,i+1),towho)+"</font>"
call chatsay("NPC����",towhoway,towho,saycolor,addwordcolor,addsays,says)
function zynpc(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn,2,2
tlj=(rs("�ȼ�")*aqjh_tlsx+5260)+rs("������")
if rs("����")<500000 then
   zynpc="##:��û��50�������ӣ�������ҽ���������ƣ������ǽ�Ǯ����ˣ�"
   exit function
end if
rs.close
rs.open "select * from npc where n����='" & to1 &"'",conn,1,3
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��NPC["&to1&"]�������ڻ���������NPC�����˰ɣ�');}</script>"
	Response.End	
end if
if InStr(";" & Application("aqjh_npc"), ";" &to1& "|")=0 then
   zynpc="��ʾ��NPC[%%]�����ߣ�����ʹ����������"
   exit function
end if
if rs("n����")<>aqjh_name then
  zynpc="##:�㲻��NPC[%%]�����ˣ��������������������ܷ�����"
  exit function
end if
n_tl=rs("n����")
if n_tl<0 then
   n_dengji=int(sqr(int(rs("n����")/50)))
   rs("n����")=n_dengji*5000
   rs("n����")=n_dengji*100
   rs("n����")=n_dengji*150
   rs("n����")=n_dengji*150
   rs.Update
   zynpc= "##��%%ʹ��ħ��<b>[������]</b>,��%%�ӽ�����������������ʱ��״̬�Ѿ�ȫ���ָ������NPC�ֳ��������յ�ɱ������ҿ���С���ˣ�"
   exit function
else
  if n_tl>=int(tlj/2) then
    zynpc= "##����NPC%%��������������1/2��,��Ȼʹ����������,����û��Ч��,��ʹ���˴�ħ��!"
    Exit Function
 else
    zynpc= "##��%%ʹ��ħ��[������],����%%NPC�������ָ������޵�1/2��!"
    rs("n����")=Int(tlj/2)
    rs.Update
    exit function
 end if
end if
end function
%>
