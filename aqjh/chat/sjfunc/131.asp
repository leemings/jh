<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_name=Session("aqjh_name")
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
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>�������ҡ�<font color=" & saycolor & ">"+sljb()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function sljb()
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if Session("sljb_time")="" then Session("sljb_time")=rs("lasttime")
mycd=DateDiff("n",Session("sljb_time"),now())
jbmoney=2
hyjb=2
if mycd<60 then
    if DateDiff("n",rs("lasttime"),now())>=60 then
       sljb="����һ��Сʱǰ�Ѿ������"
    else
       sljb="�������ݽ���ʱ�仹����һСʱ(����:"& (60-mycd) &"�ֵ�1Сʱ),����������ݵ�1Сʱ����(�˳�,����,����ʱ����0)������ȡ"& jbmoney &"�����,���ѻ�Ա����ͨ���"& hyjb &"��!"
    end if
    Response.Write "<script language=JavaScript>{alert('"&sljb&"');}</script>"
    Response.End
end if
hy=rs("��Ա�ȼ�")
mycd=int(mycd/60)
if rs("��Ա�ȼ�")=4 then
	myjbsl=jbmoney*hyjb*mycd
        conn.execute "update �û� set ���=���+"& myjbsl &" where ����='" & aqjh_name & "'"
        Session("sljb_time")=now()
	sljb="##Ϊ<font color=blue>"& hy &"</font>����Ա���˸��ѻ�Ա���ڽ���Ŭ���ݵ�<font color=red><b>"& mycd &"</b></font>Сʱ,�õ����:<font color=red>"& myjbsl &"</font>��,##�����Ŭ��,֧�ֿ��ֽ����ķ�չ!"
else
	myjbsl1=jbmoney*mycd
        conn.execute "update �û� set ���=���+"& myjbsl1 &" where ����='" & aqjh_name & "'"
        Session("sljb_time")=now()
	sljb="##Ϊ<font color=blue>"& hy &"</font>����Ա������ѻ�Ա���ڽ���Ŭ���ݵ�<font color=red><b>"& mycd &"</b></font>Сʱ,�õ����:<font color=red>"& myjbsl1 &"</font>��,##�����Ŭ��,֧�ֿ��ֽ����ķ�չ!"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>