<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�������ɡ�wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_hy then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������Ҫ["&jhdj_hy&"]���ſ��Բ�����');}</script>"
	Response.End
end if
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
says="<font color=green>���������ɡ�<font color=" & saycolor & ">"+qiuhun(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'��������
function qiuhun(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM �û� WHERE ���<>'����' and ����='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('��ʾ���������Ҳ���������ϻ����������ţ�');window.close();</script>"
		response.end
end if
if rs("����")="" or rs("����")="����" or rs("����")="��" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('��ʾ���㲢�����ɣ���');window.close();</script>"
		response.end
end if
if rs("����")="����" or rs("����")="����"  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('��ʾ�����ǳ����ˣ���������ɱ�ֲ������������뿪!����');window.close();</script>"
		response.end
end if
if rs("����")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������5���������������Ѷ�û��ѧʲô���ɰ���');}</script>"
	Response.End
end if
rs.close
conn.execute "update �û� set ����=����-50000 where ����='" & sjjh_name &"'"
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("sjjh_online")=to1
Application.UnLock
randomize()
regjm=int(rnd*3348998)
qiuhun="[##]��{%%}�������ɣ�<img src='img/29.gif'>"&fn1&"<input type=button value='ͬ��' onClick=javascript:tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('lipai.asp?name=" & sjjh_name &"&yn=1&to1="&to1&"','d') name=tongyi"&regjm&"><input type=button value='��ͬ��' onClick=javascript:;tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('lipai.asp?name=" & sjjh_name &"&yn=0&to1="&to1&"','d') name=tongyi"&regjm+1&">"
end function
%>
