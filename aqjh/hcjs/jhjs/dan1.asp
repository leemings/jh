<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');parent.history.go(-1);}</script>" 
Response.End 
end if
if Session("aqjh_inthechat")<>"1" then %>
<script language="vbscript">
MsgBox "������������ٽ��뵱�̲�������"
window.close()
</script>
<%response.end
end if
wpname=trim(lcase(request("wpname")))
if InStr(wpname,"or")<>0 or InStr(wpname,"=")<>0 or InStr(wpname,"`")<>0 or InStr(wpname,"'")<>0 or InStr(wpname," ")<>0 or InStr(wpname,"��")<>0 or InStr(wpname,"'")<>0 or InStr(wpname,chr(34))<>0 or InStr(wpname,"\")<>0 or InStr(wpname,",")<>0 or InStr(wpname,"<")<>0 or InStr(wpname,">")<>0 then Response.Redirect "../../error.asp?id=54"
lx=lcase(trim(request("lx")))
if InStr(lx,"or")<>0 or InStr(lx,"=")<>0 or InStr(lx,"`")<>0 or InStr(lx,"'")<>0 or InStr(lx," ")<>0 or InStr(lx,"��")<>0 or InStr(lx,"'")<>0 or InStr(lx,chr(34))<>0 or InStr(lx,"\")<>0 or InStr(lx,",")<>0 or InStr(lx,"<")<>0 or InStr(lx,">")<>0 then Response.Redirect "../../error.asp?id=54"
dansl=request("dansl")
allsl=request("allsl")
if not IsNumeric(dansl) or not IsNumeric(allsl) then
	Response.Write "<script language=JavaScript>{alert('����:��������ȷ����ȷ����д����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if dansl>allsl then
	Response.Write "<script language=JavaScript>{alert('����:������ô�ණ���ɵ���');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if allsl=0 or dansl=0 then
	Response.Write "<script language=JavaScript>{alert('����:û���������㲻��������������');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����ʱ��,w1,w2,w4,w5,w6,w7 from �û� where ����='"&aqjh_name&"'",conn
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Բ���ϵͳ���ƣ����["&s&"����]�ٲ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
select case lx
case "��Ƭ"
	wpsl=dansl
	zstemp=abate(rs("w5"),wpname,dansl)
	if wpsl>0 then
		rs.close
		rs.open "select h from b where a='"&wpname&"'",conn
		if rs.bof and rs.eof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('û�����뵱��������Ʒ��������ʲô�����վ����ϵ��');location.href = 'javascript:history.go(-1)';}</script>"
			response.end
		else
			bb=rs("h")*wpsl
			yin=int(bb/4)
			conn.execute "update �û� set w5='"&zstemp&"',��Ա��=��Ա��+" & yin & ",����ʱ��=now() where ����='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}���Լ����õ�"&lx&":["&wpname&"]"&wpsl&"��ԭֵ��Ա��"&bb&"Ԫ�õ����̵��˻�Ա��"&yin&"Ԫ����У�����ֵ��������<font color=#ff00ff>("&time&")</font>"
		end if
	end if
	
case "ҩƷ"
	wpsl=dansl
	zstemp=abate(rs("w1"),wpname,dansl)
	if wpsl>0 then
		rs.close
		rs.open "select h from b where a='"&wpname&"'",conn
		if rs.bof and rs.eof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('û�����뵱��������Ʒ��������ʲô�����վ����ϵ��');location.href = 'javascript:history.go(-1)';}</script>"
			response.end
		else
			bb=rs("h")*wpsl
			yin=int(bb/3)
			conn.execute "update �û� set w1='"&zstemp&"',����=����+" & yin & ",����ʱ��=now() where ����='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}���Լ����õ�"&lx&":["&wpname&"]"&wpsl&"��ԭֵ����"&bb&"���õ����̵���"&yin&"�������ǿ���Ӵ������<font color=#ff00ff>("&time&")</font>"
		end if
	end if
	
case "��ҩ"
	wpsl=dansl
	zstemp=abate(rs("w2"),wpname,dansl)
	if wpsl>0 then
		rs.close
		rs.open "select h from b where a='"&wpname&"'",conn
		if rs.bof and rs.eof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('û�����뵱��������Ʒ��������ʲô�����վ����ϵ��');location.href = 'javascript:history.go(-1)';}</script>"
			response.end
		else
			bb=rs("h")*wpsl
			yin=int(bb/3)
			conn.execute "update �û� set w2='"&zstemp&"',����=����+" & yin & ",����ʱ��=now() where ����='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}���Լ����õ�"&lx&":["&wpname&"]"&wpsl&"��ԭֵ����"&bb&"���õ����̵���"&yin&"�������ǿ���Ӵ������<font color=#ff00ff>("&time&")</font>"
		end if
	end if

case "����"
	wpsl=dansl
	zstemp=abate(rs("w4"),wpname,dansl)
	if wpsl>0 then
		rs.close
		rs.open "select h from b where a='"&wpname&"'",conn
		if rs.bof and rs.eof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('û�����뵱��������Ʒ��������ʲô�����վ����ϵ��');location.href = 'javascript:history.go(-1)';}</script>"
			response.end
		else
			bb=rs("h")*wpsl
			yin=int(bb/3)
			conn.execute "update �û� set w4='"&zstemp&"',����=����+" & yin & ",����ʱ��=now() where ����='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}���Լ����õ�"&lx&":["&wpname&"]"&wpsl&"��ԭֵ����"&bb&"���õ����̵���"&yin&"�������ǿ���Ӵ������<font color=#ff00ff>("&time&")</font>"
		end if
	end if

case "��Ʒ"
	wpsl=dansl
	zstemp=abate(rs("w6"),wpname,dansl)
	if wpsl>0 then
		rs.close
		rs.open "select h from b where a='"&wpname&"'",conn
		if rs.bof and rs.eof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('û�����뵱��������Ʒ��������ʲô�����վ����ϵ��');location.href = 'javascript:history.go(-1)';}</script>"
			response.end
		else
			bb=rs("h")*wpsl
			yin=int(bb/3)
			conn.execute "update �û� set w6='"&zstemp&"',����=����+" & yin & ",����ʱ��=now() where ����='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}���Լ����õ�"&lx&":["&wpname&"]"&wpsl&"��ԭֵ����"&bb&"���õ����̵���"&yin&"�������ǿ���Ӵ������<font color=#ff00ff>("&time&")</font>"
		end if
	end if

case "�ʻ�"
	wpsl=dansl
	zstemp=abate(rs("w7"),wpname,dansl)
	if wpsl>0 then
		rs.close
		rs.open "select h from b where a='"&wpname&"'",conn
		if rs.bof and rs.eof then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('û�����뵱��������Ʒ��������ʲô�����վ����ϵ��');location.href = 'javascript:history.go(-1)';}</script>"
			response.end
		else
			bb=rs("h")*wpsl
			yin=int(bb/3)
			conn.execute "update �û� set w7='"&zstemp&"',����=����+" & yin & ",����ʱ��=now() where ����='"&aqjh_name&"'"
			dan="{"&aqjh_name&"}���Լ����õ�"&lx&":["&wpname&"]"&wpsl&"��ԭֵ����"&bb&"���õ����̵���"&yin&"�������ǿ���Ӵ������<font color=#ff00ff>("&time&")</font>"
		end if
	end if

end select

rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>��������Ϣ��</b></font>"&dan
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
nowinroom=session("nowinroom")
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
Response.Write "<script language=JavaScript>{alert('���Լ����õ�"&lx&":["&wpname&"]"&wpsl&"��ԭֵ"&bb&"�õ����̵���"&yin&"������Ӵ��������');location.href = 'dan.asp';}</script>"
'Response.Redirect ""
%>