<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�Զ�������wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(9)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻�ܹ��Ĳ���');}</script>"
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
says=duidu(mid(says,i+1),towho)+"</font>"
call chatsay("�Զ�����",towhoway,towho,saycolor,addwordcolor,addsays,says)

'�Զ�����
'���˽����״�
function duidu(fn1,to1)
if not isnumeric(fn1) then
	Response.Write "<script language=JavaScript>{alert('["&fn1&"]��������������ʹ�����֣�');}</script>"
	Response.End 
end if
fn1=int(abs(fn1))
if fn1<10000 or fn1>3000000 then
	Response.Write "<script language=JavaScript>{alert('��ʾ���Զ�����Ӧ����10��С��300��');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ���� from �û� where ����='"&sjjh_name&"'",conn,2,2

if rs("����")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������������󲻹�ѽ����ô���˼Ҷ�??');}</script>"
	Response.End
end if
rs.close
rs.open "select ���� from �û� where ����='"&to1&"'",conn,2,2

if rs("����")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&to1&"]�����û����ô������Ӱ�?��ô��???!');}</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
randomize()
m1 = Int(6 * Rnd + 1)
m2 = Int(6 * Rnd + 1)
m3 = Int(6 * Rnd + 1)
regjm=int(rnd*3348998)+1
'duidu="[##]��{%%}����Զ����룬��עΪ��"&fn1&"���ӣ�����,��������һ�� ,<input  type=button value='˭��˭,����!' 'onClick=javascript:duidu"&regjm&".disabled=1;window.open('duidu_add.asp?name="&sjjh_name&"&fn1="&fn1&"&to1="&to1&"&m1="&m1&"&m2="&m2&"&m3="&m3&"','d') name=duidu"&regjm&">"
'duidu="<form method=POST target=d name=duidu"&regjm&" 'action=duidu_add.asp?name="&sjjh_name&"&fn1="&fn1&"&toname="&to1&"&m1="&m1&"&m2="&m2&"&m3="&m3&"><font color=green>���Զ�������<font color=" & saycolor & ">[##]��{%%}����Զ����룬��עΪ��"&fn1&"���ӣ�����,��������һ��!<input type=submit onClick=javascript:duidu"&regjm&".disabled=1;this.document.duidu"&regjm&".submit() name=duidu"&regjm&" value=˭��˭,����!></form>"

duidu="<form method=POST target=d name=duidu"&regjm&" action=duidu_add.asp?name="&urlencoding(sjjh_name)&"&fn1="&fn1&"&toname="&urlencoding(to1)&"&m1="&m1&"&m2="&m2&"&m3="&m3&"><font color=green>���Զ�������<font color=" & saycolor & ">[##]��{%%}����Զ������룬��עΪ��"&fn1&"���ӣ�����,��������һ��!<input type=submit onClick=javascript:duidu"&regjm&".disabled=1;this.document.duidu"&regjm&".submit() name=duidu"&regjm&" value=˭��˭,����!></form>"
end function

function urlencoding(vstrin)
    strreturn = ""
    for i = 1 to len(vstrin)
        thischr = mid(vstrin,i,1)
        if abs(asc(thischr)) < &hff then
            strreturn = strreturn & thischr
        else
            innercode = asc(thischr)
            if innercode < 0 then
                innercode = innercode + &h10000
            end if
            hight8 = (innercode  and &hff00)\ &hff
            low8 = innercode and &hff
            strreturn = strreturn & "%" & hex(hight8) &  "%" & hex(low8)
        end if
    next
    urlencoding = strreturn
end function

%>

