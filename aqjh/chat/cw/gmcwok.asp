<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
oldname=trim(Request.Form ("oldname"))
newname=trim(Request.Form ("newname"))
if newname="" or oldname="" or oldname=newname then
	Response.Write "<script language=JavaScript>{alert('��ʾ������������Ϊ�գ�Ҳ������ԭ����ͬ!');}</script>"
	Response.End 
end if
namelen=0
for i=1 to len(newname)
zh=mid(newname,i,1)
zhasc=asc(zh)
if zhasc<0 then
	namelen=namelen+2
else
	namelen=namelen+1
	if CStr(server.URLEncode(zh))<>CStr(zh) then Response.Redirect "../../error.asp?id=120"
end if
next
if namelen>10 then 
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������Ϊ5�����֣�');}</script>"
	Response.end
end if
badword="վ��,����,����,�侫,��,��,ʺ,��,��,��,��,����,��,��,��,�,��,��,��,��,��,����,غ��,��,��Ƥ,��ͷ,��,�P,��,�H,��,��,��,����,����,��,����,����,��Ů,����,ʮ����,��,�Ҷ�,��,��ϯ,����,����,��־,������,�ٸ�"
bad=split(badword,",")
for i=0 to ubound(bad)-1
	if InStr(LCase(newname),bad(i))<>0 then 
		Response.Write "<script language=JavaScript>{alert('��ʾ�����������ţ��������룡');}</script>"
		Response.end
	end if
next
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select cw from �û� where ����='"&aqjh_name&"'",conn,2,2
if isnull(rs("cw")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������û�г����ȥ����');</script>"
	response.end
end if
zt=split(rs("cw"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������ݳ����������¹���');</script>"
	response.end
end if
rs.close
rs.open "select i from b where a='"&zt(1)&"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������ݳ����������¹���');</script>"
	response.end
end if

'����|���|����|����|����|����|�չ�����|�չ�ʱ��
'����|����|2002-6-28 11:04:29|100|100|100|0|2002-6-28 11:04:29
temp=newname&"|"&zt(1)&"|"&zt(2)&"|"&zt(3)&"|"&zt(4)&"|"&zt(5)&"|"&zt(6)&"|"&zt(7)
conn.execute "update �û� set cw='"&temp&"' where  ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000><b>�����������</b></font>##�Ѹ��Լ��ĳ������Ϊ[<font color=red><b>"&newname&"</b></font>]�ɹ�����"
zj="<a href=javascript:parent.sw(\'[" & aqjh_name & "]\'); target=f2>"& aqjh_name & "</a>"
'br="<a href=javascript:parent.sw(\'[" & towho & "]\'); target=f2>" & towho & "</a>"
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
says=Replace(says,"##",zj,1,3,1)
'says=Replace(says,"%%",br,1,3,1)
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
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
Response.Write "<script Language=Javascript>alert('��ʾ�����������ĳ�["&newname&"]�ɹ�!');parent.f3.location.href = 'cw.asp';</script>"
%>