<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
kfnum=int(Request.Form("kfnum"))
kfnum=abs(kfnum)
hygrade=int(Request.form("hygrade"))

Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
Set rs=Server.CreateObject("ADODB.RecordSet")
if kfnum<1 or kfnum>1000 then
	Response.Write "<script language=javascript>alert('��Ǹ��һ�μӿ��ѱ�����1~1000֮�䣡��鿴�Ƿ���ȷ��');history.back();</script>"
	response.end
else
   sql="update �û� set ��Ա��=��Ա��+"&kfnum&" where ��Ա�ȼ�="& hygrade 
   Set rs=conn.Execute(sql)
end if

set rs=nothing
conn.Close
set conn=nothing
says="<font size=2 color=red>���𿨷��š�</font>��ϲվ������"&hygrade&"����Ա��"&kfnum&"Ԫ������ɣ�"
act="��Ϣ"
towhoway=1
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
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
Response.Write "<script language=javascript>alert('��ϲ��["&hygrade&"]����Ա�Ļ�Ա��"&kfnum&"������ɣ�');history.back();</script>"
%>