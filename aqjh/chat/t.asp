<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
useronlinename=Application("aqjh_useronlinename"&nowinroom)
if Instr(LCase(useronlinename),LCase(" "&Session("aqjh_name")&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
webicqname=Application("aqjh_webicqname")
For i = 1 to Application("SayCount")-Session("SayCount")
Response.Write Application("SayStr"&YuShu((Session("SayCount")+i)))
Next
session("SayCount")=Application("SayCount")
Function Yushu(a)
Yushu=(a and 31)
End Function
Response.Write "<script Language=JavaScript>"
Response.Write "setTimeout('this.location.reload();',6000);"
if Instr(webicqname," "&aqjh_name&" ")>0 then
Application("aqjh_webicqname")=replace(Application("aqjh_webicqname"),aqjh_name,"")
Response.Write "window.open('webicqread.asp','','Status=no,scrollbars=yes,resizable=no,width=430,height=160');"
end if
Response.Write "</script>"
%>