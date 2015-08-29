<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
useronlinename=Application("sjjh_useronlinename"&nowinroom)
if Instr(LCase(useronlinename),LCase(" "&Session("sjjh_name")&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 

For i = 1 to Application("SayCount")-Session("SayCount")
Response.Write Application("SayStr"&YuShu((Session("SayCount")+i)))
Next
session("SayCount")=Application("SayCount")
Function Yushu(a)
Yushu=(a and 31)
End Function
Response.Write "<script Language=JavaScript>"
Response.Write "setTimeout('this.location.reload();',3000);"
Response.Write "</script>"
%>