<%@ LANGUAGE=VBScript codepage ="936" %>
<%

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<!--#include file="h_gamble.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
dim bet
dim over
dim pages
pages="gamble.asp"
bet=trim(Request.Form("Bet"))
select case Request.Form("Action")
case "����"
over=0
call GameStart(bet)
case "����һ��"
over=0
call GameStart(bet)
case "Ҫ��"
over=0
call GiveUserPoker()
if CanAddPoker() then
GiveAntiPoker()
end if
case "����"
if CanAddPoker() then
over=0
GiveAntiPoker()
else
''''''''''''''''''''''''
Session("n_Result")=Result()	
Session("n_Begin")=false
			
dim value	
over=0	
value=CLng(Session("n_Bet"))
Session("n_TotalMoney")=Session("n_TotalMoney")+value
'value=CCur(Session("n_Bet"))
select case CInt(Session("n_Result"))
case 1 'win
				
value=value
case 2 'lost
value=0-value
case 3
value=0
case else
value=0
end select
conn.execute("Update �û� set ����=����+"&value&",�Ĵ���=�Ĵ���+1,����=����-1,ӮǮ=ӮǮ+"&value&" where ����='"&sjjh_name&"'") 				
Session("n_TotalMoney")=Session("n_TotalMoney")+value		
Session("n_Bet")=0

end if
case "������"
pages="../jh.asp"
case "����"
'�������ʼ������
'	Session("n_Init")=false
'	Session("n_TotalMoney")=1000
'case "������"
'�˳��ɡ�
pages="../jh.asp"
case else
call GameStart(bet)
end select


if over=0 then
Response.Redirect pages
end if
	
%>
<html><script language="JavaScript">                                                                  </script></html>
<html></html>