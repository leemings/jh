<%
Response.Expires=0 
sjq=DateDiff("s",Application("yx8_mhjh_dg"),now())  
Application.Lock 
Application("yx8_mhjh_dg")=now() 
Application.UnLock  
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016" 
nowdate=cstr(date()) 
bkmn=Request.Form("money") 
bkop=Request.Form("op") 
if isnumeric(bkmn) then 
if bkmn>0 then 
if bkop="ȡ��" then sqlstr="update �û� set ����=����+"&bkmn&",���=���-"&bkmn&" where ����='"&username&"' and ���>="&bkmn 
if bkop="���" then sqlstr="update �û� set ����=����-"&bkmn&",���=���+"&bkmn&" where ����='"&username&"' and ����>="&bkmn 
if sqlstr="" then 
msg="���룿�������룿���Բ��𣿣�" 
else 
set conn=server.CreateObject("adodb.connection") 
conn.Mode=16 
conn.IsolationLevel=256 
conn.Open Application("yx8_mhjh_connstr")
conn.Execute(sqlstr) 
conn.Close 
set conn=nothing 
response.redirect "ok.asp?id=100" 
end if 
else 
response.redirect "../error.asp?id=521" 
end if 
else 
response.redirect "../error.asp?id=522" 
end if 
%> 
