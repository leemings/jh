<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
horsecalc=Request.QueryString("calc")
horsewin=Request.QueryString("win")
bet=Request.QueryString("bet")
msg="<head><link rel=stylesheet href='css.css'></head><body bgcolor='"&bgcolor&"' background='"&bgimage&"'><table width=50% align=center><tr><td>共下注："&horsecalc&"</td><td>获胜："&horsewin&"号马</td><td>下注："&bet&"</td><td>"
win=horsecalc-bet*10
if win>0 then
msg=msg&"输："&win
else
msg=msg&"赢："&abs(win)
end if	
msg=msg&"</td></tr><tr><td colspan=6 align=center><input type=button onclick=javascript:location.replace('chipin.asp'); value=' 返 回 '></td></tr></table></body>"
Response.Write msg
%>
