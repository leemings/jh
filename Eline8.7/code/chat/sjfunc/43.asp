<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc.asp"-->
<%'查看本门基金♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
says="<font color=red>【基金查看】<font color=" & saycolor & ">"+seejj()+"</font>"
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'查看本门基金seejj
function seejj()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select 门派基金,门派 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn,2,2
mp=rs("门派")
if mp="游侠" or mp="未知" or mp="无" or mp="" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你还没有门派，不能查看门派基金！');}</script>"
	Response.End
end if
myjj=rs("门派基金")
rs.close
rs.open "select h FROM p WHERE a='" & mp & "'",conn,2,2
bpjj=rs("h")
seejj="##你现的对本门的贡献：<font color=red><b>"&myjj&"</b></font>两,["&mp&"]的基金数为：<font color=red><b>"&bpjj&"两！</b></font>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>