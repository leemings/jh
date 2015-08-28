<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("yx8_mhjh_id")
sjjh_name=session("yx8_mhjh_username")
sjjh_grade=session("yx8_mhjh_usergrade")
sjjh_jhdj=session("yx8_mhjh_usergrade")
'if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
if session("yx8_mhjh_inchat")="" then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
rnds=Request.Form("rnds")
if int(rnds)=session("rnds") then
score=Request.Form("sc")
score=score*50
if score>99999 then score=99999
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="update 用户 set 银两=银两+'"&score&"',体力=体力-50000 where 姓名='"&sjjh_name&"' and 体力>50000"
conn.Execute (sql)
conn.Close 
set rs=nothing
end if
%>