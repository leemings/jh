<%@ LANGUAGE=VBScript codepage ="936" %>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
grade=Trim(Request.Form("grade"))
Response.Redirect "manaccquerygradeview.asp?grade=" & server.URLEncode(grade) & "&page=1"%>