<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
id=trim(Request.QueryString("id"))
sid=trim(Request.QueryString("sid"))
if id="" or sid="" then sid=801
%>
<HTML><HEAD><TITLE>江湖秀</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK href="show.css" rel=stylesheet>
<META content="MSHTML 6.00.2600.0" name=GENERATOR></HEAD>
<FRAMESET border=0 frameSpacing=0 rows=107,*,0 frameBorder=NO cols=*>
<FRAME name=topFrame src="ttop.asp?id=<%=id%>" frameBorder=NO noResize scrolling=no>
<FRAME name=mainFrame src="main.asp?sid=<%=sid%>" frameBorder=NO noResize>
<FRAME name=myjhshow marginWidth=0 marginHeight=0  frameBorder=NO noResize>
</FRAMESET>
</HTML>
