<!--#include file="../config.asp"-->
<%Response.Buffer=true
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if Application("yx8_mhjh_advertisemen")="" then
   advertisemen=yx8_mhjh_advertisemen
else
   advertisemen=Application("yx8_mhjh_advertisemen")
end if
advertisemens=Split(Trim(advertisemen),";",-1)
total=UBound(advertisemens)
randomize timer
x=int(rnd*(total+1))
%>
<html>
<head>
<script language=javascript>setTimeout('this.location.reload();',90000);</script>
<link rel="stylesheet" href="css.css">
<title></title>
</head>
<body topmargin="0" background="bg1.gif" style="CURSOR: url('../image/banana.ani')" oncontextmenu='self.event.returnValue=false'>
<p align="center"><b><font color="#FF0000"><%=advertisemens(x)%></font></b></p>
</body>
</html>
