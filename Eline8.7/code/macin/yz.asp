<%
Randomize
session("rnds")=int(rnd*99999)
Response.Write "&rnds="&server.URLEncode(session("rnds"))

%>