<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
Response.Write "<script language=JavaScript>{top.location.href='../error.asp?id=440';}</script>"
Response.End
%>