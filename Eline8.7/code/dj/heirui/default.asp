<%  
IF Session("auth") = True THEN 
Response.ContentType = "application/x-javascript"  
Response.Expires = 0  
Response.Expiresabsolute = Now() - 1  
Response.AddHeader "pragma","no-cache"  
Response.AddHeader "cache-control","private"  
Response.CacheControl = "no-cache"  
%>
document.player.DoStop();
document.player.DoPlay();
document.player.SetSource("readrm.asp?songid="+songid)
<%ELSE%><!--��Щ�����ܰ�Ȩ����������Ȩ������--><%END IF%>