<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
Response.Write "<head><title>归来吧</title><script language=javascript>window.resizeTo(200,20);window.moveTo(screen.availWidth/3,screen.availHeight/4);</script></head><body oncontextmenu='self.event.returnValue=false' background='bg1.gif'><marquee><font color=FF0099 size=5>归来吧，归来哟，浪迹天涯的游子！再不归来江湖就要沉下大海去了。</font></marquee></body>"
%>
