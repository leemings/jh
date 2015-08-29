<%Session.Timeout = 120%>
document.write('<font color="#FF0000"><%Response.Write CInt(Application("Counter"))%></font>')