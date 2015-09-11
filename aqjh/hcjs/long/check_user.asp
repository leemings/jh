<%
if session("aqjh_name")="" or session("aqjh_id")="" then 
Response.Write "<script Language=Javascript>alert('对不起,你尚未登录，或已经超时断开了连接，请重新登录！');window.close();</script>"
 response.end
end if
myid=session("aqjh_id")

%>
</body>
</html>
