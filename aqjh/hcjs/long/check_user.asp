<%
if session("aqjh_name")="" or session("aqjh_id")="" then 
Response.Write "<script Language=Javascript>alert('�Բ���,����δ��¼�����Ѿ���ʱ�Ͽ������ӣ������µ�¼��');window.close();</script>"
 response.end
end if
myid=session("aqjh_id")

%>
</body>
</html>
