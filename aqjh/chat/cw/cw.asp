<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select cw from �û� where ����='"&aqjh_name&"'",conn,3,3
if isnull(rs("cw")) or rs("cw")="" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������û�г����ȥ����');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
zt=split(rs("cw"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������ݳ����������¹���');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if clng(zt(3))<=0 then
	conn.execute "update �û� set cw='' where  ����='"&aqjh_name&"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����Ϊ�㲻�ú��չ˳������ˣ���');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
rs.open "select i from b where a='"&zt(1)&"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������ݳ����������¹���');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
mypic=rs("i")
cwsm=clng(DateDiff("d",zt(7),now()))
if cwsm>0 then
	temp=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))-cwsm*15)&"|"&zt(4)&"|"&zt(5)&"|"&zt(6)&"|"&zt(7)
	conn.execute "update �û� set cw='"&temp&"' where  ����='"&aqjh_name&"'"
		if (clng(zt(3))-cwsm*15)<=0 then
			conn.execute "update �û� set cw='' where  ����='"&aqjh_name&"'"
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('��ʾ����Ϊ�㲻�ú��չ˳������ˣ���');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
			response.end
		end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
if DateDiff("h",zt(7),now())<5 then
	zg=clng(zt(6))
else
	zg=0
end if
zgg=(4+int(DateDiff("d",zt(2),now())/10))
cssj=clng(DateDiff("d",zt(2),now()))
gs=clng(zt(5))
if gs<(80+cssj*5) then
	czzt="����"
elseif gs>((100+cssj*5+50)) then
	czzt="˳��"
else
	czzt="��ͨ"
end if
%>
<html>
<head>
<title>�ҵĳ���</title>
<script language="JavaScript">
function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}</script>
<style>
td{font-size:9pt;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../<%=chatimage%>" bgproperties="fixed" leftmargin="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<p style="font-size:12pt; color:red" align="center"> <font color="#FFFFFF"><b>�� �� �� ��</b><br>
  </font> <img src="../../hcjs/jhjs/images/<%=mypic%>"></p>
<table width="80%" border="1" align="center" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="1" cellspacing="0" >
  <tr> 
    <td>��Ƭ:<img src="../../hcjs/jhjs/images/<%=mypic%>"><br><br>
      ����:<%=zt(0)%><br><br>
      ���:<%=zt(1)%><br><br>
      ����:<%=(clng(zt(3))-cwsm*15)%><br><br>
      ����:<%=zt(4)%><br><br>
      ����:<%=zt(5)%><br><br>
      ״̬:<%=czzt%><br><br>
      �չ�:<%=zg%>/<font color=red><b><%=zgg%></b></font><br><br>
      ����:<%=Year(zt(2))%>-<%=Month(zt(2))%>-<%=day(zt(2))%><br><br>
      ����:<%=DateDiff("d",zt(2),now())%>��<br><br>
    </td>
  </tr>
  <tr>
    <td><div align="center"><a href="yang.asp" title="ʹ���ʻ���ι���������������">��</a> <a href="dlcw.asp" title="�����������ӹ���!" target="d">��</a> 
        <a href="pbcw.asp" title="��������Ӹ���!" target="d">��</a> <a href="gmcw.asp" title="ȡ�������ְ�!" >��</a> <a href="javascript:s('/����$ <%=zt(1)%>')" >��</a> <a href="javascript:s('/�����Ա�$ <%=zt(1)%>')" >��</a></div></td>
  </tr>
</table>
</body>
</html>