<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ��ż,�Ա� from �û� where ����='"&aqjh_name&"'",conn,3,3
mysex=rs("�Ա�")
if isnull(rs("��ż")) or rs("��ż")="" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������û�У����Լ�һ������ѽ��');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
rs.open "select �������� from �û� where ����='"&aqjh_name&"'",conn,3,3
if DateDiff("d",rs("��������"),date())<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����Ŀ��ֻ������ȶ�����2���ɣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
rs.open "select boy,boysex from �û� where ����='"&aqjh_name&"'",conn,3,3
if isnull(rs("boy")) or rs("boy")="" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�������������У���û����С����');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
mypic=rs("boysex")

zt=split(rs("boy"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��С�����ݳ���������������');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if clng(zt(3))<=0 then
	conn.execute "update �û� set boy='' where  ����='"&aqjh_name&"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����Ϊ�㲻�ú��չ�С��,С�������ˣ���');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
cwsm=clng(DateDiff("d",zt(7),now()))
if cwsm>0 then
	temp=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))-cwsm*15)&"|"&zt(4)&"|"&zt(5)&"|"&zt(6)&"|"&zt(7)
	conn.execute "update �û� set boy='"&temp&"' where  ����='"&aqjh_name&"'"
		if (clng(zt(3))-cwsm*15)<=0 then
			conn.execute "update �û� set boy='' where  ����='"&aqjh_name&"'"
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('��ʾ����Ϊ�㲻�ú��չ�С��,С�������ˣ���');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
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
<script Language="Javascript">
function del(){if(confirm("�����Ҫ�������С����")){window.location.href="delboy.asp";return true;}}
</script>
<html>
<head>
<title>�ҵ�С��</title>
<script language="JavaScript">
function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}</script>
<style>
body{
CURSOR: url('boy.cur');
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff}
td{font-size:9pt;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgproperties="fixed" leftmargin="0" topmargin="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" background="../<%=chatimage%>">
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
    <td width="100%" height="28"> <div align="center"><font color="#ccccff" size="2"><b>�ҵ�С��</b></font></div></td>
  </tr>
</table>
<table background="../../bg2.gif" border="1" cellspacing="0" width="140" bordercolorlight="#000000"
bordercolordark="#FFFFFF" cellpadding="4" align="center">
  <tr> 
    <td>��Ƭ:<img src="<%=mypic%>" width=40 height=40><br><br>
      ����:<%=zt(0)%><br><br>
      �Ա�:<%=zt(1)%><br><br>
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
    <td><div align="center"><a href="yang.asp" title="��С���Ե��׷��ɣ�����������"><font color="#ffffff">��</font></a> <a href="dlboy.asp" title="����С�����ӹ���!" target="d"><font color="ffffff">��</font></a> 
        <a href="pbboy.asp" title="���С�������Ӹ���!" target="d"><font color="ffffff">��</font></a> <a href="gmboy.asp" title="ȡ�������ְ�!" ><font color="ffffff">��</font></a> <a href="javascript:s('/С��$ <%=zt(0)%>')" ><font color="ffffff">��</font></a> <a href="javascript:s('/С������$ <%=zt(0)%>')" title="С������">��</a></div></td>
  </tr>
<%if mysex="Ů" then%>
<tr><td align=center><a href=# onclick="javascript:del()">����С��</a></td></tr>
<%end if%>
</table>
</body></html>