<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if session("myroom")<>"h_ӵ����" then
	Response.Write "<script Language=Javascript>alert('��ʾ����Ǵ˷���������Ȩ����');parent.email.style.visibility='hidden';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM house WHERE h_ӵ����='" & aqjh_name &"'",conn,1,3
if rs.Eof or rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������û�й����ӣ���������');top.close();</script>"
	Response.End
end if
myroomname=trim(rs("h_С����"))
roomlb=rs("h_����")
roomjb=rs("h_����")
naijd=rs("h_�;ö�")
if naijd>=90 then
	mshow="�����������!"
elseif naijd>=80 then
	mshow="���˵���ȫû����!"
elseif naijd>=70 then
	mshow="������һ���𻵣�"
elseif naijd>=50 then
	mshow="���Ӷ�����,��©��!"
elseif naijd>=40 then
	mshow="�����ķ��ӻ�������ס?"
elseif naijd>=30 then
	mshow="�����ıߵ�ǽû��ʲô��!"
elseif naijd>=20 then
	mshow="�쵱���ص���,���һ����ɽǽ!"
else
	mshow="���ﲻ����ס����!"
end if
rs.close
act=lcase(trim(Request("act")))
if act="xiuli" then
	Submit=trim(Request.form("Submit"))
	if Submit="ȷ������" then
		rs.open "SELECT ���� FROM [�û�] WHERE ����='" & aqjh_name &"'",conn,1,3
		mf=(1000+((roomjb-1)*500))
		if rs("����")<mf then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('��ʾ���㶹���������["& mf &"]�㲻������');location.href = 'xiuli.asp';</script>"
			Response.End
		end if
		rs("����")=rs("����")-mf
		rs.update
		rs.close
		rs.open "SELECT * FROM house WHERE h_ӵ����='" & aqjh_name &"'",conn,1,3
		rs("h_�;ö�")=rs("h_�;ö�")+50
		rs.update
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ�����Ѷ���"& mf &"���;ö�����50�㣡');location.href = 'xiuli.asp';</script>"
		Response.End
	end if
end if
set rs=nothing
conn.close
set conn=nothing
%>
<LINK href="style.css" rel=stylesheet>
<script language="JavaScript" type="text/JavaScript">
function show(x,y,sm){Layer1.style.visibility="visible";mess.innerHTML="<img src=pic/hut"+x+y+".jpg width=185 height=125><br><div align=center>"+sm+y+"��</div>";}
function hidden(){Layer1.style.visibility="hidden";}
document.onmousedown=click;
function click(){if(event.button==2){alert("��ӭ�������߿��ֽ�����");}}
function shLayers(n,n1){
if(n.style.visibility=="visible"){n.style.visibility="hidden";}else if(n.style.visibility=="hidden"){n.style.visibility="visible";}
if(n1.style.visibility=="visible"){n1.style.visibility="hidden";}
}
</script>
<title>������</title>
<body background="IMAGES/bg.jpg" text="#FFFFFF" leftmargin="0" topmargin="0">
<table width="423" height="324" border="0" cellpadding="0" cellspacing="0">
    <td width="423" height="324" align="left" valign="top"> 
        <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><font color="#FFFF00"><strong>�� �� �� ��</strong></font><br>
              ����һ��ʱ���ʹ��,���С�ݾ����������Ҳ����,���Ծ�Ҫ����������������������ҿ��԰����ݵȼ��շѵ�,���;�ֵС��<font color="#FFFF00"><strong>50</strong></font>��ʱ���������޸�ÿ���޸������;ö�<font color="#FFFF00"><strong>50</strong></font>��,ÿ���շ�<font color="#FFFF00"><strong>1000</strong></font>���������,����ÿ���1������������<font color="#FFFF00"><strong>50%</strong></font>,��˵�Һ�,����ͷΪ�������ұ˴˱˴�....<br>
              <br>
              ״̬:<%=naijd%>��<br>
              <br>
              ����:<%=mshow%><br>
              <br>
              <br>
<%if naijd<=50 then%>
<form name="form" method="post" action="xiuli.asp?act=xiuli" >
<input type="submit" name="Submit" value="ȷ������"> </td>
</form><%end if%>
</tr>
</table>
        <br>
        <br>

</td>
</tr>
</table>
<div id="Layer0" style="position:absolute; width:200px; height:115px; z-index:3; left: 195px; top: 76px;"><img src="pic/xiuli.gif" width="223" height="247"></div>












