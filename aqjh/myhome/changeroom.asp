<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if session("myroom")<>"h_ӵ����" then
	Response.Write "<script Language=Javascript>alert('��ʾ����Ǵ˷���������Ȩ������');parent.email.style.visibility='hidden';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM house WHERE h_ӵ����='" & aqjh_name &"'",conn,1,1
if rs.Eof or rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������û�й����ӣ�����������');top.close();</script>"
	Response.End
end if
myroomname=trim(rs("h_С����"))
roomlb=rs("h_����")
roomjb=rs("h_����")
rs.close
rs.open "Select * from [housetype] where ht_���="&roomlb,conn,1,1
if rs.Eof or rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����ķ�������������վ����ϵ��');top.close();</script>"
	Response.End
end if
tj2=replace(replace(rs("ht_2������"),";","&nbsp;"),"|",":")
tj3=replace(replace(rs("ht_3������"),";","&nbsp;"),"|",":")
tj4=replace(replace(rs("ht_4������"),";","&nbsp;"),"|",":")
rs.close
act=lcase(trim(Request("act")))
if act="changeroom" then
	Response.Write "<html><body background=IMAGES/bg.jpg><div align=center><a href=javascript:history.go(-1);><font color=#0000FF>�����ﷵ��</font></a></div>"
	roomid=clng(Request.form("roomid"))
	if roomid<0 or roomid>5 then
		Response.Write "<script Language=Javascript>alert('��ʾ��������𲻶ԣ����Ϊ0-5��');location.href = 'changeroom.asp';</script>"
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End
	end if
	if roomid=roomlb then
		Response.Write "<script Language=Javascript>alert('��ʾ����ѡ��������������ڵķ�������һ�������ܲ�����');location.href = 'changeroom.asp';</script>"
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End
	end if 
	rs.open "SELECT * FROM housetype WHERE ht_���=" & roomid,conn,1,1
	xlwp=rs("ht_1������")
	if isnull(xlwp) or xlwp="" or instr(xlwp,"|")=0 or instr(xlwp,";")=0 then 
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ�����������д��ܲ���\n���ҳ��򿪷�����ϵ!');</script>"
		response.end
	end if
	xadata=split(xlwp,";")
	xadata1=UBound(xadata)
	rs.close
	rs.open "SELECT w6,�ȼ�,��Ա�ȼ�,����,��� from [�û�] WHERE ����='" & aqjh_name & "'",conn,1,3
	duyao=rs("w6")
	if isnull(duyao) or duyao="" then 
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ����ʲô��ƷҲû��!');location.href = 'changeroom.asp';</script>"
		response.end
	end if
	yin=0
	for ii=0 to xadata1
		xadata2=split(xadata(ii),"|")
		mysl=clng(xadata2(1))
		myxlwp=trim(xadata2(0))
		select case myxlwp
		case "�ȼ�"
			if rs("�ȼ�")<mysl then
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('��ʾ�����ĵȼ�δ��["& mysl &"]�������ܸ�������!');location.href = 'changeroom.asp';</script>"
				response.end
			end if
		case "���"
			if rs("���")<mysl then
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('��ʾ�����Ľ��û��["& mysl &"]�������ܸ�������!');location.href = 'changeroom.asp';</script>"
				response.end
			end if
			rs("���")=rs("���")-mysl
		case "����"
			yin=yin+mysl
			if rs("����")<mysl then
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('��ʾ����������û��["& mysl &"]�������ܸ�������!');location.href = 'changeroom.asp';</script>"
				response.end
			end if
			rs("����")=rs("����")-mysl
		case "��Ա�ȼ�"
			if rs("��Ա�ȼ�")<mysl then
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('��ʾ�����Ļ�Ա�ȼ�δ��["& mysl &"]�������ܸ�������!');location.href = 'changeroom.asp';</script>"
				response.end
			end if
		case else
			if mywpsl(duyao,myxlwp)<mysl then
    			Set rs = Nothing
   				conn.Close
  				Set conn = Nothing
   				Response.Write "<script Language=Javascript>alert('��ʾ����Ʒ"& myxlwp &"�������㣿');</script>"
				response.end
			end if
			duyao=abate(duyao,myxlwp,mysl)
		end select
	next
	rs("w6")=duyao
	rs.update
	rs.close
	rs.open "SELECT * FROM housetype WHERE ht_���=" & roomid,conn,1,3
	h_name=rs("ht_С����")
	h_nai=rs("ht_1���;ö�")
	rs("ht_С���ʲ�")=rs("ht_С���ʲ�")+int(yin/10000)
	rs.update
	rs.close
	rs.open "SELECT * FROM house WHERE h_ӵ����='" & aqjh_name &"'",conn,1,3
	rs("h_С����")=h_name
	rs("h_����")=1
	rs("h_�;ö�")=h_nai
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=""Javascript"">alert(""��ϲ�㣬���ݸ�����ɣ�"");location.href = 'changeroom.asp';;</script>"
end if
%>

<head>
<LINK href="style.css" rel=stylesheet>
<script language="JavaScript" type="text/JavaScript">
function show(x,y,sm){
if(x==0){Layer1.style.top="182px";}
else if(x==1){Layer1.style.top="182px";}
else if(x==2){Layer1.style.top="182px";}
else if(x==3){Layer1.style.top="1px";}
else if(x==4){Layer1.style.top="1px";}
else if(x==4){Layer1.style.top="1px";}
Layer1.style.visibility="visible";mess.innerHTML="<img src=pic/hut"+x+y+".jpg width=185 height=125><br><div align=center>"+sm+y+"��</div>";
}
function hidden(){Layer1.style.visibility="hidden";}
document.onmousedown=click;
function click(){if(event.button==2){alert("��ӭ�������߰��齭����");}}
function shLayers(n,n1){
if(n.style.visibility=="visible"){n.style.visibility="hidden";}else if(n.style.visibility=="hidden"){n.style.visibility="visible";}
if(n1.style.visibility=="visible"){n1.style.visibility="hidden";}
}
function check(){
if(document.form.roomid[0].checked==false && document.form.roomid[1].checked==false && document.form.roomid[2].checked==false && document.form.roomid[3].checked==false && document.form.roomid[4].checked==false && document.form.roomid[5].checked==false)
{alert("��ʾ����ѡ����Ҫ���ķ������ͣ�");return false;}
if(confirm("���ڽ�Ҫ���л����ݲ�������ԭ���ķ��ݽ���\n��ϵͳ���գ������޷��ָ�����ȷ�������·���")){return true;}else{return false;}
}
</script>
</head>

<body bgcolor="#000000" leftmargin="0" topmargin="0">
<table width="423" height="324" border="0" cellpadding="0" cellspacing="0">
      <form name="form" method="post" action="changeroom.asp?act=changeroom" onsubmit='return(check());'>
<tr>
<td width="423" height="324" align="left" valign="top" background="pic/croom.jpg"> 
<table width="423" border="0" cellpadding="0" cellspacing="0">
<tr>
            <td><strong> </strong>ס���˻����»���,��һ��������С��,���¸������С��ȫ��1������,��������Ҫ�ؿ���. 
            </td>
</tr>
</table>
<table width="423" align="center">
<%rs.open "Select * from [housetype] Order by ht_���",conn,1,1
do while (Not RS.Eof)
%><tr><td>
<label><input type="radio" name="roomid" value="<%=rs("ht_���")%>" <%if roomlb=rs("ht_���") then%>disabled=1 <%end if%> onclick="show(<%=rs("ht_���")%>,1,'<%=rs("ht_С����")%>');" ><%=rs("ht_С����")%>1��:<%=replace(replace(rs("ht_1������"),";","&nbsp;"),"|",":")%></label>
</td></tr>
<%RS.MoveNext
Loop
%>
</table>
<div align="center"><input type="submit" name="Submit" value="ȷ������"></div>
</td>
</tr></form>
</table>
<div id="Layer1" style="position:absolute; width:189px; height:139px; z-index:2; left: 230px; top: 182px; visibility: hidden; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000;"><font color="blue" id="mess"></font></div>
