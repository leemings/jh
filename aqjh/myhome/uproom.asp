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
set rs=nothing
conn.close
set conn=nothing
act=lcase(trim(Request("act")))
if act="uproom" then
	Response.Write "<html><body background=IMAGES/bg.jpg><div align=center><a href=javascript:history.go(-1);><font color=#0000FF>�����ﷵ��</font></a></div>"
	uproomdj=clng(Request.form("uproomdj"))
	if uproomdj<>2 and uproomdj<>3 and uproomdj<>4 then
		Response.Write "<script Language=Javascript>alert('��ʾ����ѡ�񼶱�2-4����');location.href = 'uproom.asp';</script>"
		Response.End
	end if
	if roomjb>=uproomdj then
		Response.Write "<script Language=Javascript>alert('��ʾ��������ȷ,�����ڵļ�����ڵ��ڴ˼���');location.href = 'uproom.asp';</script>"
		Response.End
	end if 
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
	rs.open "SELECT * FROM housetype WHERE ht_���=" & roomlb,conn,1,1
	xlwp=rs("ht_"& uproomdj &"������")
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
		Response.Write "<script Language=Javascript>alert('��ʾ����ʲô��ƷҲû��!');location.href = 'uproom.asp';</script>"
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
				Response.Write "<script Language=Javascript>alert('��ʾ�����ĵȼ�δ��["& mysl &"]������������!');location.href = 'uproom.asp';</script>"
				response.end
			end if
		case "���"
			if rs("���")<mysl then
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('��ʾ�����Ľ��û��["& mysl &"]������������!');location.href = 'uproom.asp';</script>"
				response.end
			end if
			rs("���")=rs("���")-mysl
		case "����"
			yin=yin+mysl
			if rs("����")<mysl then
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('��ʾ����������û��["& mysl &"]������������!');location.href = 'uproom.asp';</script>"
				response.end
			end if
			rs("����")=rs("����")-mysl
		case "��Ա�ȼ�"
			if rs("��Ա�ȼ�")<mysl then
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('��ʾ�����Ļ�Ա�ȼ�δ��["& mysl &"]������������!');location.href = 'uproom.asp';</script>"
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
	rs.open "SELECT * FROM housetype WHERE ht_���=" & roomlb,conn,1,3
	h_name=rs("ht_С����")
	h_nai=rs("ht_"& uproomdj &"���;ö�")
	rs("ht_С���ʲ�")=rs("ht_С���ʲ�")+int(yin/10000)
	rs.update
	rs.close
	rs.open "SELECT * FROM house WHERE h_ӵ����='" & aqjh_name &"'",conn,1,3
	rs("h_С����")=h_name
	rs("h_����")=uproomdj
	rs("h_�;ö�")=h_nai
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=""Javascript"">alert(""��ϲ�㣬����������ɣ�"");location.href = 'uproom.asp';;</script>"
end if

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
function check(){
if(document.form.uproomdj[0].checked==false && document.form.uproomdj[1].checked==false && document.form.uproomdj[2].checked==false)
{alert("��ʾ����ѡ����Ҫ�����ļ���");return false;
}
}
</script>
<body bgcolor="#000000" leftmargin="0" topmargin="0">
<table width="423" height="324" border="0" cellpadding="0" cellspacing="0">
    <td width="423" height="324" align="left" valign="top" background="pic/uproom.jpg"> 
      <form name="form" method="post" action="uproom.asp?act=uproom" onsubmit='return(check());'>
<table width="98%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td><font color="#0000FF"><strong>���ֽ���С������ϵͳ</strong></font><br>�����ɲ���һ��򵥵����飬 ����<br>
Ҫ׼��������Ʒ����������С �ݻ�<br>
�и��¸�Ư�������񣬹��ܶ�Ӧ��<br>
����Ҳ�����ǿ��</td>
</tr>
</table><br>
<br>
        ��ѡ������Ҫ����ĵȼ�<br>
<table width="100%" align="center">
<tr>
<td><label>
<input type="radio" name="uproomdj" value="2" <%if roomjb>=2 then%>disabled=1 <%end if%> onclick="show(<%=roomlb%>,2,'<%=myroomname%>');" >2��:<%=tj2%></label></td>
</tr>
<tr>
<td><label>
<input type="radio" name="uproomdj" value="3" <%if roomjb>=3 then%>disabled=1 <%end if%> onclick="show(<%=roomlb%>,3,'<%=myroomname%>');">3��:<%=tj3%></label></td>
</tr>
<tr>
<td><label>
<input type="radio" name="uproomdj" value="4" <%if roomjb>=4 then%>disabled=1 <%end if%> onclick="show(<%=roomlb%>,4,'<%=myroomname%>');">4��:<%=tj4%></label></td>
</tr>
</table>
<div align="center">
<input type="submit" name="Submit" value="ȷ������"></div>
<br></form>
</td>
</tr>
</table>
<div id="Layer1" style="position:absolute; width:189px; height:139px; z-index:2; left: 229px; top: 0px; visibility: hidden;"><font color="blue" id="mess"></font></div>










