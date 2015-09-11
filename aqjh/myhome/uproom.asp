<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if session("myroom")<>"h_拥有者" then
	Response.Write "<script Language=Javascript>alert('提示：你非此房间主人无权升级！');parent.email.style.visibility='hidden';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM house WHERE h_拥有者='" & aqjh_name &"'",conn,1,1
if rs.Eof or rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
	Response.Write "<script Language=Javascript>alert('提示：您并没有购买房子，不能升级！');top.close();</script>"
	Response.End
end if
myroomname=trim(rs("h_小区名"))
roomlb=rs("h_类型")
roomjb=rs("h_级别")
rs.close
rs.open "Select * from [housetype] where ht_序号="&roomlb,conn,1,1
if rs.Eof or rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
	Response.Write "<script Language=Javascript>alert('提示：你的房子类别出错，请与站长联系！');top.close();</script>"
	Response.End
end if
tj2=replace(replace(rs("ht_2级条件"),";","&nbsp;"),"|",":")
tj3=replace(replace(rs("ht_3级条件"),";","&nbsp;"),"|",":")
tj4=replace(replace(rs("ht_4级条件"),";","&nbsp;"),"|",":")
rs.close
set rs=nothing
conn.close
set conn=nothing
act=lcase(trim(Request("act")))
if act="uproom" then
	Response.Write "<html><body background=IMAGES/bg.jpg><div align=center><a href=javascript:history.go(-1);><font color=#0000FF>按这里返回</font></a></div>"
	uproomdj=clng(Request.form("uproomdj"))
	if uproomdj<>2 and uproomdj<>3 and uproomdj<>4 then
		Response.Write "<script Language=Javascript>alert('提示：请选择级别2-4级！');location.href = 'uproom.asp';</script>"
		Response.End
	end if
	if roomjb>=uproomdj then
		Response.Write "<script Language=Javascript>alert('提示：级别不正确,你现在的级别大于等于此级别！');location.href = 'uproom.asp';</script>"
		Response.End
	end if 
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
	rs.open "SELECT * FROM housetype WHERE ht_序号=" & roomlb,conn,1,1
	xlwp=rs("ht_"& uproomdj &"级条件")
	if isnull(xlwp) or xlwp="" or instr(xlwp,"|")=0 or instr(xlwp,";")=0 then 
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：房屋数据有错不能操作\n请找程序开发商联系!');</script>"
		response.end
	end if
	xadata=split(xlwp,";")
	xadata1=UBound(xadata)
	rs.close
	rs.open "SELECT w6,等级,会员等级,银两,金币 from [用户] WHERE 姓名='" & aqjh_name & "'",conn,1,3
	duyao=rs("w6")
	if isnull(duyao) or duyao="" then 
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：你什么物品也没有!');location.href = 'uproom.asp';</script>"
		response.end
	end if
	yin=0
	for ii=0 to xadata1
		xadata2=split(xadata(ii),"|")
		mysl=clng(xadata2(1))
		myxlwp=trim(xadata2(0))
		select case myxlwp
		case "等级"
			if rs("等级")<mysl then
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('提示：您的等级未到["& mysl &"]级，不能升级!');location.href = 'uproom.asp';</script>"
				response.end
			end if
		case "金币"
			if rs("金币")<mysl then
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('提示：您的金币没有["& mysl &"]个，不能升级!');location.href = 'uproom.asp';</script>"
				response.end
			end if
			rs("金币")=rs("金币")-mysl
		case "银两"
			yin=yin+mysl
			if rs("银两")<mysl then
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('提示：您的银两没有["& mysl &"]两，不能升级!');location.href = 'uproom.asp';</script>"
				response.end
			end if
			rs("银两")=rs("银两")-mysl
		case "会员等级"
			if rs("会员等级")<mysl then
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('提示：您的会员等级未到["& mysl &"]级，不能升级!');location.href = 'uproom.asp';</script>"
				response.end
			end if
		case else
		if mywpsl(duyao,myxlwp)<mysl then
    		Set rs = Nothing
   			conn.Close
  			Set conn = Nothing
   			Response.Write "<script Language=Javascript>alert('提示：物品"& myxlwp &"数量不足？');</script>"
			response.end
		end if
		duyao=abate(duyao,myxlwp,mysl)
		end select
	next
	rs("w6")=duyao
	rs.update
	rs.close
	rs.open "SELECT * FROM housetype WHERE ht_序号=" & roomlb,conn,1,3
	h_name=rs("ht_小区名")
	h_nai=rs("ht_"& uproomdj &"级耐久度")
	rs("ht_小区资产")=rs("ht_小区资产")+int(yin/10000)
	rs.update
	rs.close
	rs.open "SELECT * FROM house WHERE h_拥有者='" & aqjh_name &"'",conn,1,3
	rs("h_小区名")=h_name
	rs("h_级别")=uproomdj
	rs("h_耐久度")=h_nai
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=""Javascript"">alert(""恭喜你，房屋升级完成！"");location.href = 'uproom.asp';;</script>"
end if

%>
<LINK href="style.css" rel=stylesheet>
<script language="JavaScript" type="text/JavaScript">
function show(x,y,sm){Layer1.style.visibility="visible";mess.innerHTML="<img src=pic/hut"+x+y+".jpg width=185 height=125><br><div align=center>"+sm+y+"级</div>";}
function hidden(){Layer1.style.visibility="hidden";}
document.onmousedown=click;
function click(){if(event.button==2){alert("欢迎来到勇者爱情江湖！");}}
function shLayers(n,n1){
if(n.style.visibility=="visible"){n.style.visibility="hidden";}else if(n.style.visibility=="hidden"){n.style.visibility="visible";}
if(n1.style.visibility=="visible"){n1.style.visibility="hidden";}
}
function check(){
if(document.form.uproomdj[0].checked==false && document.form.uproomdj[1].checked==false && document.form.uproomdj[2].checked==false)
{alert("提示：请选择你要升级的级别！");return false;
}
}
</script>
<body bgcolor="#000000" leftmargin="0" topmargin="0">
<table width="423" height="324" border="0" cellpadding="0" cellspacing="0">
    <td width="423" height="324" align="left" valign="top" background="pic/uproom.jpg"> 
      <form name="form" method="post" action="uproom.asp?act=uproom" onsubmit='return(check());'>
<table width="98%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td><font color="#0000FF"><strong>爱情江湖小屋升级系统</strong></font><br>升级可不是一项简单的事情， 你需<br>
要准备各种物品，升过级的小 屋会<br>
有更新更漂亮的形像，功能对应的<br>
功能也会更加强大！</td>
</tr>
</table><br>
<br>
        请选择你所要处理的等级<br>
<table width="100%" align="center">
<tr>
<td><label>
<input type="radio" name="uproomdj" value="2" <%if roomjb>=2 then%>disabled=1 <%end if%> onclick="show(<%=roomlb%>,2,'<%=myroomname%>');" >2级:<%=tj2%></label></td>
</tr>
<tr>
<td><label>
<input type="radio" name="uproomdj" value="3" <%if roomjb>=3 then%>disabled=1 <%end if%> onclick="show(<%=roomlb%>,3,'<%=myroomname%>');">3级:<%=tj3%></label></td>
</tr>
<tr>
<td><label>
<input type="radio" name="uproomdj" value="4" <%if roomjb>=4 then%>disabled=1 <%end if%> onclick="show(<%=roomlb%>,4,'<%=myroomname%>');">4级:<%=tj4%></label></td>
</tr>
</table>
<div align="center">
<input type="submit" name="Submit" value="确定升级"></div>
<br></form>
</td>
</tr>
</table>
<div id="Layer1" style="position:absolute; width:189px; height:139px; z-index:2; left: 229px; top: 0px; visibility: hidden;"><font color="blue" id="mess"></font></div>










