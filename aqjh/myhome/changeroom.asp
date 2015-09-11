<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if session("myroom")<>"h_拥有者" then
	Response.Write "<script Language=Javascript>alert('提示：你非此房间主人无权换房！');parent.email.style.visibility='hidden';</script>"
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
act=lcase(trim(Request("act")))
if act="changeroom" then
	Response.Write "<html><body background=IMAGES/bg.jpg><div align=center><a href=javascript:history.go(-1);><font color=#0000FF>按这里返回</font></a></div>"
	roomid=clng(Request.form("roomid"))
	if roomid<0 or roomid>5 then
		Response.Write "<script Language=Javascript>alert('提示：房屋类别不对，类别为0-5！');location.href = 'changeroom.asp';</script>"
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End
	end if
	if roomid=roomlb then
		Response.Write "<script Language=Javascript>alert('提示：你选择的类型与你现在的房屋类型一至，不能操作！');location.href = 'changeroom.asp';</script>"
		set rs=nothing
		conn.close
		set conn=nothing
		Response.End
	end if 
	rs.open "SELECT * FROM housetype WHERE ht_序号=" & roomid,conn,1,1
	xlwp=rs("ht_1级条件")
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
		Response.Write "<script Language=Javascript>alert('提示：你什么物品也没有!');location.href = 'changeroom.asp';</script>"
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
				Response.Write "<script Language=Javascript>alert('提示：您的等级未到["& mysl &"]级，不能更换房屋!');location.href = 'changeroom.asp';</script>"
				response.end
			end if
		case "金币"
			if rs("金币")<mysl then
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('提示：您的金币没有["& mysl &"]个，不能更换房屋!');location.href = 'changeroom.asp';</script>"
				response.end
			end if
			rs("金币")=rs("金币")-mysl
		case "银两"
			yin=yin+mysl
			if rs("银两")<mysl then
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('提示：您的银两没有["& mysl &"]两，不能更换房屋!');location.href = 'changeroom.asp';</script>"
				response.end
			end if
			rs("银两")=rs("银两")-mysl
		case "会员等级"
			if rs("会员等级")<mysl then
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('提示：您的会员等级未到["& mysl &"]级，不能更换房屋!');location.href = 'changeroom.asp';</script>"
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
	rs.open "SELECT * FROM housetype WHERE ht_序号=" & roomid,conn,1,3
	h_name=rs("ht_小区名")
	h_nai=rs("ht_1级耐久度")
	rs("ht_小区资产")=rs("ht_小区资产")+int(yin/10000)
	rs.update
	rs.close
	rs.open "SELECT * FROM house WHERE h_拥有者='" & aqjh_name &"'",conn,1,3
	rs("h_小区名")=h_name
	rs("h_级别")=1
	rs("h_耐久度")=h_nai
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=""Javascript"">alert(""恭喜你，房屋更换完成！"");location.href = 'changeroom.asp';;</script>"
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
Layer1.style.visibility="visible";mess.innerHTML="<img src=pic/hut"+x+y+".jpg width=185 height=125><br><div align=center>"+sm+y+"级</div>";
}
function hidden(){Layer1.style.visibility="hidden";}
document.onmousedown=click;
function click(){if(event.button==2){alert("欢迎来到勇者爱情江湖！");}}
function shLayers(n,n1){
if(n.style.visibility=="visible"){n.style.visibility="hidden";}else if(n.style.visibility=="hidden"){n.style.visibility="visible";}
if(n1.style.visibility=="visible"){n1.style.visibility="hidden";}
}
function check(){
if(document.form.roomid[0].checked==false && document.form.roomid[1].checked==false && document.form.roomid[2].checked==false && document.form.roomid[3].checked==false && document.form.roomid[4].checked==false && document.form.roomid[5].checked==false)
{alert("提示：请选择你要换的房屋类型！");return false;}
if(confirm("现在将要进行换房屋操作，你原来的房屋将会\n被系统回收，并将无法恢复，你确定购买新房吗？")){return true;}else{return false;}
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
            <td><strong> </strong>住久了换换新环境,试一下其它的小区,重新更换后的小屋全以1级计算,所以请大家要重考虑. 
            </td>
</tr>
</table>
<table width="423" align="center">
<%rs.open "Select * from [housetype] Order by ht_序号",conn,1,1
do while (Not RS.Eof)
%><tr><td>
<label><input type="radio" name="roomid" value="<%=rs("ht_序号")%>" <%if roomlb=rs("ht_序号") then%>disabled=1 <%end if%> onclick="show(<%=rs("ht_序号")%>,1,'<%=rs("ht_小区名")%>');" ><%=rs("ht_小区名")%>1级:<%=replace(replace(rs("ht_1级条件"),";","&nbsp;"),"|",":")%></label>
</td></tr>
<%RS.MoveNext
Loop
%>
</table>
<div align="center"><input type="submit" name="Submit" value="确定升级"></div>
</td>
</tr></form>
</table>
<div id="Layer1" style="position:absolute; width:189px; height:139px; z-index:2; left: 230px; top: 182px; visibility: hidden; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000;"><font color="blue" id="mess"></font></div>
