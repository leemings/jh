<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if session("myroom")<>"h_拥有者" then
	Response.Write "<script Language=Javascript>alert('提示：你非此房间主人无权修理！');parent.email.style.visibility='hidden';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM house WHERE h_拥有者='" & aqjh_name &"'",conn,1,3
if rs.Eof or rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
	Response.Write "<script Language=Javascript>alert('提示：您并没有购买房子，不能修理！');top.close();</script>"
	Response.End
end if
myroomname=trim(rs("h_小区名"))
roomlb=rs("h_类型")
roomjb=rs("h_级别")
naijd=rs("h_耐久度")
if naijd>=90 then
	mshow="完好无需修理!"
elseif naijd>=80 then
	mshow="旧了但完全没问题!"
elseif naijd>=70 then
	mshow="房子有一点损坏！"
elseif naijd>=50 then
	mshow="房子都破了,还漏雨!"
elseif naijd>=40 then
	mshow="这样的房子还会有人住?"
elseif naijd>=30 then
	mshow="除了四边的墙没有什么了!"
elseif naijd>=20 then
	mshow="天当被地当床,身边一堵破山墙!"
else
	mshow="这里不能再住人了!"
end if
rs.close
act=lcase(trim(Request("act")))
if act="xiuli" then
	Submit=trim(Request.form("Submit"))
	if Submit="确定修理" then
		rs.open "SELECT 豆点 FROM [用户] WHERE 姓名='" & aqjh_name &"'",conn,1,3
		mf=(1000+((roomjb-1)*500))
		if rs("豆点")<mf then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('提示：你豆点点数不足["& mf &"]点不能修理！');location.href = 'xiuli.asp';</script>"
			Response.End
		end if
		rs("豆点")=rs("豆点")-mf
		rs.update
		rs.close
		rs.open "SELECT * FROM house WHERE h_拥有者='" & aqjh_name &"'",conn,1,3
		rs("h_耐久度")=rs("h_耐久度")+50
		rs.update
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：消费豆点"& mf &"点耐久度增加50点！');location.href = 'xiuli.asp';</script>"
		Response.End
	end if
end if
set rs=nothing
conn.close
set conn=nothing
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
</script>
<title>修理房屋</title>
<body background="IMAGES/bg.jpg" text="#FFFFFF" leftmargin="0" topmargin="0">
<table width="423" height="324" border="0" cellpadding="0" cellspacing="0">
    <td width="423" height="324" align="left" valign="top"> 
        <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><font color="#FFFF00"><strong>修 理 房 屋</strong></font><br>
              经过一段时间的使用,你的小屋经过风风雨雨也会损坏,所以就要请我这个工匠来进行修理我可以按房屋等级收费的,当耐久值小于<font color="#FFFF00"><strong>50</strong></font>点时可以找我修复每次修复增加耐久度<font color="#FFFF00"><strong>50</strong></font>点,每次收费<font color="#FFFF00"><strong>1000</strong></font>个豆点点数,房屋每提高1级加收手续费<font color="#FFFF00"><strong>50%</strong></font>,别说我黑,这年头为了生活大家彼此彼此....<br>
              <br>
              状态:<%=naijd%>点<br>
              <br>
              建议:<%=mshow%><br>
              <br>
              <br>
<%if naijd<=50 then%>
<form name="form" method="post" action="xiuli.asp?act=xiuli" >
<input type="submit" name="Submit" value="确定修理"> </td>
</form><%end if%>
</tr>
</table>
        <br>
        <br>

</td>
</tr>
</table>
<div id="Layer0" style="position:absolute; width:200px; height:115px; z-index:3; left: 195px; top: 76px;"><img src="pic/xiuli.gif" width="223" height="247"></div>












