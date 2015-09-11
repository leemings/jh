<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,身份,内力,体力,武功,攻击,防御,等级,魅力,道德,银两 from 用户 where 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=440"
	response.end
end if
if rs("等级")<200 or rs("内力")<500000 or rs("体力")<1500000 or rs("武功")<550000 or rs("攻击")<340600 or rs("防御")<336700 or rs("魅力")<50000 or rs("道德")<10000 or rs("银两")<500000000 then
	Response.Write "<script Language=Javascript>alert('提示：你的状态未达到条件，请看清了再来！');window.close();</script>"
	response.end
end if
if rs("身份")="掌门" then
	Response.Write "<script Language=Javascript>alert('提示：你现在已经是掌门了，还想要多少门派呀！');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("门派")="官府" or (rs("门派")<>"无" and rs("门派")<>"游侠") then
	Response.Write "<script Language=Javascript>alert('提示：要想自创门派就必须首先离开现在的门派！');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
tmprs=conn.execute("Select count(*) As 数量 from 用户 where 等级>=80 and 介绍人='"& aqjh_name &"'")
lr=tmprs("数量")
set tmprs=nothing
if lr<8 then
	Response.Write "<script Language=Javascript>alert('提示：你的拉人记录不足8人，或你所拉的人的等级还没有到80级以上！');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
%>
<html>
<head>
<title>创派</title>
<link rel="stylesheet" href="../../css.css">
</head>
<BODY oncontextmenu=self.event.returnValue=false bgColor=#339966 text=#ffffff>
<p align="center"><img border="0" src="22.GIF" width="111" height="22"></p>
<table width="363" border="1" cellspacing="0" cellpadding="3" align="center" height="15" bordercolorlight="#999999" bordercolordark="#FFFFFF">
  <tr align="center"> 
    <td height="35">恭喜你来此自立门户</td>
  </tr>
  <tr> 
    <td height="108"> 
      <form method="post" action="newzt.asp">
        <p align="center">
          <table width="77%" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
            <tr> 
              <td height="36" align="center">门　　派：</td>
              <td height="36"> 
                <input type="text" name="subject" size="10" maxlength="4">
              </td>
            </tr>
            <tr>
              <td align="center">帮派简介：</td>
              <td>
                <input type="text" name="comment" size="20" maxlength="20">
              </td>
            </tr>
          </table>
        <p align="center">
          帮派简介最多不能超过50字。
        </p>
        <p align="center">
          <input type="submit" name="Submit" value="确认">
          <input type="reset" name="reset" value="取消"> 
        </p>
      </form>
 </td>
  </tr>  
</table>
</body>
</html>