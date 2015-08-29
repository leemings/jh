<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 门派,身份,内力,体力,武功,攻击,防御,等级,魅力,道德,银两,保留3 from 用户 where 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=440"
	response.end
end if

if rs("等级")<30 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的等级不够30级，创什么派呀！');location.href = 'CHUANGP.ASP';</script>"
	Response.End
end if
if rs("内力")<500000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的内力不够500000，创什么派呀！');location.href = 'CHUANGP.ASP';</script>"
	Response.End
end if
if rs("体力")<1000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的体力不够1000000，创什么派呀！');location.href = 'CHUANGP.ASP';</script>"
	Response.End
end if
if rs("武功")<50000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的武功不够50000，创什么派呀！');location.href = 'CHUANGP.ASP';</script>"
	Response.End
end if
if rs("攻击")<40000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的攻击不够40000，创什么派呀！');location.href = 'CHUANGP.ASP';</script>"
	Response.End
end if
if rs("防御")<30000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的防御不够30000，创什么派呀！');location.href = 'CHUANGP.ASP';</script>"
	Response.End
end if
if rs("魅力")<10000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的魅力不够10000，创什么派呀！');location.href = 'CHUANGP.ASP';</script>"
	Response.End
end if
if rs("道德")<10000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的道德不够10000，创什么派呀！');location.href = 'CHUANGP.ASP';</script>"
	Response.End
end if
if rs("银两")<500000000  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的银两不够5亿，创什么派呀！');location.href = 'CHUANGP.ASP';</script>"
	Response.End
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
if rs("保留3")=true then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你曾经创过门派了，自己不好好珍惜还想作什么呀！');window.close();</script>"
	Response.End
end if
tmprs=conn.execute("Select count(*) As 数量 from 用户 where 等级>=16 and 介绍人='"& sjjh_name &"'")
lr=tmprs("数量")
set tmprs=nothing
if lr<8 then
	Response.Write "<script Language=Javascript>alert('提示：你的拉人记录不足8人，或你所拉的人的等级还没有到16级以上！');window.close();</script>"
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
<link rel="stylesheet" href="../dg/setup.css">
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




