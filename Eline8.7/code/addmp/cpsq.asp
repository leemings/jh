<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
session("sjjh_sj")=now()
Session("sjjh_jl")=0
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr
rs.open "select 门派,等级,内力,体力,魅力,攻击,银两 from 用户 where 姓名='"&sjjh_name&"'",conn
if trim(rs("等级"))<30 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的等级低于30所以暂时不能创派！！');window.close();</script>"
	response.end
end if
if trim(rs("内力"))<100000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的内力低于10万所以暂时不能创派！！');window.close();</script>"
	response.end
end if
if trim(rs("体力"))<100000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的体力低于10万所以暂时不能创派！！');window.close();</script>"
	response.end
end if
if trim(rs("攻击"))<10000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的攻击低于1万所以暂时不能创派！！');window.close();</script>"
	response.end
end if
if trim(rs("魅力"))<5000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的魅力低于5千所以暂时不能创派！！');window.close();</script>"
	response.end
end if

if trim(rs("银两"))<50000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的银两低于5千万所以暂时不能创派！！');window.close();</script>"
	response.end
end if
if trim(rs("门派"))<>"游侠" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您已经有门派了无法创派要创派请先脱离自己门派！！');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr
rs.Open ("select count(*) from 用户 where 介绍人='"&sjjh_name&"'"),conn
count=rs(0)
if (count<15) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你拉人数低于15人所以暂时不能创派你目前只为本江湖拉了！');window.close();}</script>"
	response.end
end if
%>

<html>

<head>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
body {font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt; color:'ffff77';
	scrollbar-face-color:'222222'; 
	scrollbar-shadow-color:'222222'; 
	scrollbar-highlight-color:'EFBA56'; 
	scrollbar-3dlight-color:'222222'; 
	scrollbar-darkshadow-color: 'AB7510'; 
	scrollbar-track-color:'FBEED7'; 
	scrollbar-arrow-color:'AB7510' ;

	}

td	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
div 	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
form	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
select	{font-size: 9pt; background-color:F7DEAC}
option	{font-size: 9pt; background-color:F7DEAC}
	
p	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
br	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
A 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:'ffff77'; font-size: 9pt;}
A:hover 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:ff4d4d;  font-size: 9pt}
.U1	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 10px; }
.U2	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 11px; }


.informat01{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color:'F7DEAC'}
	
.i02	{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; background-color: 'F7DEAC'; 
	color: '291C03'; border: 1 solid '291C03'}

.i03{ background-color:'F7DEAC'; 
	BORDER-TOP-WIDTH: 1px; 
	PADDING-RIGHT: 1px; PADDING-LEFT: 1px; 
	BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; 
	BORDER-LEFT-COLOR:'ffffff'; BORDER-BOTTOM-WIDTH: 1px; 
	BORDER-BOTTOM-COLOR:'EBA82C'; PADDING-BOTTOM: 1px; 
	BORDER-TOP-COLOR:'ffffff'; PADDING-TOP: 1px; 
	HEIGHT: 20px; BORDER-RIGHT-WIDTH: 1px; 
	BORDER-RIGHT-COLOR:'EBA82C'}

.i04{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'ffffff'; color: '333333'; 
	border: 333333px solid; background=ffffff; } 
	
.i05{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'F3CB81'; color: '83590C'; 
	border: 1px solid AB7510;} 
.i06{ background-color:'F7DEAC'; 
	border-top-width: 1px; 
	padding-right: 1px; padding-left: 1px; 
	border-left-width: 1px; font-size: 9pt; font-family: verdana,arial,helvetica; 
	border-left-color:'ffffff'; border-bottom-width: 1px; 
	border-bottom-color:'EBA82C'; padding-bottom: 1px; 
	border-top-color:'ffffff'; padding-top: 1px; 
	border-right-width: 1px; 
	border-right-color:'EBA82C'}
</style>
<style>.skin0 {
	BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; VISIBILITY: hidden; BORDER-LEFT: 1px solid; WIDTH: 80px; CURSOR: default; LINE-HEIGHT: 20px; BORDER-BOTTOM: 1px solid; FONT-FAMILY: Verdana; POSITION: absolute; BACKGROUND-COLOR: white
}
.menuitems {
	PADDING-RIGHT: 10px; PADDING-LEFT: 15px
}
                </style>



<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>自创门派</title>
</head>

<body bgcolor="#000000">

<div align="center">
  <center>
  <table border="0" width="80%" cellspacing="0" cellpadding="0" height="225" background="../picc/bg.jpg">
    <tr>
      <td width="100%" colspan="3" background="../picc/k2.gif" height="18">　</td>
    </tr>
    <tr>
      <td width="4%" rowspan="4" background="../picc/k3.gif" height="189">　</td>
      <td width="88%" align="center" height="26"><img border="0" src="22.gif" width="111" height="22"></td>
      <td width="4%" background="../picc/k3.gif" height="189" rowspan="4">　</td>
    </tr>
    <tr>
      <td width="88%" height="70" align="center"><img border="0" src="king3.jpg" width="503" height="256"></td>
    </tr>
    <tr>
      <td width="88%" height="145" valign="top">
        <table border="0" width="100%">

          <tr>
            <td width="100%" align="center">『快乐江湖』创立帮派请慎选帮派名称</td>
          </tr>

          <tr>
            <td width="100%" align="center"><form method="post" action="newzt.asp">
     门派名称：                           
            <input type="text" name="subject" size="37" style="background-color: #FFFFCC; color: #000000; border-style: solid; border-width: 1"> 
            <br> 
            帮派简介：                           
            <input type="text" name="comment" size="37" style="color: #000000; background-color: #FFFFCC; border-style: solid; border-width: 1"> 
            
     &nbsp;              
     <p> 
            
            <input type="submit" name="Submit" value="确认" style="background-color: #FFFFCC; color: #000000; border-style: solid; border-width: 1"> 
            <input type="reset" name="reset" value="取消" style="background-color: #FFFFCC; color: #000000; border-style: solid; border-width: 1">            
                 
     </p> 
                 
         </form> 
           
        <a href=# onclick="window.close()"><font color="#00FF00">关闭</font></a>  
</td> 
          </tr> 
         <td><hr> 
 
          <tr> 
            <td width="100%" align="center">『快乐江湖』</td> 
          </tr> 
        </table> 
      </td> 
    </tr> 
    
    <tr> 
      <td width="100%" colspan="3" background="../picc/k2.gif" height="18">　</td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>



