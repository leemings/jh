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
rs.open "select id,门派,身份,内力,体力,武功,攻击,防御,等级,魅力,道德,银两 from 用户 where 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=440"
	response.end
end if
aqjh_id=rs("id")
if rs("等级")<35 or rs("内力")<100000 or rs("体力")<500000 or rs("武功")<50000 or rs("攻击")<40600 or rs("防御")<36700 or rs("魅力")<40000 or rs("道德")<50000 or rs("银两")<100000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=460"
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
tmprs=conn.execute("Select count(*) As 数量 from 用户 where 等级>=18 and 介绍人='"& aqjh_name &"'")
lr=tmprs("数量")
set tmprs=nothing
if lr<5 then
	Response.Write "<script Language=Javascript>alert('提示：你的拉人记录不足5人，或你所拉的人的等级还没有到18级以上！');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
function chuser(u)
  dim filter,xx,usernameenable,su
  for i=1 to len(u)
    su=mid(u,i,1)
    xx=asc(su)
    zhengchu = -1*xx \ 256
    yushu = -1*xx mod 256
    if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
      chuser=true
      exit function
    end if
  next
  chuser=false
end function
if trim(request.form("subject"))="" or trim(request.form("comment"))=""  then%>
	<script language=vbscript>
	  MsgBox "错误：帮派名称和简介必须填写！"
	  location.href = "javascript:history.back()"
	</script>
	<%
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
mypai1=trim(request.form("subject"))
jjjj=request.form("comment")
namelen=0
for i=1 to len(mypai1)
	zh=mid(mypai1,i,1)
	zhasc=asc(zh)
	if zhasc<0 then
		namelen=namelen+2
	else
		namelen=namelen+1
	end if
next
if namelen>10 or namelen<4 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=027"
end if
if zhasc>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=027"
end if
namelen=0
if len(mypai1)>10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect"../../error.asp?id=027"
end if
if len(jjjj)>50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=028"
end if
if mypai1="无" or mypai1="六扇门" or mypai1=" 六扇门" or mypai1="衙门" or mypai1="官府" or mypail=" 官府" or mypai1="政府" or mypai1="游侠" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=027"
end if
if InStr(mypai1,"逼")<>0 or InStr(mypai1,"无")<>0 or InStr(mypai1,"官")<>0 or InStr(mypai1,"管")<>0 or InStr(mypai1,"衙")<>0 or InStr(mypai1,"爸")<>0  or  InStr(mypai1,"祖宗")<>0  or Instr(mypai1,"爷")<>0  or Instr(mypai1,"爹")<>0  or Instr(mypai1,"妈")<>0 or Instr(mypai1,"靠")<>0 or Instr(mypai1,"叼")<>0 or Instr(mypai1,"操")<>0 or Instr(mypai1,"扇")<>0 or Instr(mypai1,"娘")<>0 or Instr(mypai1,"管")<>0 or Instr(mypai1,"性")<>0 or Instr(mypai1,"阴")<>0 or Instr(mypai1,"色")<>0 or Instr(mypai1,"天网")<>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=027"
end if
if left(mypai1,3)="%20" or InStr(mypai1,"=")<>0 or InStr(mypai1,"`")<>0 or InStr(mypai1,"'")<>0 or InStr(mypai1," ")<>0 or InStr(mypai1,"　")<>0 or InStr(mypai1,"'")<>0 or InStr(mypai1,chr(34))<>0 or InStr(mypai1,"\")<>0 or InStr(mypai1,",")<>0 or InStr(mypai1,"<")<>0 or InStr(mypai1,">")<>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=027"
end if
if chuser(mypai1) then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=027"
end if
if left(jjjj,3)="%20" or instr(jjjj,"or")<>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=027"
end if
sql="Select a from p where a='"& mypail &"'"
set rs=conn.execute(sql)
if not(rs.bof or rs.eof) then %>
	<script language=vbscript>
	  MsgBox "错误：帮派名称已经存在！"
	  location.href = "javascript:history.back()"
	</script>
	<%
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
'e加入限制
conn.execute "Insert Into p (a,b,c,d,e,g,h,i) Values('"&mypai1&"','"&aqjh_name&"',0,'" & jjjj & "','加入无限制','无',50000000,'"&now()&"')"
conn.execute "update 用户 set 门派='"&mypai1&"',身份='掌门',银两=银两-100000000,grade=5 where 姓名='"&aqjh_name&"'"
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body {font-size:9pt;SCROLLBAR-FACE-COLOR: #c2cea8;SCROLLBAR-HIGHLIGHT-COLOR: #ffffff;SCROLLBAR-SHADOW-COLOR: #000000;SCROLLBAR-3DLIGHT-COLOR: #ffffff;SCROLLBAR-ARROW-COLOR: #ffffff;SCROLLBAR-TRACK-COLOR: #ffffff;SCROLLBAR-DARKSHADOW-COLOR: #ffffff;SCROLLBAR-BASE-COLOR: #c2cea8}
table {font-size: 9pt}
a:link {CURSOR:crosshair;color: yellow; text-decoration: none}
a:visited {CURSOR:crosshair;color: yellow; text-decoration: none}
a:active {CURSOR:crosshair;color: #white; text-decoration: none}
a:hover { CURSOR:crosshair;color: #white; text-decoration: underline overline}
-->
</style>
<script language="javascript">
function opench(url,win){
        controlWindow=window.open(url,win,"fullscreen=no,toolbar=yes,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=20,left=2,width=450,height=300");
        controlWindow.focus();
        }
function ch(){
	document.forms[0].userpassword.value=document.forms[0].pwd.value;
	document.forms[0].pwd.value="";
//	document.forms[0].submit();
	}
</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=#000000 text=#ffffff link="#000080" alink="#800000" vlink="#000080">
<div align="center"> 
  <p><font size="3"><b>帮派已成功添加！</b></font></p>
</div>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>