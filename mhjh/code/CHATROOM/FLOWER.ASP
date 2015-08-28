<%@ LANGUAGE = VBScript.Encode %>
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
chatroomsn=session("yx8_mhjh_userchatroomsn")
allonline=split(Application("yx8_mhjh_allonline"),";")
allonlineubd=ubound(allonline)
nickname=session("yx8_mhjh_username")
id=request.QueryString("ID")
if id<>"1" then
%>
<html>
<head>
<style>
body,p{font-size:9pt; color:0}
select,input{font-size:9pt; color:#000000;}
</style>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<style>
body,p{font-size:14; color:red; line-height:16pt}
</style>
<title>送礼</title>
</head>
<body oncontextmenu=self.event.returnValue=false background="bg1.gif" bgproperties="fixed" text="#FFFFFF" leftmargin='3'>
<center>
  <p style='font-size:12pt;color:red'><font color="#FF0000">江湖送花</font></p>
<img id=flower src="flower/f21.gif" alt=花的图案><br>
<form method=POST action='flower.asp?id=1'>
<select name=flower size=1 onChange="document.images['flower'].src=options[selectedIndex].value;">
<option value="flower/f00.gif">花00
<option value="flower/f01.gif">花01
<option value="flower/f02.gif">花02
<option value="flower/f03.gif">花03
<option value="flower/f04.gif">花04
<option value="flower/f05.gif">花05
<option value="flower/f06.gif">花06
<option value="flower/f07.gif">花07
<option value="flower/f08.gif">花08
<option value="flower/f09.gif">花09
<option value="flower/f10.gif">花10
<option value="flower/f11.gif">花11
<option value="flower/f12.gif">花12
<option value="flower/f13.gif">花13
<option value="flower/f14.gif">花14
<option value="flower/f15.gif">花15
<option value="flower/f16.gif">花16
<option value="flower/f17.gif">花17
<option value="flower/f18.gif">花18
<option value="flower/f19.gif">花19
<option value="flower/f20.gif">花20
<option value="flower/f21.gif" selected>花21
<option value="flower/f22.gif">花22
<option value="flower/f23.gif">花23
<option value="flower/f24.gif">花24
<option value="flower/f25.gif">花25
<option value="flower/f26.gif">花26
<option value="flower/f33.gif">花33
<option value="flower/f34.gif">花34
<option value="flower/f35.gif">花35
<option value="flower/f36.gif">花36
<option value="flower/f37.gif">花37
<option value="flower/f38.gif">花38
<option value="flower/f39.gif">花39
<option value="flower/f40.gif">花40
<option value="flower/g01.gif">果01
<option value="flower/g02.gif">果02
<option value="flower/g27.gif">果27
<option value="flower/g28.gif">果28
<option value="flower/g29.gif">果29
<option value="flower/g30.gif">果30
<option value="flower/g31.gif">果31
<option value="flower/g32.gif">果32
<option value="flower/g33.gif">果33
<option value="flower/g34.gif">果34
<option value="flower/g35.gif">果35
<option value="flower/g36.gif">果36
<option value="flower/g37.gif">果37
<option value="flower/g38.gif">果38
<option value="flower/g39.gif">果39
<option value="flower/h40.gif">壶01
<option value="flower/h41.gif">壶02
<option value="flower/h42.gif">瓶01
<option value="flower/h43.gif">瓶02
<option value="flower/h44.gif">瓶03
<option value="flower/h45.gif">瓶04
<option value="flower/h46.gif">瓶05
<option value="flower/h47.gif">杯01
<option value="flower/h48.gif">杯02
<option value="flower/i49.gif">物01
<option value="flower/i50.gif">物02
<option value="flower/i51.gif">物03
<option value="flower/j65.gif">物04
<option value="flower/j52.gif">糖果
<option value="flower/i53.gif">饼干
<option value="flower/i54.gif">食物
<option value="flower/i55.gif">蛋羔
<option value="flower/i56.gif">面条
<option value="flower/i57.gif">米饭
<option value="flower/i58.gif">蛋羔
<option value="flower/I59.gif">蛋羔
<option value="flower/I60.gif">食物
<option value="flower/I61.gif">面包
<option value="flower/I62.gif">小凡
<option value="flower/j63.gif">药凡
<option value="flower/j64.gif">针筒
<option value="flower/j66.gif">灯泡
<option value="flower/j68.gif">纸卷
<option value="flower/j69.gif">锁匙
<option value="flower/j70.gif">香01
</select><br>
    <p><font size="2">送花给谁：</font> <font size="2">    
      <select name="towho">
<%for i=1 to allonlineubd-1
if allonline(i)<>Application("yx8_mhjh_admin") then
Response.Write "<option value='"&allonline(i)&"'>"&allonline(i)&"</option>"
end if
next
%>
      </select>
    <p>想说的话： <font size="2"> 
      <input value="我爱你" size=15 name=says maxlength=50>
      </font> 
    <p><font size="2">
      <input type=submit value="确定">
      </font>
  </form>
</center>
  <font color="#000000">注意：<br>
送礼需要100000两<br>会员优惠50000两     
<center><b><%=msg%></b></center>
  </font>
</body>
</html>
<%
else
	image=left(request.form("flower"),20)
	image=lcase(image)
	image=replace(image,"http","")
	towho=request.form("towho")
	says=request.form("says")
      says=Server.HtmlEncode(says)
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
if towho="" then Response.end
rs.open "select 会员,银两 FROM 用户 WHERE 姓名='"&nickname&"'",conn
hy=rs("会员")
yl=rs("银两")
rs.close
if yl<100000 then 
msg="对不起，你的银两不够，需要100000两才能够送礼"
abc=0
set rs=nothing
conn.close
set conn=nothing
else
	 if towho=nickname then
      Response.Write "<script language=JavaScript>{alert('"&nickname&",不能自己给自己送!');parent.history.go(-1);}</script>"
	 Response.End
	 end if
msg="恭喜你，你花了100000两从魔幻仙子那里买了礼物，深情地送给了"&towho&"."
abc=1
song="捧着从<font color=red>魔幻仙子</font>那里花了<b><font color=red>100000</font></b>两买来的<img src="&image&">，("&towho&"资质加1000)，对着<b><font color=green>"&towho&"</font></b>深情的说:<marquee width=100% behavior=alternate scrollamount=10><font color=red><font face=Webdings class=p>&yuml;</font></font>" & says&  "</marquee>"
if hy=true then
conn.execute "update 用户 set 银两=银两-50000 where 姓名='"&nickname&"'"
conn.execute "update 用户 set 资质=资质+1000 where 姓名='"&towho&"'"
else
conn.execute "update 用户 set 银两=银两-100000 where 姓名='"&nickname&"'"
conn.execute "update 用户 set 资质=资质+1000 where 姓名='"&towho&"'"
end if
Response.Write "<script language=JavaScript>{alert('[送礼物成功]');parent.history.go(-1);}</script>"	
set rs=nothing
conn.close
set conn=nothing
end if
end if
if abc=1 then
dim newtalkarr(600)
talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=nickname
		newtalkarr(595)=towho
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=#E7761B>【江湖送礼】"& nickname &""& song &"</font>"
        newtalkarr(600)=chatroomsn
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr	
end if
%>

