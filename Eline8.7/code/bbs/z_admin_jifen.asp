<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->

<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。"
		call dvbbs_error()
	else
		call main()
	end if

	sub main()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
  <tr>
    <td>
      <table cellpadding=3 cellspacing=1 border=0 width=100%>
        <tr >
        <th>欢迎<b><%=membername%></b>进入管理中心</th>
        </tr>
            <tr>
          <td width="100%" valign=top>
            <table width="90%" class=tb cellspacing="0" cellpadding="0" align="center" height="22">
              <tr>
                <td align="center" valign="middle">社区积分 V1.2 管理中心</td>
              </tr>
            </table>
            <br>
            <table width="90%" class=tb cellspacing="0" cellpadding="0" align="center" height="22">
              <tr> 
                <td align="left" valign="middle"> 
                  <p> 1．注意事项： 在下面，您将看到对社区积分的设定，如果想要修改，请在相应的地方输入你的设置，自定能力更加强大！<br>
                    <br>
                    2. 设置的值不要太小，也不要太大；太小会使各项属性增加太多，太大增加的属性值有很少，甚至为零；因为计算中是取整的，如果在线时间行<font color="#FF6633">小于</font>设置的折算率 
                    结果将会是零。</p>
                  <p>3. <font color="#9933CC">时间间隔</font>是指在线时间必须<font color="#FF6600">大于</font>在这个时间才可以保存积分，保存积分之后在线时间将会自动重置为零，重新开始计算；这个时间的设置最为重要，要慎重！<font color="#9933CC">财产折算率(W)</font>是用来计算增加的财产值，这个值可以小一点，一般可以<font color="#FF9933">小于</font>时间间隔的值；<font color="#9933CC">魅力折算率(C)</font>可以尽量设大一点，一般为这四个值中最大的一个，计算中如果结果为零，增加的魅力值将会置为1，所有这个值大也不怕。</p>
                  <p>4. 设置值必须在0～99之间，否则将会截取前两位作为新值<br>
                    <br>
                    3.<font color="#FF0000"> 特别注意：</font>新设定的值将会覆盖掉原来所设置的值，请小心使用！！！</p></td>
              </tr>
            </table>
            <br>
            <form action="z_admin_jifen.asp" method="post">
              <table width="90%" class=tb cellspacing="0" cellpadding="0" align="center" height="249">
                <tr> 
                  <td align="center" valign="middle"> 
                    <%
	dim fsObj,txtObj
	dim jf_StartTime,jf_WealthRate,jf_CPRate,jf_EPRate     '积分间隔时间，计算财产比率，计算魅力比率，计算经验比率
	dim SuanfaValue   '算法选择
	 
if request.form("submit")<>"提交"  then
	Set fsObj=Server.CreateObject("Scripting.FileSystemObject")
	set txtObj = fsObj.OpenTextFile(Server.MapPath("./")&"\z_jifen.inc",1)
	txtObj.skipline()
'jf_StartTime	
	txtObj.skipline()
	txtObj.skip(13)
	jf_StartTime=txtObj.read(2)
'jf_WealthRate	
	txtObj.skipline()
	txtObj.skip(14)
	jf_WealthRate=txtObj.read(2)
'jf_CPRate	
	txtObj.skipline()
	txtObj.skip(10)
	jf_CPRate=txtObj.read(2)
'jf_EPRate	
	txtObj.skipline()
	txtObj.skip(10)
	jf_EPRate=txtObj.read(2)
'SuanfaValue			
	txtObj.skipline()
	txtObj.skip(12)
	SuanfaValue=txtObj.read(1)
	
	set txtObj=nothing
	set fsObj=nothing
%>
                    <p>&nbsp;</p>
                    <p>==================================== <font color="#0066FF">参数设置</font> 
                      ====================================</p>
                    <p> 时间间隔(分钟)：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                      <input type="text" name="jf_StartTime" size="20" value="<%=jf_StartTime%>">
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;可以保存积分的时间间隔<br>
                      财产折算率(W)：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                      <input type="text" name="jf_WealthRate" size="20" value="<%=jf_WealthRate%>">
                      &nbsp;&nbsp;&nbsp;&nbsp; 财产值 = 在线时间 / W <br>
                      魅力折算率(C)：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                      <input type="text" name="jf_CPRate" size="20" value="<%=jf_CPRate%>">
                      &nbsp;&nbsp;&nbsp;&nbsp; 魅力值 = 在线时间 / C <br>
                      经验折算率(E)：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                      <input type="text" name="jf_EPRate" size="20" value="<%=jf_EPRate%> ">
                      &nbsp;&nbsp;&nbsp;&nbsp; 经验值 = 在线时间 / E </p>
                    <p><font color="#FF99CC">参考值：时间间隔:T=<font color="#FF0000">20</font> 
                      分钟, 财产折算率: W=<font color="#FF0000">10</font>, 魅力折算率: C=<font color="#FF0000">30</font>, 
                      经验折算率: E=<font color="#FF0000">20</font></font></p>
                    <p> ==================================== <font color="#0066FF">算法选择</font> 
                      ====================================<br>
                      <br>
                      <input type="radio" name="suanfa" value="1" <%if SuanfaValue="1" then response.write("checked") end if%>>
                      算法一: 一般算法<br>
                      <br>
                      财产值 = 在线时间 / W <br>
                      魅力值 = 在线时间 / C <br>
                      经验值 = 在线时间 / E <br>
                      <br>
                      <input type="radio" name="suanfa" value="2" <%if SuanfaValue="2" then response.write("checked") end if%>>
                      算法二: 四舍五入算法<br>
                      <br>
                      财产值 = (在线时间+5) / W <br>
                      魅力值 = (在线时间+5) / C <br>
                      经验值 = (在线时间+5) / E <br>
                      <br>
                      <input type="radio" name="suanfa" value="3" <%if SuanfaValue="3" then response.write("checked") end if%>>
                      算法三: 四舍五入、最小值不小于1算法<br>
                      <br>
                      财产值 = MAX（(在线时间+5) / W ，1）<br>
                      魅力值 = MAX（(在线时间+5) / C ，1）<br>
                      经验值 = MAX（(在线时间+5) / E ，1）</p>
                    <p><font color="#FF99CC">推荐算法：算法三</font><br>
                    </p>
                    <p> 
                      <input type="submit" name="Submit" value="提交">
                      　　　　　 
                      <input type="reset" name="Submit2" value="重置">
                    </p>
                    <p>&nbsp; </p>
                    <%
else
    jf_StartTime=trim(request.form("jf_StartTime"))
    jf_WealthRate=trim(request.form("jf_WealthRate"))
    jf_CPRate=trim(request.form("jf_CPRate"))
    jf_EPRate=trim(request.form("jf_EPRate"))
    SuanfaValue=trim(request.form("Suanfa"))
	
	Set fsObj=Server.CreateObject("Scripting.FileSystemObject")
	set txtObj = fsObj.OpenTextFile(Server.MapPath("./")&"\z_jifen.inc",2,true)
	txtObj.writeline("<%")
	txtObj.writeline("dim jf_StartTime,jf_WealthRate,jf_CPRate,jf_EPRate,SuanfaValue    '积分间隔时间，计算财产比率，计算魅力比率，计算经验比率，算法选择")
	txtObj.write("jf_StartTime=")
	txtObj.writeline(jf_StartTime)
	txtObj.write("jf_WealthRate=")
	txtObj.writeline(jf_WealthRate)
	txtObj.write("jf_CPRate=")
	txtObj.writeline(jf_CPRate)
	txtObj.write("jf_EPRate=")
	txtObj.writeline(jf_EPRate)
	txtObj.write("SuanfaValue=")
	txtObj.writeline(SuanfaValue)			
	txtObj.writeline("")
	txtObj.writeline("")
	txtObj.writeline("%\>")
	set txtObj=nothing
	set fsObj=nothing
	response.write("恭喜您，设置成功！！！<br>")
end if
%>
                  </td>
                </tr>
              </table>
            </form>
			<br><br>
          </td>
            </tr>
        </table>
        </td>
    </tr>
</table>
<br>
<br>
<br>
<%
end sub
%>