<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "Select * from 用户 where 姓名='" & sjjh_name &"'",conn,3,3
mygj=rs("等级")*sjjh_gjsx+4500 
myfy=rs("等级")*sjjh_fysx+3000
if rs("攻击")>mygj then
	conn.execute("update 用户 set 攻击="& mygj &" where 姓名='"&sjjh_name&"'")
end if
if rs("防御")>myfy then
	conn.execute("update 用户 set 防御="& myfy &" where  姓名='"&sjjh_name&"'")
end if
if rs("会员")=true and clng(DateDiff("d",now(),rs("会员结束")))>0 then
	pdstr="<font size=-1>[泡点制会员]"&rs("会员结束")&"</font>"
else
	pdstr=""
end if
wgj=rs("武功加")
nlj=rs("内力加")
tlj=rs("体力加")
%>
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<title>快乐江湖 - [<%=sjjh_name%>]状态</title>
<body style="BACKGROUND-COLOR: buttonface; BORDER-BOTTOM: medium none; BORDER-LEFT: medium none; BORDER-RIGHT: medium none; BORDER-TOP: medium none; PADDING-BOTTOM: 6px; PADDING-LEFT: 6px; PADDING-RIGHT: 6px; PADDING-TOP: 6px" marginwidth="0" marginheight="0" leftmargin="0" topmargin="10" LANGUAGE="javascript"  >
<table width="280" border="1" cellspacing="0" cellpadding="0" class="p150" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="left">
  <tr valign="top" align="left"> 
    <td height="250" colspan="4" class="p9"> 
      <!--#include file="../z_showvisualmy.asp"-->
      <p>&nbsp;</p></td>
  </tr>
  <tr bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9"><font color="#000000" size="2">年龄:</font></td>
    <td width="26%"><font size="2"><%=rs("年龄")%></font>　</td>
    <td width="13%"><font color="#000000" size="2">职业:</font></td>
    <td width="26%"><font size="2"><%=rs("职业")%></font>　</td>
  </tr>
  <tr bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9"><font color="#000000" size="2">师傅:</font></td>
    <td width="26%"><font size="2"><%=rs("师傅")%></font>　</td>
    <td width="13%"><font color="#000000" size="2">门派:</font></td>
    <td width="26%"><font size="2"><%=rs("门派")%></font>　</td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9" height="8"><font color="#000000" size="2">地区:</font></td>
    <td width="26%" height="8"><font size="2"><%=rs("地区")%></font></td>
    <td width="13%" height="8"><font color="#000000" size="2">宝物:</font></td>
    <td width="26%" height="8"><font size="2"><%=rs("宝物")%></font></td>
  </tr>
  <%if rs("会员等级")>0 then%>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9"><font color="#000000" size="2">会员:</font></td>
    <td width="26%"><font size="2"><%=rs("会员等级")%></font>　</td>
    <td width="13%"><font color="#000000" size="2">会期</font></td>
    <td width="26%"><font size="2"><%=rs("会员日期")%></font>　</td>
  </tr>
  <%end if%>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9"><font color="#000000" size="2">杀人:</font></td>
    <td width="26%"><font size="2"><%=rs("杀人数")%><font color="#FF0000">/</font><font size="2" color="#FF0000"><b><%=int(Application("sjjh_killman"))%></b></font></font></td>
    <td width="13%"><font color="#000000" size="2">总数:</font></td>
    <td width="26%"><font size="2"><%=rs("总杀人")%>人</font></td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9"><font size="2" color="#000000">等级: </font></td>
    <td width="26%"><font size="2"><%=rs("等级")%></font>　</td>
    <td width="13%"><font size="2" color="#000000">管理: </font></td>
    <td width="26%"><font size="2" color="#000000"><%=rs("grade")%></font>　</td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9" height="2"><font size="2" color="#000000">银子:</font></td>
    <td width="26%" height="2" nowrap><font size="2"><%=rs("银两")%></font></td>
    <td width="13%" height="2" nowrap><font size="2" color="#000000">存款: </font></td>
    <td width="26%" height="2" nowrap><font size="2"><%=rs("存款")%></font></td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9"><font size="2" color="#000000">存点: </font></td>
    <td width="26%"><font size="2"><%=rs("allvalue")%></font>　</td>
    <td width="13%"><font size="2" color="#000000">豆点:</font></td>
    <td width="26%"><font size="2"> 
      <%if rs("grade")=3 and rs("身份")="护法" then%>
      <%=rs("泡豆点数")-int(rs("泡豆点数")/500)*500%><font color="#FF0000">/500</font> 
      <%else%>
      <%=rs("泡豆点数")-int(rs("泡豆点数")/1000)*1000%><font color="#FF0000">/1000</font> 
      <%end if%>
      </font></td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9"><font size="2" color="#000000">豆子: </font></td>
    <td width="26%"><font size="2"><font color=red> 
      <%if rs("grade")=3 and rs("身份")="护法" then%>
      <%=int(rs("泡豆点数")/500)%> 
      <%else%>
      <%=int(rs("泡豆点数")/1000)%> 
      <%end if%>
      </font></font>　</td>
    <td width="13%"><font size="2" color="#000000">练功: </font></td>
    <td width="26%"><font size="2"> 
      <%if rs("保护")=true then%>
      练功保护 
      <%else%>
      没有保护 
      <%end if%>
      </font></td>
  </tr>
  <%
sj=DateDiff("n",rs("暴豆时间"),now())
if sj<=60 then%>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td colspan="4" class="p9"><div align="center"> <font size="2" color="#000000">威力还剩:</font><font size="2" color=red><b><%=60-sj%>分</b></font></div></td>
  </tr>
  <%end if%>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%"><font size="2" color="#000000">金卡:</font></td>
    <td width="26%" height="20"><font color="#0000FF"><font color=red size="2"><%=rs("会员金卡")%>元 
      </font></font></td>
    <td width="13%" height="20"><font size="2" color="#000000">金币:</font></td>
    <td width="26%" height="20"><font color="#0000FF"><font color=red size="2"><%=rs("金币")%>个</font></font></td>
  </tr>
  <tr  bgcolor=""onmouseover="this.bgColor='#DFEFFF';" onmouseout="this.bgColor='';"> 
    <td><font size="2" color="#000000">名号:</font></td>
    <td height="24"><font color="#0000FF"><font size="2"> 
      <%
if rs("等级")<5  then response.write("初来乍道")
if rs("等级")>=5 and rs("等级")<10  then response.write("江湖初行")
if rs("等级")>=10 and rs("等级")<15  then response.write("小有成就")
if rs("等级")>=15 and rs("等级")<20  then response.write("声名显赫")
if rs("等级")>=20 and rs("等级")<35  then response.write("行闯江湖")
if rs("等级")>=35 and rs("等级")<45  then response.write("一代大侠")
if rs("等级")>=45 and rs("等级")<55  then response.write("江湖剑客")
if rs("等级")>=55 and rs("等级")<65  then response.write("闻名天下")
if rs("等级")>=65 and rs("等级")<75  then response.write("云游仙胜")
if rs("等级")>=75 and rs("等级")<80  then response.write("已入仙道")
if rs("等级")>=80 and rs("等级")<85  then response.write("小仙")
if rs("等级")>=85 and rs("等级")<90  then response.write("大仙")
if rs("等级")>=90 and rs("等级")<95  then response.write("极帝大仙")
if rs("等级")>=95 and rs("等级")<100  then response.write("仙人")
if rs("等级")>=100 then response.write("剑仙")%>
      </font></font></td>
    <td height="24"><font color="#000000" size="2">身份:</font></td>
    <td height="24"><font color="#0000FF"><font size="2"><%=rs("身份")%></font><font color=red size="2"> 
      </font></font></td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%" class="p9" height="12"><font size="2" color="#000000">攻击: 
      </font></td>
    <td width="26%" height="12"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("攻击")%></font></font> 
    </td>
    <td width="13%" height="12"><font size="2" color="#000000">防御: </font></td>
    <td width="26%" height="12"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("防御")%></font> 
      </font></td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%"><font size="2" color="#000000">道德: </font></td>
    <td width="26%" height="24"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("道德")%></font> 
      </font></td>
    <td width="13%" height="24"><font size="2" color="#000000">知质: </font></td>
    <td width="26%" height="24"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("知质")%></font> 
      </font></td>
  </tr>
  <tr  bgcolor="" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#DFEFFF';"> 
    <td width="13%"><font size="2" color="#000000">魅力: </font></td>
    <td width="26%" height="24"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("魅力")%></font> 
      </font></td>
    <td width="13%" height="24"><font size="2" color="#000000">神灵:</font></td>
    <td width="26%" height="24"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("sl")%></font></font></td>
  </tr>
  <tr  bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="13%"><font size="2" color="#000000">配偶: </font></td>
    <td width="26%" height="24"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("配偶")%></font> 
      </font></td>
    <td width="13%" height="24"><font size="2" color="#000000">情人:</font></td>
    <td width="26%" height="24"><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("情人")%></font></font></td>
  </tr>
  <%if rs("配偶")<>"无" then%>
  <tr  bgcolor="" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#DFEFFF';"> 
    <td colspan="4" class="p9"> 
      <div align="center"><font size="2" color="#000000">婚恋:</font><font color="#0000FF"><font color="#0000FF" size="2"><%=rs("结婚次数")%>次</font></font><font color="#0000FF" size="2"> 
        </font><font size="2" color="#000000">记念:<%=rs("结婚记念日")%> 结婚时间:<%=DateDiff("d",rs("结婚记念日"),date())%>天</font></div>
    </td>
  </tr>
  <%end if%>
  <tr> 
    <td colspan="4"> <div align="center"><font size="2" color=red><b><%if sjjh_sx=1 then%>上限打开<%else%>上限关闭<%end if%></b></font><br><font size="2" color="#000000">体力<br>
        </font><img height=12
<%if int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*200)<200 then
abc=int(int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*200)/50)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*200)%>><img height=12 src="img/5.gif" width=<%=200-int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*200)%>> 
        <%else%>
        src="img/4.gif" width=200> 
        <%end if%>
        <br>
        <font size="2"><%=rs("体力")%><font color="#FF0000"><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_tlsx+5260+tlj%><%end if%></font></font></div></td>
  </tr>
  <tr> 
    <td colspan="4" class="p9" height="23"> <div align="center"><font size="2" color="#000000">内力<br>
        </font><img height=12
<%if int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+nlj))*200)<200 then
abc=int(int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+nlj))*200)/50)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+nlj))*200)%>><img height=12 src="img/5.gif" width=<%=200-int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+nlj))*200)%>> 
        <%else%>
        src="img/4.gif" width=200> 
        <%end if%>
        <br>
        <font size="2"><%=rs("内力")%><font color="#FF0000"><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_nlsx+2000+nlj%><%end if%></font></font></div></td>
  </tr>
  <tr> 
    <td colspan="4" class="p9"> <div align="center"><font size="2" color="#000000">武功<br>
        </font><img height=12
<%if int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+wgj))*200)<200 then
abc=int(int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+wgj))*200)/50)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+wgj))*200)%>><img height=12 src="img/5.gif" width=<%=200-int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+wgj))*200)%>> 
        <%else%>
        src="img/4.gif" width=200> 
        <%end if%>
        <br>
        <font size="2"><%=rs("武功")%><font color="#FF0000"><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_wgsx+3800+wgj%><%end if%></font></font></div></td>
  </tr>
</table>
<%rs.close
set rs=nothing
conn.close
%>