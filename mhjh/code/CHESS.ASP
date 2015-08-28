<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=session("yx8_mhjh_username")
sjjh_grade=session("yx8_mhjh_usergrade")
sjjh_jhdj=session("yx8_mhjh_usergrade")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%><td bgcolor="" align=center><font color=""><b>ÖÐ ¹ú Ïó Æå</b></font></td> 
</tr> 
<tr><td bgcolor=""> 
<CENTER> 
<APPLET code=b.class archive=1.jar codebase=http://202.96.140.39:800/game/xq WIDTH=760 HEIGHT=440 HSPACE=0 VSPACE=0> 
<param name=host value="202.96.140.39"> 
<param name=soundurl value="http://202.96.140.39:800/game/xq/sounds/";> 
<param name=port value="1220"> 
<param name="username" value="<%=sjjh_name%>"> 
</APPLET></CENTER> 
</td></tr> 