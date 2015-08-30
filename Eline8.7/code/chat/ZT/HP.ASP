<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../config.asp"-->
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
mygj=rs("等级")*sjjh_gjsx
myfy=rs("等级")*sjjh_fysx
'if rs("攻击")>mygj then
	conn.execute("update 用户 set 攻击="& mygj &" where 姓名='"&sjjh_name&"'")
'end if
'if rs("防御")>myfy then
	conn.execute("update 用户 set 防御="& myfy &" where  姓名='"&sjjh_name&"'")
'end if
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
<title>♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
<!--
td{font-size:9pt;}
a:link {color:#0000FF;}
a:visited {color:0000ff;}
a:hover{color:#FFFFFF;; cursor: hand}
.hand{background-color:rgb(208,207,192);cursor:hand;}
}-->
</style>

</head>
<body leftmargin="0" topmargin="2" style="color: #FFFFFF" oncontextmenu=self.event.returnValue=false>
<DIV ID="overDiv" STYLE="position:absolute; visibility:hide; z-index: 1;"></DIV>
<SCRIPT LANGUAGE="JavaScript" src="img/overlib.js"></SCRIPT>
<table width="255" border="0" cellspacing="0" cellpadding="0"  height="90">
             <td width="255" height="90">
             <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" class="qx" background="img/zt_r31_c1.jpg" height="99">
                <tr>
                 <td width="100%" height="18">
                        <p align="left">
                        <font color="#FFFFFF">&nbsp;</font>				                

                        </p>

					</td>

                </tr>
                <tr>
                 <td width="75%" height="21">
                        <font color="#FFFFFF">&nbsp; 气血：<img src="img/qx1.gif" height="12" height="8" >         

                        <img <%if int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*200)<200 then %>src="img/qx.gif" name=qx_pic width="<%=int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*180)%>" height="8" onMouseOver="drs('体力<%=rs("体力")%><%if sjjh_sx=1 then%>/上限<%=rs("等级")*sjjh_tlsx+5260+tlj%><%end if%>'); return true;" onMouseOut="nd(); return true;"><img height="8" src="img/sx.gif" width=<%=180-int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*180)%> onMouseOver="drs('体力<%=rs("体力")%><%if sjjh_sx=1 then%>/上限<%=rs("等级")*sjjh_tlsx+5260+tlj%><%end if%>'); return true;" onMouseOut="nd(); return true;">				                                
                         <%else%>
                         src="img/qx.gif" width=180 onMouseOver="drs('体力<%=rs("体力")%><%if sjjh_sx=1 then%>/上限<%=rs("等级")*sjjh_tlsx+5260+tlj%><%end if%>'); return true;" onMouseOut="nd(); return true;" > 
                         <%end if%>
                        </font>				                

					</td>

                </tr>
	            <tr>
					<td width="75%" height="21">
                    <font color="#FFFFFF">&nbsp;&nbsp;内力：<img src="img/nl1.gif" height="12" >    
                        <img <%if int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+nlj))*200)<200 then %>src="img/nl.gif" name=nl_pic width="<%=int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+nlj))*180)%>>" height="8" onMouseOver="drs('内力<%=rs("内力")%><%if sjjh_sx=1 then%>/上限<%=rs("等级")*sjjh_nlsx+2000+nlj%><%end if%>'); return true;" onMouseOut="nd(); return true;"><img height="8" src="img/sx.gif" width=<%=180-int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+nlj))*180)%> onMouseOver="drs('内力<%=rs("内力")%><%if sjjh_sx=1 then%>/上限<%=rs("等级")*sjjh_nlsx+2000+nlj%><%end if%>'); return true;" onMouseOut="nd(); return true;">              
                        <%else%>
                        src="img/nl.gif" width=180 onMouseOver="drs('内力<%=rs("内力")%><%if sjjh_sx=1 then%>/上限<%=rs("等级")*sjjh_nlsx+2000+nlj%><%end if%>'); return true;" onMouseOut="nd(); return true;" > 
                        <%end if%>
                    </font>      
					</td>
                </tr>
	            <tr>
					<td width="75%" height="19">
                    <font color="#FFFFFF">&nbsp; 精力：<img src="img/jl1.gif" height="12">
	                    <img <%if int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+wgj))*200)<200 then %>src="img/jl.gif" name=jl_pic width="<%=int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+wgj))*180)%>" height="8" onMouseOver="drs('武功<%=rs("武功")%><%if sjjh_sx=1 then%>/上限<%=rs("等级")*sjjh_wgsx+3800+wgj%><%end if%>'); return true;" onMouseOut="nd(); return true;" ><img height="8" src="img/sx.gif" width=<%=180-int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+wgj))*180)%> onMouseOver="drs('武功<%=rs("武功")%><%if sjjh_sx=1 then%>/上限<%=rs("等级")*sjjh_wgsx+3800+wgj%><%end if%>'); return true;" onMouseOut="nd(); return true;">              
                    <%else%>
                    src="img/jl.gif" width=180 onMouseOver="drs('武功<%=rs("武功")%><%if sjjh_sx=1 then%>/上限<%=rs("等级")*sjjh_wgsx+3800+wgj%><%end if%>'); return true;" onMouseOut="nd(); return true;"> 
                   <%end if%>
                    </font>      
					</td>
                </tr>
			  	<tr>
					<td width="75%" height="20">
					</td>
				</tr>
			  </table>
</table>
</body>
</html>
<%rs.close
set rs=nothing
conn.close
%>

