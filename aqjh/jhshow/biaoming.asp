<!--#include file="../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
if Application("jhshowurl")<>"http://qqshow-item.tencent.com" then
	showgif=Application("jhshowurl") & "gif/pic/[img].gif"
	showgif1="src="& Application("jhshowurl") & "gif/'+j+'/'+showArray[i]+'.gif"
else
	showgif=Application("jhshowurl") & "/[img]/0/00/00.gif"
	showgif1="src="& Application("jhshowurl") &"/'+showArray[i]+'/'+j+'/00/00.gif"
end if
aqjh_name=Session("aqjh_name")
ndate=Day(date())
if ndate>20 then
		Response.Write "<script Language=Javascript>alert('��ʾ���������������ֻ��ÿ��ǰ20�죬����ʱ���ѹ��������ٲμӣ�');window.close();</script>"
		response.end
end if
myok=trim(lcase(cstr(server.HTMLEncode(Request("ok")))))
%>
<html><title>�������������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet></head>
<body leftmargin="0" topmargin="0">
<%if myok="ok" then
	sm=trim(lcase(cstr(server.HTMLEncode(Request.form("sm")))))
	sm=replace(sm,"&#","")
	sm=Replace(sm,"&amp;","&")
	sm=Replace(sm,"&amp;","&")
	sm=Replace(sm,"'","&#" & asc("'"))
	sm=Replace(sm,"\","")
	sm=Replace(sm,",","&#" & asc(","))
	sm=Replace(sm,"""","&#" & asc(""""))
	sm=Replace(sm,"<","&#" & asc("<"))
	sm=Replace(sm,">","&#" & asc(">"))
	sm=replace(sm,"..","")
	sm=Replace(sm,"\x3c","")
	sm=Replace(sm,"\x3e","")
	sm=Replace(sm,"\074","")
	sm=Replace(sm,"\74","")
	sm=Replace(sm,"\75","")
	sm=Replace(sm,"\76","")
	sm=Replace(sm,"&lt","")
	sm=Replace(sm,"&gt","")
	sm=Replace(sm,"\076","")
	if len(sm)<3 then
		Response.Write "<script Language=Javascript>alert('��ʾ��˵��������С��3�����֣�');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
	rs.open "select ��� from [�û�] where ����='" & Session("aqjh_name") & "'",conn,1,3
	if rs("���")<10 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	 	Response.Write "<script language=JavaScript>{alert('��ʾ����û��10����ң����ܽ��б���������');location.href = 'javascript:history.go(-1)';}</script>"
		response.end
	end if
	rs("���")=rs("���")-10
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("showmdb")
	rs.Open "select * from use where a='"& Session("aqjh_name") &"'", conn, 1,3
	if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ���㻹û��ע�Ὥ���㣬���ܹ��μӱ�����');window.close();</script>"
		response.end
	end if
	if rs("f")="||||||14|13|12||11||10|9||||8|||||||" or rs("f")="||||||7|6|5||4||3|2||||1|||||||" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ��������ͼΪ��ʼ��ͼ����������װ�κ�������');window.close();</script>"
		response.end
	end if
	myshow=rs("f")
	if rs("c")="M" then
		sex="��"
	else
		sex="Ů"
	end if
	rs.close
	newbao=0
	rs.Open "select * from sai where s_����='"& Session("aqjh_name") &"'", conn, 1,3
	if rs.eof and rs.bof then
		newbao=1
		rs.AddNew
	end if
	rs("s_����")=Session("aqjh_name")
	rs("s_����")=myshow
	rs("s_Ʊ��")=0
	rs("s_����ʱ��")=now()
	rs("s_˵��")=sm
	rs("s_�Ա�")=sex
	rs.update
	if newbao=1 then
		conn.execute "update pool set p_���=p_���+10,p_��������=p_��������+1,p_�е�һ��='',p_�еڶ���='',p_Ů��һ��='',p_Ů�ڶ���='',p_�е�һ�콱=false,p_�еڶ��콱=false,p_Ů��һ�콱=false,p_Ů�ڶ��콱=false"
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	says="<font color=#ff0000>����Ϣ��</font>["&session("aqjh_name")&"]<font color=green>�μ��˽�����ѡ���������˴�������21�ս��й���ͶƱ����ͶƱ����<font color=blue><b>"& (21-ndate) &"</b></font>�죡<font class=t>(" & time() & ")</font></font>"
	call showchat(says)
	Response.Write "<script Language=Javascript>alert('��ʾ�������ɹ�������ص�������ҳ��ˢ�¼��ɿ�����ı������ϣ�');window.close();</script>"
	response.end
else%>
<table width="100%" height="100%" border="0" align="center" bgcolor="#FF9900">
     <form name="form" method="post" action="biaoming.asp?ok=ok">
     <tr> 
      <td height="83" align="center"> 
        <%Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("showmdb")
	rs.Open "select * from use where a='"& Session("aqjh_name") &"'", conn, 1,3
	if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ���㻹û��ע�Ὥ���㣬���ܹ��μӱ�����');window.close();</script>"
		response.end
	end if
	if rs("f")="||||||14|13|12||11||10|9||||8|||||||" or rs("f")="||||||7|6|5||4||3|2||||1|||||||" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ��������ͼΪ��ʼ��ͼ����������װ�κ�������');window.close();</script>"
		response.end
	end if
	Response.Write "<script language=JavaScript>var jhshow='"& rs("f") &"';"
	Response.Write "document.write ("& chr(34) &"<div id=show"& myrnd &" style='PADDING-RIGHT: 0px; PADDING-LEFT: 0px; LEFT: 0px; PADDING-BOTTOM: 0px; WIDTH: 140px; PADDING-TOP: 0px; POSITION: relative; TOP: 0px; HEIGHT: 226px;'></div>"& chr(34) &");"
	Response.Write "var showArray = jhshow.split('|');"
	Response.Write "var s='';"
	Response.Write "for (var i=0; i<25; i++){if(showArray[i] != ''){j = i + 1;"
	Response.Write "s+='<IMG id=s'+j+' "& showgif1 &" style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:'+j+';"& chr(34) &">';"
	Response.Write "}}"
	Response.Write "s+='<IMG src=face/blank.gif style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:50;"& chr(34) &">';"
	Response.Write "show"& myrnd &".innerHTML=s;</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
      </td>
  </tr>
  <tr> 
      <td><font color="#FFFF00">ȷ����ʹ�õ�ǰ������б�����������ϵͳ���۳���10�������Ϊ�������ã�һ���������ܳ����������Ҫ�޸�������21��ǰ��������±��������±�����ʹ�������񣬵�ϵͳ��ȻҪ�ٿ۳�10����ң������������ؿ��ǡ�</font></td>
  </tr>
  <tr>
    <td align="center">����˵��(40����)�� 
      <input name="sm" type="text" size="20" maxlength="40"><br><br> 
      <input type="submit" name="Submit" value="ȷ����Ҫ����">
    </td>
  </tr></form>
</table>
<%end if%>
</body></html>