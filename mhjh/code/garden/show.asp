<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
Sub Msg (v)
 Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><title>�ֻ�</title><meta http-equiv='pragma' content='no-cache'><style type=text/css>body{color:black;font-family:����;font-size:9pt;background-color:buttonface;border-bottom:medium none;border-left:medium none;border-right:medium none;border-top:medium none;padding-bottom:0px;padding-left:0px;padding-right:0px;padding-top:0px}</style></head><body leftMargin=0 topMargin=0 marginheight=0 marginwidth=0>"
 Response.Write "<script Language=JavaScript>alert('" & v & "');window.close();</script></body></html>"
 Response.End
End Sub
sjjh_name=session("yx8_mhjh_username")
If sjjh_name="" Then Msg "����δ��¼�������ֻ���"
actid=Request.Form("actid")
flowerid=Request.Form("flowerid")
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
select case actid
case 7
rs.open "select ����,hua from �û� where ����='"&sjjh_name&"'",conn,2,2
if rs("����")<50000 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����û�� 50000 ������������ʧ�ܣ�');parent.gg();}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.end
end if
huasj=rs("hua")

if huasj="" then
	Response.Write "<script language=JavaScript>{alert('��ʾ������û�п��ѻ�԰�����ȿ��ѣ�');parent.gg();}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.end
end if
fhuasj=split(huasj,"|")
if fhuasj(0)<=0 or fhuasj(0)="" then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㻹û�����ӣ����ȿ��ѣ�');parent.gg();}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.end
else
	seedsm=int(fhuasj(0))
	fhuasj(0)=seedsm-1
	huasj = Join(fhuasj,"|")
	rs("hua")=huasj
	rs.update
end if
huasj=rs("hua")
fhuasj=split(huasj,"|")
seedsm=fhuasj(0)
hua_sj=now()
ghuasj=seedsm&"|"&hua_sj&"|"
for i=2 to 26 step 1
huasj=split(fhuasj(i),";")
huasj_cs=huasj(0)
huasj_zl=huasj(1)
huasj_cz=huasj(2)
if i=flowerid+2 then
randomize timer
huasj_zl=int(rnd*5)
huasjnew="1;"&huasj_zl&";"&huasj_cz
Response.Write "<script language=JavaScript>{parent.go("&flowerid&",1,"&huasj_zl&",0);}</script>"
else
huasjnew=huasj_cs&";"&huasj_zl&";"&huasj_cz
end if
ghuasj=ghuasj&huasjnew&"|"
next
conn.execute "update �û� set hua='"&ghuasj&"',����=����-50000 where ����='" &sjjh_name&"'"
says="<font color=0000FF>"&sjjh_name&"</font>����50000�����ӿ�����һƬ�ĵؿ�ʼ�ֻ��ˣ���Ȼ�����½���50000�㣬���ǻ�������밮�˷������һ��ܸ��Լ����Ӵ���������õľ�������������һ�����!"		
Response.Write "<script language=JavaScript>{parent.gs("&seedsm&");parent.gg();}</script>"
says1="<font color=red><b>���ֻ���Ϣ��</b></font>"&says
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
		newtalkarr(594)=name
		newtalkarr(595)="���" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="008000" 
		newtalkarr(599)=""&says1&""
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
case 8
rs.open "select ����,hua from �û� where ����='"&sjjh_name&"'",conn,2,2
if rs("����")<50000 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����û�� 50000 ������������ʧ�ܣ�');parent.gg();}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.end
end if
conn.execute "update �û� set hua='3|"&now()&"|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|',����=����-50000 where ����='" & sjjh_name &"'"
seedsm=3
Response.Write "<script language=JavaScript>{for(var iii=0;iii<25;iii++){parent.go(iii,0,0,0);}parent.gs("&seedsm&");parent.gg();}</script>"
Response.Write "<script language=JavaScript>{alert('��ʾ���������˻�԰������3�����ӣ��ú��ֻ���');}</script>"
case 1,2,3,4
rs.open "select ����,hua from �û� where ����='"&sjjh_name&"'",conn,2,2
if rs("����")<50000 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����û�� 50000 ������������ʧ�ܣ�');parent.gg();}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

	Response.end
end if
huasj=rs("hua")

if huasj="" then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㻹û�п��ѻ�԰�����ȿ��ѣ�');}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.end
end if
fhuasj=split(huasj,"|")
seedsm=fhuasj(0)
hua_sj=fhuasj(1)
for i=2 to 26 step 1
huasj=split(fhuasj(i),";")
huasj_cs=huasj(0)
huasj_zl=huasj(1)
huasj_cz=huasj(2)
if i=flowerid+2 and huasj_cz=actid then
huasj_cs=huasj_cs+1
huasjnew=huasj_cs&";"&huasj_zl&";0"
Response.Write "<script language=JavaScript>{parent.go("&flowerid&","&huasj_cs&","&huasj_zl&",0);}</script>"
else
huasjnew=huasj_cs&";"&huasj_zl&";"&huasj_cz
end if
ghuasj=ghuasj&huasjnew&"|"
next
ghuasj=seedsm&"|"&hua_sj&"|"&ghuasj
Response.Write "<script language=JavaScript>{parent.gs("&seedsm&");parent.gg();}</script>"
conn.execute "update �û� set hua='"&ghuasj&"' where ����='" &sjjh_name&"'"
case 5
rs.open "select hua from �û� where ����='"&sjjh_name&"'",conn,2,2
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㻹û�п��ѻ�԰����ô�ջ�');}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.end
else
	huaz=rs("hua")
end if
fhuasj=split(huaz,"|")
seedsm=fhuasj(0)
hua_sj=fhuasj(1)
for i=2 to 26 step 1
huasj=split(fhuasj(i),";")
huasj_cs=huasj(0)
huasj_zl=huasj(1)
huasj_cz=huasj(2)

if i=flowerid+2 then
  If huasj_cs<>100 Then
        Response.Write "<script language=JavaScript>{parent.gs("&seedsm&");parent.gg();alert('��ʾ�������л�ô����');}</script>"
        conn.Close
        Set conn = Nothing
        Response.End

    End If
huasjnew="0;0;0"
shuazl=huasj_zl
Response.Write "<script language=JavaScript>{parent.go("&flowerid&",0,0,0);}</script>"
else
huasjnew=huasj_cs&";"&huasj_zl&";"&huasj_cz
end if
ghuasj=ghuasj&huasjnew&"|"
next
ghuasj=seedsm&"|"&hua_sj&"|"&ghuasj
if shuazl=0 then huaname="����"
if shuazl=1 then huaname="����"
if shuazl=2 then huaname="��ܿ"
if shuazl=3 then huaname="��ѩ"
if shuazl=4 then huaname="����"
if shuazl=0 then huajb=70000
if shuazl=1 then huajb=80000
if shuazl=2 then huajb=90000
if shuazl=3 then huajb=100000
if shuazl=4 then huajb=150000
conn.execute "update �û� set hua='"&ghuasj&"',����=����+"&huajb&" where ����='" &sjjh_name&"'"
Response.Write "<script language=JavaScript>{parent.gg();alert('��ʾ�����ջ������ֵĻ� ["&huaname&"] ����þ��� "&huajb&"��');}</script>"

fhuasj=split(rs("hua"),"|")
seedsm=fhuasj(0)
Response.Write "<script language=JavaScript>{parent.gs("&seedsm&");parent.gg();}</script>"
case 6
rs.open "select hua from �û� where ����='"&sjjh_name&"'",conn,2,2
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㻹û�п��ѻ�԰����ô�ջ�');}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.end
else
	huaz=rs("hua")
end if
fhuasj=split(huaz,"|")
seedsm=fhuasj(0)
hua_sj=fhuasj(1)
for i=2 to 26 step 1
huasj=split(fhuasj(i),";")
huasj_cs=huasj(0)
huasj_zl=huasj(1)
huasj_cz=huasj(2)
if i=flowerid+2 then
  If huasj_cs<>100 Then
        Response.Write "<script language=JavaScript>{parent.gs("&seedsm&");parent.gg();alert('��ʾ�������л�ô����');}</script>"
        conn.Close
        Set conn = Nothing
        Response.End

    End If
huasjnew="0;0;0"
shuazl=huasj_zl
Response.Write "<script language=JavaScript>{parent.go("&flowerid&",0,0,0);}</script>"
else
huasjnew=huasj_cs&";"&huasj_zl&";"&huasj_cz
end if
ghuasj=ghuasj&huasjnew&"|"
next
if seedsm="" then seedsm=0
seedsm=seedsm+3
ghuasj=seedsm&"|"&hua_sj&"|"&ghuasj
if shuazl=0 then huaname="����"
if shuazl=1 then huaname="����"
if shuazl=2 then huaname="��ܿ"
if shuazl=3 then huaname="��ѩ"
if shuazl=4 then huaname="����"
conn.execute "update �û� set hua='"&ghuasj&"' where ����='" &sjjh_name&"'"
Response.Write "<script language=JavaScript>{parent.gs("&seedsm&");parent.gg();alert('��ʾ�����ջ������ֵĻ� ["&huaname&"] ,����������3��');}</script>"

end select
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.end
%>