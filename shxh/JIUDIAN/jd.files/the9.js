MSIEIndex = navigator.userAgent.indexOf("MSIE");
ns = ((MSIEIndex) == -1);

function logout()
{
        if (confirm('您确实要离开第九城市?'))
        {
                window.top.location.href = "http://www.the9.com/logout.php";
        }
}

function demit(demit)
{
        if (confirm('您确实要辞职吗？\n\n先前的辛勤劳动可会付诸东流噢！'))
        {
           var url = "/work/demit.php?demit="+demit;
		   window.self.location.href = url;
        }
}

function ViewInfo(name)
{
        if (arguments.length == 0)
        {
                return;
        }
        if (name == "")
        {
                return;
        }
        //modify by emilyzhou on 2000/7/17 
		if (ns)
        {
                name = escape(name);
		}
        var i;
        var newname = "";
        for (i=0; i<name.length; i++)
        {
                if (name.charAt(i) == '+')
                {
                        newname = newname + "%2B";
                }
                else
                {
                        newname = newname + name.charAt(i);
                }
        }
		//var infourl = "http://www.the9.com/resident/ctl_resident.php?action=detail&name=" + newname;
       	var infourl = "http://www.the9.com/resident/detail.php?name=" + newname;
		window.open(infourl, 'infowin', 'left=120, top=100, width=540,height=450,resizable=1,status=0,menubar=0,scrollbars=1');
}

function Visit(name)
{
   var infourl = "/home/livingroom/visit.php?host="+name;
   window.open(infourl, 'visit', 'left=10, top=10, width=760,height=450,resizable=0,status=0,menubar=0,scrollbars=0');
}

function ViewHelp(name)
{
        if (arguments.length == 0)
        {
                name = "";
        }
        if (ns)
        {
                name = escape(name);
        }
        if (name == "")
        {
                var url = escape(window.location);
                //var helpurl = "http://www.the9.com/help/index.php?url=" + url;
				var helpurl = "http://www.the9.com/help/ctl_help.php?url=" + url;
        }
        else
        {
                //var helpurl = "http://www.the9.com/help/index.php?kind=" + name;
				var helpurl = "http://www.the9.com/help/ctl_help.php?kind=" + name;
        }
        window.open(helpurl, 'helpwin', 'left=100, top=80, width=540,height=450,resizable=1,status=0,menubar=0,scrollbars=1');
}

function CheckMail()
{
        //var newurl = "http://www.the9.com/message/show.php";
		var newurl = "http://www.the9.com/message/ctl_message.php?action=show";
        window.open(newurl, 'mailwin', 'left=50, top=50, width=700,height=450,resizable=1,status=0,menubar=0,scrollbars=1');
}

function WriteMailReply(name,serial)
{
        //var newurl = "http://www.the9.com/message/message.php";
		var newurl = "http://www.the9.com/message/ctl_message.php?action=send";
        if (arguments.length != 0)
        {
                if (ns)
                {
                        name = escape(name);
                }
                newurl = newurl + "&replyto=" + name + "&serial=" + serial;
        }
        window.open(newurl, 'writereplywin', 'left=80, top=50, width=700,height=460,resizable=1,status=0,menubar=0,scrollbars=1');
}

function WriteMail(name)
{
        //var newurl = "http://www.the9.com/message/message.php";
		var newurl = "http://www.the9.com/message/ctl_message.php?action=send";
        if (arguments.length != 0)
        {
                if (ns)
                {
                        name = escape(name);
                }
                var i;
                var newname = "";
                for (i=0; i<name.length; i++)
                {
                        if (name.charAt(i) == '+')
                        {
                                newname = newname + "%2B";
                        }
                        else
                        {
                                newname = newname + name.charAt(i);
                        }
                }

                newurl = newurl + "&replyto=" + newname;
        }
        window.open(newurl, 'writewin', 'left=80, top=50, width=700,height=460,resizable=1,status=0,menubar=0,scrollbars=1');
}

function OpenNews()
{
        window.open("", 'newswin', 'left=150, top=50, width=600,height=450,resizable=1,status=0,menubar=0,scrollbars=1');
}

function ViewGuid()
{
        guidurl = "http://www.the9.com/other/navigation.php";
		window.open(guidurl, 'guidwin', 'left=150, top=50, width=540,height=400,resizable=1,status=0,menubar=0,scrollbars=1');
}


function ViewTrInfo(tr_id)
{
        if (arguments.length == 0)
        {
                return;
        }
        if (tr_id == "")
        {
                return;
        }
        var infourl = "/treasure/tr_info.php?tr_id=" + tr_id;
        window.open(infourl, 'infowin', 'left=120, top=100, width=540,height=450,resizable=1,status=0,menubar=0,scrollbars=1');
}

function ViewMyTrInfo(tr_serial, tr_id)
{
        if (arguments.length == 0)
        {
                return;
        }
        if (tr_id == "")
        {
                return;
        }
        if (tr_serial == "")
        {
                return;
        }
        var infourl = "/treasure/tr_info.php?tr_id=" + tr_id + "&tr_serial=" + tr_serial;
        window.open(infourl, 'infowin', 'left=120, top=100, width=540,height=450,resizable=1,status=0,menubar=0,scrollbars=1');
}

function ViewFileInfo(fileID)
{
        if (arguments.length == 0)
        {
                return;
        }
        if (fileID == "")
        {
                return;
        }
        if (ns)
        {
                fileID = escape(fileID);
        }
        
        var infourl = "/download/dl_detail.php?dl_file_id=" + fileID;
        window.open(infourl, 'fileinfowin', 'left=120, top=100, width=540,height=450,resizable=1,status=0,menubar=0,scrollbars=1');
}

function DownloadFile(type1,type2)
{
        if (ns)
        {
                type1 = escape(type1);
                type2 = escape(type2);
        }
        
        //modify by EmilyZhou on 2001-03-15
        var infourl = "http://upload.the9.com/upload/ul_upload.php?ul_type1=" + type1 + "&ul_type2=" + type2;
        window.open(infourl, 'downloadwin', 'left=120, top=100, width=540,height=450,resizable=1,status=0,menubar=0,scrollbars=1');
        //alert("站点维护，暂时关闭上传。敬请谅解！");
}