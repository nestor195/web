
<SCRIPT type="text/javascript">
var win= null;
function NewWindow(mypage,myname,w,h,scroll){
  var winl = (screen.width-w)/2;
  var wint = (screen.height-h)/2;
  var settings  ='height='+h+',';
      settings +='width='+w+',';
      settings +='top='+wint+',';
      settings +='left='+winl+',';
      settings +='scrollbars='+scroll+',';
      settings +='resizable=yes';
  win=window.open(mypage,myname,settings);
  if(parseInt(navigator.appVersion) >= 4){win.window.focus();}
}
</SCRIPT>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>My Websites Events</title>
<style>
.TableMonthHeader{font-family:Tahoma;font-size:9pt;font-weight:bold;color:black}
.MonthHeadings{width:14%;text-align:center;font-size:9pt;font-family:Tahoma;background-color:#525252;border-top:1px solid #757575;border-bottom:1px solid #757575;color:white;font-weight:bold;filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr='#848485', EndColorStr='#525252')}
.TableMonthCalendar{font-family:Tahoma;height:95%;padding:0;background-color:white;border-collapse:collapse;border-style:none;border-color:#757575;}
.TableMonthDayCellToday{font-family:Tahoma;border-style:solid;border-width:1;border-color:#757575;text-align:left;vertical-align:top;background-color:#CCCCCC;}
.TableMonthDayCell{font-family:Tahoma;border-style:solid;border-width:1;border-color:#757575;text-align:left;vertical-align:top;background-color:#FFFFFF;border-collapse:collapse;}
.MonthSubHeadings{font-family:Tahoma;font-size:8pt;background-color:#CFCFCF;color:black;font-weight:normal;filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr='#EFEFEF', EndColorStr='#CFCFCF')}
.TableMonthOtherDayCell{font-family:Tahoma;border-style:solid;border-width:1;border-color:#757575;text-align:center;vertical-align:top;background-color:#CFCFCF;border-collapse:collapse;}
.EventTable{font-family:Tahoma;border-style:solid;border-width:1;border-color:black;border-collapse:collapse;border-width:1;text-align:left;background-color:white;padding:1;width:100%;}
.EventTitleFont{font-family:Tahoma;font-size:7pt;}
.EventTimeCell{font-family:Tahoma;font-size:7pt;width:10%;text-align:left;background-color:#DDDDDD;}
.EventTimeFont{font-family:Tahoma;font-size:7pt;}
.EventTitleCell{font-family:Tahoma;font-size:7pt;width:90%;text-align:left;background-color:white;}
.EventTitleFont{font-family:Tahoma;font-size:7pt;}
.EventTitleCellAllDay{font-family:Tahoma;font-size:7pt;text-align:center}
.MonthDayDiv{width:100%;height:85%;overflow:visible;}
.MiniHeadingBar{background-color:#848485;height:19px;text-align:center;border-top:1px solid #757575;border-bottom:1px solid #757575;font-family:Tahoma;font-size:8pt;color:black;font-weight:bold;filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr='#EFEFEF', EndColorStr='#CFCFCF')}
.TableMiniHeader{height:1;padding:0;background-color:white;border-style:solid;border-color:#757575;border-width:0;border-collapse:collapse;}
.MiniCalHeading{width:14%;font-family:Tahoma;font-size:8pt;color:black;font-weight:normal;background-color:#EFEFEF;text-align:center;}
.TableMiniCalendar{padding:0;background-color:white;border-collapse:collapse;border-width:0;border-style:none;}
.TableMiniDayCellToday{border-style:solid;border-width:1;border-color:white;text-align:center;vertical-align:top;background-color:silver;padding:0;}
.TableMiniDayCell{border-style:solid;border-width:1;border-color:white;text-align:center;vertical-align:center;background-color:white;padding:0;border-collapse:collapse;cursor:hand;}
.TableMiniDayCellWithEvent{border-style:solid;border-width:1;border-color:white;text-align:center;vertical-align:center;background-color:#FBE694;padding:0;border-collapse:collapse;cursor:hand}
.FontCalendarDay{font-family:Tahoma;font-size:7pt;}
.TableMiniOtherDayCell{border-style:solid;border-width:1;border-color:white;text-align:center;vertical-align:center;background-color:#CFCFCF;padding:0;border-collapse:collapse;}
.EventLeftTD{width:20%;font-family:Tahoma;font-size:8pt;font-weight:bold;background-color:#EFEFEF;}
.EventRightTD{width:80%;font-family:Tahoma;font-size:8pt;}
.EventTitleBar{background-color:#CFCFCF;height:19px;text-align:left;border-top:1px solid #757575;border-bottom:1px solid #757575;font-family:Tahoma;font-size:10pt;color:black;font-weight:bold;filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr='#848485', EndColorStr='#CFCFCF')}
.ButtonBar{background-color:#EFEFEF;padding-top:1px;width:100%;height:30px;filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr='#EFEFEF', EndColorStr='#CFCFCF')}
.Button{background-color:#EFEFEF;cursor:hand;padding:1px 1px 1px 1px;height:27px;filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr='#EFEFEF', EndColorStr='#CFCFCF')}
.ButtonOver{background-color:#CFCFCF;cursor:hand;border: 1px solid #757575;height:27px;filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr='#FBE694', EndColorStr='#EE9515')}
.ButtonFont{font-family:Tahoma;font-size:9pt;font-weight:bold;}
.PageBody{background-color:#525252;filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr='#848485', EndColorStr='#525252')}
.SideBar{background-color:#525252;filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr='#848485', EndColorStr='#525252')}
.StandardFont{font-family:Tahoma;font-size:8pt;color:black;font-weight:bold;}
.StandardTextBox{font-family:Tahoma;font-size:8pt;color:black;font-weight:normal;width:100%;}
.DescriptionHeadingFont{font-family:Tahoma;font-size:13pt;color:yellow;font-weight:bold;}
.DescriptionFont{font-family:Tahoma;font-size:8pt;color:white;font-weight:normal;}
.EditPaneTable{width:99%; border:0px;}
.EditPaneLeft{width:25%;font-family:Arial;font-size:8pt;}
.EditPaneRight{width:75%;font-family:Arial;font-size:8pt;}
.ErrorFont{font-family:Tahoma;font-size:8pt;color:red;font-weight:bold;}

</style>
</head>

<body>
<iframe src="https://www.google.com/calendar/embed?src=gpc3vbkg86854mmsq7svpu8ni8%40group.calendar.google.com&ctz=America/Argentina/Buenos_Aires" style="border: 0" width="800" height="600" frameborder="0" scrolling="no"></iframe>
</body>

</html>
