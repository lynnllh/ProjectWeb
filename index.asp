<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/login.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_login_STRING
Recordset1_cmd.CommandText = "SELECT * FROM teamactivity ORDER BY moddate DESC" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 8
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<%
Dim MM_paramName 
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<!doctype html>
<html><!-- InstanceBegin template="/Templates/ProjectWeb.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<meta charset="utf-8">
<!-- InstanceBeginEditable name="doctitle" -->
<title>课题组</title>
<meta name="keywords" content="课题组">
<script src="flexslider/js/jquery.js"></script>
<script src="flexslider/js/jquery.flexslider-min.js"></script>
<link rel="stylesheet"  type="text/css"  href="flexslider/css/flexslider.css">
<script type="text/javascript">
		$(function() {
			$(".flexslider").flexslider({
				slideshowSpeed: 4000, //展示时间间隔ms
				animationSpeed: 400, //滚动时间ms
				touch: true, //是否支持触屏滑动
				animation: "slide",
				pauseOnHover: true
			});
		});	
</script>
<link href="index.css" rel="stylesheet" type="text/css">
<!-- InstanceEndEditable -->
<style type="text/css">
body {
	background-color: #D5DADE;
}
</style>
<link href="main.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="head" -->
<!-- InstanceEndEditable -->
</head>

<body>
<div class="webbody">
  <div class="head">
    <p><strong><img src="indeximages/ground.jpg" width="900" height="204"></strong></p>
    <div class="navbar">
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <th width="128" height="50" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="index.asp" class="navword">首页</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="researchteam/researchteamprofessor.asp" class="navword">团队成员</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="researchdirection/researchdirection.asp" class="navword">研究方向</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="researchresult/researchresult.asp" class="navword">论文成果</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="teamactivity/teamactivity.asp" class="navword">团队动态</a></th>
          <th scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="contactus/contactus.asp" class="navword">联系我们</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a  href="login/backsystemlogin.asp" class="navword">管理入口</a></th>
        </tr>
      </table>
      <br></div>  
      <div style="height:50">
    <p></p>
    </div>
    <!-- InstanceBeginEditable name="EditRegion" -->
    <div class="leaderdiv">
<table width="220" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <th colspan="2" scope="col" class="leadertable" align="left" valign="middle"><img src="indeximages/箭头.jpg" 

width="20" height="17" style="vertical-align:middle">&nbsp;&nbsp;学术带头人</th>
  </tr>
  <tr>
  <td height="5" colspan="2">
  </td>
  </tr>
  <tr>
    <td rowspan="3"  class="leaderphoto"><img src="indeximages/教授.jpg" width="100" height="120"></td>
    <td  class="leaderrank">张三教授</td>
  </tr>
  <tr>
    <td class="leaderrank">机械电子系系主任</td>
  </tr>
  <tr>
    <td  class="leaderrank">智能机器人所所长</td>
  </tr>
    <tr>
  <td height="5" colspan="2">
  </td>
  </tr>
  <tr>
    <td colspan="2" class="leaderintroductionword"> &nbsp; 黄继光（1931年1月8日—1952年10月19日），民族英雄。1931年生于四川省中江县，中国人民志愿军第45师135团9连的通讯员。1952年10月19日在朝鲜上甘岭地区597.9高地牺牲,年仅21岁。被中国人民志愿军领导机关追记特等功，并授予“特级英雄”称号[1]  ；所在部队党委追授他为中国共产党正式党员；朝鲜民主主义人民共和国最高人民会议常务委员会授予他“朝鲜民主主义人民共和国英雄”称号和金星奖章和一级国旗勋章。</td>
  </tr>
</table>
</div>
  	<div class="indexmiendispaly">
      <div class="flexslider" style="width:663px; height:250px; border:#FFF; padding-left:4px; padding-top:4px;"> 
      <ul class="slides"> 
						<li><img src="images/s1.jpg" height="291"/></li>
						<li><img src="images/s2.jpg" height="291"/></li>
						<li><img src="images/s3.jpg" height="291"/></li>
						<li><img src="images/s4.jpg" height="291"/></li>
                        <li><img src="images/sc1.jpg" height="291"/></li>
						<li><img src="images/sc2.jpg" height="291"/></li>
						<li><img src="images/sc3.jpg" height="291"/></li>
						<li><img src="images/sc4.jpg" height="291"/></li>
      </ul> 
</div> 
 <div id="indexpaper">
      <table width="470" border="0" cellspacing="0" cellpadding="0">
		<tr>
        <td  height="30" style="background-color:#61a3ff" class="friendlinkword"><img src="indeximages/箭头.jpg" width="20" height="17" style="vertical-align:middle">&nbsp;组内新闻</td>
        <td width="60" align="left" style="background-color:#61a3ff" ><a href="teamactivity/teamactivity.asp" class="message"  onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">更多内容</a></td>
        </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
  <tr>
    <td width="410" height="20">&nbsp;&nbsp;<A HREF="teamactivity/teamactivitydetail.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "ID=" & Recordset1.Fields.Item("ID").Value %>" class="message" onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'"><%=(Recordset1.Fields.Item("title").Value)%></A></td>
    <td class="message"><%=(Recordset1.Fields.Item("moddate").Value)%></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
      </table> 
      </div>
	 <div class="linkfriendlinkk">   

     	<table width="190" align="left" cellpadding="0" cellspacing="0">
          <tr align="left">
    <th scope="col" height="30" style="background-color:#61a3ff" class="friendlinkword" align="left"><img src="indeximages/箭头.jpg" width="20" height="17" style="vertical-align:middle">&nbsp;友情链接</th>
  </tr>
  <tr>
    <td  height="20">&nbsp;<a href="http://www.baidu.com/" class="message"  onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">百度</a></td>
  </tr>
  <tr>
    <td  height="20">&nbsp;<a href="http://www.baidu.com/" class="message"  onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">百度</a></td>
  </tr>
  <tr>
    <td  height="20">&nbsp;<a href="http://www.baidu.com/" class="message"  onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">百度</a></td>
  </tr>
  <tr>
    <td  height="20">&nbsp;<a href="http://www.baidu.com/" class="message"  onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">百度</a></td>
  </tr>
  <tr>
    <td  height="20">&nbsp;<a href="http://www.baidu.com/" class="message"  onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">百度</a></td>
  </tr>
  <tr>
    <td  height="20">&nbsp;<a href="http://www.baidu.com/" class="message"  onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">百度</a></td>
  </tr>   
  <tr>
    <td  height="20">&nbsp;<a href="http://www.baidu.com/" class="message"  onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">百度</a></td>
  </tr>
  <tr>
    <td  height="20">&nbsp;<a href="http://www.baidu.com/" class="message"  onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">百度</a></td>
  </tr>

</table>
      </div>
	
      <div id="footer"><table width="570" border="0" cellspacing="5" cellpadding="2" align="center">
  <tr>
    <th  align="center" id="footerwords">&copy;&nbsp;Copyright&nbsp;2007-2014&nbsp;百度百度百度课题组版权所有

</th>
  </tr>
  <tr>
    <td align="center" id="footerwords">北京市西长安街174号中南海新华门|邮编：100017</td>
  </tr>
  <tr>
    <td align="center" id="footerwords">电话：010-11111111|邮箱：xiaobai@163.com|推荐兼容模式打开</td>
  </tr>
</table>
</div>

	
	
	<!-- InstanceEndEditable --></div>

  
    
   
  
</div>
</body>
<!-- InstanceEnd --></html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
