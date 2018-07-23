<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/login.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_login_STRING
Recordset1_cmd.CommandText = "SELECT * FROM message WHERE ID = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 5, 1, -1, Recordset1__MMColParam) ' adDouble

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!doctype html>
<html><!-- InstanceBegin template="/Templates/ProjectWebbacksystem.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<meta charset="utf-8">
<!-- InstanceBeginEditable name="doctitle" -->
<title>后台管理</title>
<link href="login.css" rel="stylesheet" type="text/css">
<style type="text/css">
#condiv {
	
		width: 670px;
	background-color: #FFF;
	height: 410px;
	float:right;
	margin-top: -220px;
}
#footer{
	width: 900px;
	height: 85px;
	background-color: #3388ff;
	margin-top: 200px;

	
}
</style>
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
    <p><strong><img src="indeximages/backsystem.jpg" width="900" height="204"></strong></p>
    <div class="navbar">
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <th width="114" height="50" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="backsystemindex.asp" class="navword">首页</a></th>
          <th width="114" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="researchteam/backresearchteamlist.asp" class="navword">团队成员</a></th>
          <th width="114" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="researchdirection/backresearchdirectionlist.asp" class="navword">研究方向</a></th>
          <th width="114" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="researchresult/researchresultlist.asp" class="navword">论文成果</a></th>
          <th width="114" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="teamactivity/teamactivitylist.asp" class="navword">团队动态</a></th>
          <th scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a   href="contactus/backcontactusmod.asp" class="navword">联系我们</a></th>
          <th width="114" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a  href="login/leavemessage.asp" class="navword">组内建议</a></th>
          <th width="114" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a  href="index.asp" class="navword">返回前台</a></th>
        </tr>
      </table>
      <br></div>
      <div style="height:50">
    <p></p>
    </div>
    <!-- InstanceBeginEditable name="EditRegion" -->
          <div class="steamdiv">
      <table width="220" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <th class="commonbtable">后台管理系统</th>
          </tr>
        <tr>
          <th class="commonsword" scope="col" onmouseover="this.bgColor='#3388ff'" onmouseout="this.bgColor=''" ><a href="backsystemindex.asp" class="commonsword">留言列表</a></th>
        </tr>
        <tr>
          <th class="commonsword" scope="col" onmouseover="this.bgColor='#3388ff'" onmouseout="this.bgColor=''"><a href="backsystemindexaddresslist.asp" class="commonsword">通讯录</a></th>
          </tr>

      </table>
    </div>
    <div class="currentplace"><table width="145" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <th class="currentptable" scope="col"><img src="indeximages/箭头.jpg" width="20" height="17" style="vertical-align:middle"></th>
	<th class="currentptable" scope="col">当前位置：</th>
    <th class="currentptable" scope="col"><a href="backsystemindex.asp"  class="dellink" onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'" >首页</a></th>
  </tr>
</table>
</div>
<div id="condiv">
<table width="670" border="0" cellspacing="0" cellpadding="0">
    <tr height="50">
  <td>

  </td>
  </tr>
  <tr>
    <th scope="col"  style="font-size: 18px"><%=(Recordset1.Fields.Item("title").Value)%></th>
  </tr>
    <tr height="10">
  <td>

  </td>
  </tr>
    <tr>
    <th scope="col"  class="footsong"><%=(Recordset1.Fields.Item("moddate").Value)%></th>

  </tr>
  <tr>
  <td><hr>

  </td>
  </tr>

    <tr>
    <th scope="col" align="left"  class="footsong">&nbsp;&nbsp;<%=(Recordset1.Fields.Item("content").Value)%></th>

  </tr>

</table>

</div>
	
	     <div id="footer"><table width="570" border="0" cellspacing="5" cellpadding="2" align="center">
  <tr>
    <th  align="center" id="footerwords">&copy;&nbsp;Copyright&nbsp;2007-2014&nbsp;机器人与多体系统课题组版权所有

</th>
  </tr>
  <tr>
    <td align="center" id="footerwords">南京市秦淮区御道街29号南航明故宫校区|邮编：210016</td>
  </tr>
  <tr>
    <td align="center" id="footerwords">电话：025-84892503|邮箱：chenbye@nuaa.edu.cn|推荐兼容模式打开</td>
  </tr>
</table>
</div>
	
	
	
	<!-- InstanceEndEditable -->
  </div>
  
    
   
  
</div>
</body>
<!-- InstanceEnd --></html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
