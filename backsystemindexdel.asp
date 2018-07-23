<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/login.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_login_STRING
    MM_editCmd.CommandText = "DELETE FROM message WHERE ID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "backsystemindex.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If

End If
%>
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
<script type="text/javascript">
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
</script>
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
    <th class="currentptable" scope="col"><a href="backsystemindex.asp" class="dellink" onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">首页</a></th>
  </tr>
</table>
</div>
<div id="condiv"><form name="form1" action="<%=MM_editAction%>" method="POST">


<table border="0" width="670">
	<tr>
    <th>&nbsp;
    </th>
    </tr>
    	<tr>
    <th>&nbsp;
    </th>
    </tr>
    	<tr>
    <th>&nbsp;
    </th>
    </tr>
    	<tr>
    <th>&nbsp;
    </th>
    </tr>
	<tr>
    <th>确定要删除吗？
    </th>
    </tr>
    <tr align="center">
    <td><input name="提交" type="submit" id="提交" value="提交"><input name="返回" type="button" id="返回" onClick="MM_goToURL('parent','backsystemindex.asp');return document.MM_returnValue" value="返回">
    </td>
    </tr>
    </table>
<input type="hidden" name="MM_delete" value="form1">
<input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("ID").Value %>">
</form>

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
