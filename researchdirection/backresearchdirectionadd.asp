<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/login.asp" -->
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
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_login_STRING
    MM_editCmd.CommandText = "INSERT INTO researchdirection (recordnum, dir, shortdir, longdir1, file1, longdir2, file2, longdir3, file3, longdir4, file4, longdir5, file5) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, 255, MM_IIF(Request.Form("recordnum"), Request.Form("recordnum"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 255, Request.Form("dir")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 255, Request.Form("shortdir")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 203, 1, 1073741823, Request.Form("longdir1")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 255, Request.Form("file1")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 203, 1, 1073741823, Request.Form("longdir2")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 202, 1, 255, Request.Form("file2")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 203, 1, 1073741823, Request.Form("longdir3")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 202, 1, 255, Request.Form("file3")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param10", 203, 1, 1073741823, Request.Form("longdir4")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param11", 202, 1, 255, Request.Form("file4")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param12", 203, 1, 1073741823, Request.Form("longdir5")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param13", 202, 1, 255, Request.Form("file5")) ' adVarWChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "backresearchdirectionlist.asp"
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
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_login_STRING
Recordset1_cmd.CommandText = "SELECT * FROM researchdirection" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!doctype html>
<html><!-- InstanceBegin template="/Templates/ProjectWebbacksystem.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<meta charset="utf-8">
<!-- InstanceBeginEditable name="doctitle" -->
<title>研究方向</title>
<link href="backdir.css" rel="stylesheet" type="text/css">
<!-- InstanceEndEditable -->
<style type="text/css">
body {
	background-color: #D5DADE;
}
</style>
<link href="../main.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="head" -->
<!-- InstanceEndEditable -->
</head>

<body>
<div class="webbody">
  <div class="head">
    <p><strong><img src="../indeximages/backsystem.jpg" width="900" height="204"></strong></p>
    <div class="navbar">
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <th width="114" height="50" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../backsystemindex.asp" class="navword">首页</a></th>
          <th width="114" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../researchteam/backresearchteamlist.asp" class="navword">团队成员</a></th>
          <th width="114" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="backresearchdirectionlist.asp" class="navword">研究方向</a></th>
          <th width="114" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../researchresult/researchresultlist.asp" class="navword">论文成果</a></th>
          <th width="114" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../teamactivity/teamactivitylist.asp" class="navword">团队动态</a></th>
          <th scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a   href="../contactus/backcontactusmod.asp" class="navword">联系我们</a></th>
          <th width="114" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a  href="../login/leavemessage.asp" class="navword">组内建议</a></th>
          <th width="114" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a  href="../index.asp" class="navword">返回前台</a></th>
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
    <th colspan="2" scope="col" class="leadertable" align="left" valign="middle"><img src="../indeximages/箭头.jpg" width="20" height="17" style="vertical-align:middle">&nbsp;&nbsp;学术带头人</th>
  </tr>
  <tr>
  <td height="5" colspan="2">
  </td>
  </tr>
  <tr>
    <td rowspan="3"  class="leaderphoto"><img src="../indeximages/教授.jpg" width="100" height="120"></td>
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




        <div id="backcurrentplace"><table width="230" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <th class="currentptable" scope="col"><img src="../indeximages/箭头.jpg" width="20" height="17" style="vertical-align:middle"></th>
	<th class="currentptable" scope="col">当前位置：</th>
    <th class="currentptable" scope="col"><a href="../backsystemindex.asp" class="dellink" onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">首页</a>-><a href="backresearchdirectionlist.asp" class="dellink" onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">研究方向</a></th>
  </tr>
</table>
</div>
<div id="backcondiv">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1"><table width="570" border="0" align="center" cellpadding="2" cellspacing="0">
  <tr valign="baseline">
    <td  nowrap width="150" scope="col" align="right">序号</td>
    <td align="left" scope="col" colspan="3"><input name="recordnum" type="text" id="recordnum" size="10" maxlength="255"></td>
  </tr>
  <tr valign="baseline">
    <td  nowrap width="150" scope="col" align="right">研究方向简称</td>
    <td align="left" scope="col" colspan="3"><input name="dir" type="text" id="dir" size="15" maxlength="255"></td>
  </tr>
  <tr valign="baseline">
    <td  nowrap width="150" scope="col" align="right">总体研究方向</td>
    <td align="left" scope="col" colspan="3"><input name="shortdir" type="text" id="shortdir" size="20" maxlength="255"></td>
  </tr>
  <tr>
    <td nowrap scope="col" align="right" valign="top">短期研究方向1</td>
    <td align="left" scope="col"><textarea name="longdir1" cols="20" rows="2" id="longdir1"></textarea></td>
    <td align="right">文件名：</td>
    <td><input name="file1" type="text" id="file1" size="20" maxlength="255"></td>
  </tr>
  <tr>
    <td nowrap scope="col" align="right" valign="top">短期研究方向2</td>
    <td align="left" scope="col"><textarea name="longdir2" cols="20" rows="2" id="longdir2"></textarea></td>
    <td align="right">文件名：</td>
    <td><input name="file2" type="text" id="file2" size="20" maxlength="255"></td>
  </tr>
  <tr>
    <td nowrap scope="col" align="right" valign="top">短期研究方向3</td>
    <td align="left" scope="col"><textarea name="longdir3" cols="20" rows="2" id="longdir3"></textarea></td>
        <td align="right">文件名：</td>
    <td><input name="file3" type="text" id="file3" size="20" maxlength="255"></td>
  </tr>
    <tr>
    <td nowrap scope="col" align="right" valign="top">短期研究方向4</td>
    <td align="left" scope="col"><textarea name="longdir4" cols="20" rows="2" id="longdir4"></textarea></td>
        <td align="right">文件名：</td>
    <td><input name="file4" type="text" id="file4" size="20" maxlength="255"></td>
  </tr>
    <tr>
    <td nowrap scope="col" align="right" valign="top">短期研究方向5</td>
    <td align="left" scope="col"><textarea name="longdir5" cols="20" rows="2" id="longdir5"></textarea></td>
        <td align="right">文件名：</td>
    <td><input name="file5" type="text" id="file5" size="20" maxlength="255"></td>
  </tr>
  <tr>
  <td colspan="4"  align="center">注意事项：文件名需带上拓展名，并且不支持中文
  </td>
  </tr>
  <tr>
  <td></td>
  <td><input name="提交" type="submit" id="提交" value="提交"><input name="重置" type="reset" value="重置"></td>
  </tr>
</table>
  <input type="hidden" name="MM_insert" value="form1">
</form>
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

	
	<!-- InstanceEndEditable -->
  </div>
  
    
   
  
</div>
</body>
<!-- InstanceEnd --></html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
