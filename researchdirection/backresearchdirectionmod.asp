<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/login.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="../login/backsystemlogin.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
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
If (CStr(Request("MM_update")) = "form2") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_login_STRING
    MM_editCmd.CommandText = "UPDATE researchdirection SET recordnum = ?, dir = ?, shortdir = ?, longdir1 = ?, longdir2 = ?, longdir3 = ?, longdir4 = ?, longdir5 = ?, file1 = ?, file2 = ?, file3 = ?, file4 = ?, file5 = ? WHERE ID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 255, Request.Form("recordnum")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 255, Request.Form("dir")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 255, Request.Form("shortdir")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 203, 1, 1073741823, Request.Form("longdir1")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 203, 1, 1073741823, Request.Form("longdir2")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 203, 1, 1073741823, Request.Form("longdir3")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 203, 1, 1073741823, Request.Form("longdir4")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 203, 1, 1073741823, Request.Form("longdir5")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param9", 202, 1, 255, Request.Form("file1")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param10", 202, 1, 255, Request.Form("file2")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param11", 202, 1, 255, Request.Form("file3")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param12", 202, 1, 255, Request.Form("file4")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param13", 202, 1, 255, Request.Form("file5")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param14", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
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
Recordset1_cmd.CommandText = "SELECT * FROM researchdirection WHERE ID = ?" 
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
      <form method="post" action="<%=MM_editAction%>" name="form2">
        <table align="center" cellpadding="2">
          <tr valign="baseline">
            <td nowrap align="right" width="150">序号</td>
            <td colspan="3"><input type="text" name="recordnum" value="<%=(Recordset1.Fields.Item("recordnum").Value)%>" size="10"></td>
          </tr>
          <tr valign="baseline">
            <td nowrap align="right">研究方向简称</td>
            <td colspan="3"><input type="text" name="dir" value="<%=(Recordset1.Fields.Item("dir").Value)%>" size="15"></td>
          </tr>
          <tr valign="baseline">
            <td nowrap align="right">总体研究方向</td>
            <td colspan="3"><input type="text" name="shortdir" value="<%=(Recordset1.Fields.Item("shortdir").Value)%>" size="20"></td>
          </tr>
          <tr>
            <td nowrap align="right">短期研究方向1</td>
            <td><textarea name="longdir1" cols="20" rows="2"><%=(Recordset1.Fields.Item("longdir1").Value)%></textarea></td>
 			<td nowrap align="right">文件名</td>
            <td><input name="file1" type="text" value="<%=(Recordset1.Fields.Item("file1").Value)%>" size="20"></td>
          </tr>
          <tr >
            <td nowrap align="right">短期研究方向2</td>
            <td><textarea name="longdir2" cols="20" rows="2"><%=(Recordset1.Fields.Item("longdir2").Value)%></textarea></td>
            <td nowrap align="right">文件名</td>
            <td><input type="text" name="file2" value="<%=(Recordset1.Fields.Item("file2").Value)%>" size="20"></td>
          </tr>
          <tr >
            <td nowrap align="right">短期研究方向3</td>
            <td><textarea name="longdir3" cols="20" rows="2"><%=(Recordset1.Fields.Item("longdir3").Value)%></textarea></td>
            <td nowrap align="right">文件名</td>
            <td><input type="text" name="file3" value="<%=(Recordset1.Fields.Item("file3").Value)%>" size="20"></td>
          </tr>
          <tr >
            <td nowrap align="right">短期研究方向4</td>
            <td><textarea name="longdir4" cols="20" rows="2"><%=(Recordset1.Fields.Item("longdir4").Value)%></textarea></td>
            <td nowrap align="right">文件名</td>
            <td><input type="text" name="file4" value="<%=(Recordset1.Fields.Item("file4").Value)%>" size="20"></td>
          </tr>
          <tr >
            <td nowrap align="right">短期研究方向5</td>
            <td><textarea name="longdir5" cols="20" rows="2"><%=(Recordset1.Fields.Item("longdir5").Value)%></textarea></td>
            <td nowrap align="right">文件名</td>
            <td><input type="text" name="file5" value="<%=(Recordset1.Fields.Item("file5").Value)%>" size="20"></td>
          </tr>

            <tr>
              <td colspan="4"  align="center">注意事项：文件名需带上拓展名，并且不支持中文
              </td>
              </tr>
          <tr valign="baseline">
            <td nowrap align="right">&nbsp;</td>
            <td><input type="submit" value="更新记录"></td>
          </tr>
        </table>
        <input type="hidden" name="MM_update" value="form2">
        <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("ID").Value %>">
      </form>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
    
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
