<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/login.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_login_STRING
Recordset1_cmd.CommandText = "SELECT dir FROM researchdirection" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Recordset2
Dim Recordset2_cmd
Dim Recordset2_numRows

Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
Recordset2_cmd.ActiveConnection = MM_login_STRING
Recordset2_cmd.CommandText = "SELECT * FROM researchresult ORDER BY moddate DESC" 
Recordset2_cmd.Prepared = true

Set Recordset2 = Recordset2_cmd.Execute
Recordset2_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 18
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = 6
Repeat2__index = 0
Recordset2_numRows = Recordset2_numRows + Repeat2__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim Recordset2_total
Dim Recordset2_first
Dim Recordset2_last

' set the record count
Recordset2_total = Recordset2.RecordCount

' set the number of rows displayed on this page
If (Recordset2_numRows < 0) Then
  Recordset2_numRows = Recordset2_total
Elseif (Recordset2_numRows = 0) Then
  Recordset2_numRows = 1
End If

' set the first and last displayed record
Recordset2_first = 1
Recordset2_last  = Recordset2_first + Recordset2_numRows - 1

' if we have the correct record count, check the other stats
If (Recordset2_total <> -1) Then
  If (Recordset2_first > Recordset2_total) Then
    Recordset2_first = Recordset2_total
  End If
  If (Recordset2_last > Recordset2_total) Then
    Recordset2_last = Recordset2_total
  End If
  If (Recordset2_numRows > Recordset2_total) Then
    Recordset2_numRows = Recordset2_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (Recordset2_total = -1) Then

  ' count the total records by iterating through the recordset
  Recordset2_total=0
  While (Not Recordset2.EOF)
    Recordset2_total = Recordset2_total + 1
    Recordset2.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (Recordset2.CursorType > 0) Then
    Recordset2.MoveFirst
  Else
    Recordset2.Requery
  End If

  ' set the number of rows displayed on this page
  If (Recordset2_numRows < 0 Or Recordset2_numRows > Recordset2_total) Then
    Recordset2_numRows = Recordset2_total
  End If

  ' set the first and last displayed record
  Recordset2_first = 1
  Recordset2_last = Recordset2_first + Recordset2_numRows - 1
  
  If (Recordset2_first > Recordset2_total) Then
    Recordset2_first = Recordset2_total
  End If
  If (Recordset2_last > Recordset2_total) Then
    Recordset2_last = Recordset2_total
  End If

End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = Recordset2
MM_rsCount   = Recordset2_total
MM_size      = Recordset2_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
Recordset2_first = MM_offset + 1
Recordset2_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (Recordset2_first > MM_rsCount) Then
    Recordset2_first = MM_rsCount
  End If
  If (Recordset2_last > MM_rsCount) Then
    Recordset2_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
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
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!doctype html>
<html><!-- InstanceBegin template="/Templates/ProjectWeb.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<meta charset="utf-8">
<!-- InstanceBeginEditable name="doctitle" -->
<title>论文成果</title>
<link href="researchresult.css" rel="stylesheet" type="text/css">
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
    <p><strong><img src="../indeximages/ground.jpg" width="900" height="204"></strong></p>
    <div class="navbar">
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <th width="128" height="50" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../index.asp" class="navword">首页</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../researchteam/researchteamprofessor.asp" class="navword">团队成员</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../researchdirection/researchdirection.asp" class="navword">研究方向</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="researchresult.asp" class="navword">论文成果</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../teamactivity/teamactivity.asp" class="navword">团队动态</a></th>
          <th scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a href="../contactus/contactus.asp" class="navword">联系我们</a></th>
          <th width="128" scope="col" onmouseover="this.bgColor='#026bff'" onmouseout="this.bgColor=''"><a  href="../login/backsystemlogin.asp" class="navword">管理入口</a></th>
        </tr>
      </table>
      <br></div>  
      <div style="height:50">
    <p></p>
    </div>
    <!-- InstanceBeginEditable name="EditRegion" -->
         <div id="steamdiv">
      <table width="220" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <th class="commonbtable">论文成果</th>
          </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
  <tr>
    <th class="commonsword" scope="col" onmouseover="this.bgColor='#3388ff'" onmouseout="this.bgColor=''" ><A HREF="researchresultclassify.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "dir=" & Recordset1.Fields.Item("dir").Value %>" class="commonsword"><%=(Recordset1.Fields.Item("dir").Value)%></A></th>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
      </table>
    </div>
    <div id="frontcurrentplace"><table width="230" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <th class="currentptable" scope="col" ><img src="../indeximages/箭头.jpg" width="20" height="17" style="vertical-align:middle"></th>
	<th class="currentptable" scope="col" >当前位置：</th>
    <th class="currentptable" scope="col" ><a href="../index.asp"  class="dellink" onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">首页</a>-><a href="researchresult.asp" class="dellink" onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'">论文成果</a></th>
  </tr>
</table>
</div>
<div id="frontcondivresult"><table width="670" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <th scope="col" width="550" align="left">&nbsp;&nbsp;学术论文列表</th>
    <th width="20"></th>
    <th align="left" width="100">发布日期</th>
  </tr>
  <% 
While ((Repeat2__numRows <> 0) AND (NOT Recordset2.EOF)) 
%>
    <tr>
      <th scope="col" align="left" class="footsong" height="45">&nbsp;<a href="upload/<%=(Recordset2.Fields.Item("place").Value)%>" class="dellink" onmouseover="this.style.color='blue'" onmouseout="this.style.color='black'" target="_blank"><%=(Recordset2.Fields.Item("conname").Value)%></a></th>
      <th></th>
      <th align="left" class="footsong" ><%=(Recordset2.Fields.Item("moddate").Value)%></th>
    </tr>
    <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  Recordset2.MoveNext()
Wend
%>
</table>
<hr>

  <table border="0" align="center">
    <tr>
      <td><% If MM_offset <> 0 Then %>
          <a href="<%=MM_moveFirst%>">第一页</a>
          <% End If ' end MM_offset <> 0 %></td>
      <td><% If MM_offset <> 0 Then %>
          <a href="<%=MM_movePrev%>">前一页</a>
          <% End If ' end MM_offset <> 0 %></td>
      <td><% If Not MM_atTotal Then %>
          <a href="<%=MM_moveNext%>">下一个</a>
          <% End If ' end Not MM_atTotal %></td>
      <td><% If Not MM_atTotal Then %>
          <a href="<%=MM_moveLast%>">最后一页</a>
          <% End If ' end Not MM_atTotal %></td>
    </tr>
  </table>
  <table align="center">
  <tr>
  <td>
记录 <%=(Recordset2_first)%> 到 <%=(Recordset2_last)%> (总共 <%=(Recordset2_total)%>)
</td>
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
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>
