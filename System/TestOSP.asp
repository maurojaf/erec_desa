<%
Set xmlDOM = Server.CreateObject("MSXML2.DOMDocument")
xmlDOM.async = False
xmlDOM.setProperty "ServerHTTPRequest", True
xmlDOM.Load("external.xml")
 
Set itemList = XMLDom.SelectNodes("news/article")
 
For Each itemAttrib In itemList
   newsSubject =itemAttrib.SelectSingleNode("subject").text
   newsExtract =itemAttrib.SelectSingleNode("extract").text
   newsDate =itemAttrib.SelectSingleNode("published").text
   %>
   <tr>
      <td><%=newsSubject%></td>
      <td><%=newsDate%></td>
      <td><%=newsExtract%></td>
   </tr>
<%
Next
 
Set xmlDOM = Nothing
Set itemList = Nothing
%>