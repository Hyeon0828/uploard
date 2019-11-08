<html>
<%
  db_id = int(request("text_text"))
  Set db = Server.createObject ("ADODB.Connection")
  dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("board.accdb")
  db.Open (dbMain)

    Dim dummy_db, sql
    sql = "select * from board_table where ID = " & db_id
    Set dummy_db = db.execute(sql)

    db_title = dummy_db("db_title")
    db_date = dummy_db("db_date")
    db_name = dummy_db("db_name")
    db_pw = dummy_db("db_pw")
    db_lookup = int(dummy_db("db_lookup"))
    db_lookup = db_lookup + 1

    update_db = "update board_table set db_lookup = " & db_lookup & " where ID = " & db_id
    db.Execute(update_db)
    %>
  <script language = javascript>
  </script>
  <style>

  </style>

<body>
<form action="board_updat.asp" method="get" name="frm_content">
  <center>
  <table width=80%>
  <tr>
   <td><h2><% response.write db_title %><h2></td>
  </tr>

  <tr>
   <td>작성자 <% response.write db_name %></td>
   <td align="right">작성날짜 : <% response.write db_date %><br></td>
   </tr>

   <tr>
    <td  class="border" colspan="2">
    <iframe width= 100% height=100% frameborder=1 framespacing=1 marginheight=0 marginwidth=0 scrolling=no vspace=0 src="board_content2.asp?id=<%=db_id%>"></iframe>
    </td>
   </tr>

   <tr>
	  <td align=right colspan="2">
    <input type="button" value="목록보기" style="margin-left:30px;" onclick="window.open('board_li.asp?page=<%=int(request("page"))+1%>&text_sc=<%=request("text_sc")%>&sel_sc=<%=request("sel_sc")%>', '_self');">
    </td>
	 </tr>

  </table>
</center>
</form>
</body>
<% db.close %>
</html>
