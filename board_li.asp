<html>
  <head>
  </head>
 <script language = javascript>

   function ok() {
     window.open("board_insert.asp", "_self");
   }

 </script>
 <style>



 </style>


<body>

<center>
<form name="board" action="board_li.asp">

<table align=center>
<div>
<tr>
	<td>
	 <h1 size="10">게시판</h1>
	</td>
</tr>

<tr>
   <td align=center>작성자</td>&nbsp
   <td align=center>제목</td>&nbsp
   <td align=center>날짜</td>&nbsp
   <td align=center>조회수</td>&nbsp
</tr>
</div>
<%
Dim cnt_id, sql
Set db = Server.createObject ("ADODB.Connection")
dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("board.accdb")
db.Open (dbMain)

text_sc = request("text_sc")
sel_sc = request("sel_sc")
page_count = 0

For sp_i = 1 to len(text_sc)
if mid(text_sc, sp_i, 1) <> chr(32) then
  text_sc_x = txt_sc_x + mid(text_sc, sp_i, 1)
end if

next

all_data_num = "select * from db_num"

Set dummy_db = db.execute(all_data_num)
db_count = int(dummy_db("db_datanum"))

if int(request("page")) = 0 then '첫 페이지와 클릭 페이지 나눔
  db_page = 0

  else
  db_page = int(request("page")) - 1 '이전페이지
 end if

if text_sc_x = "" then  '노검색 노데이터 일때

sql = "select * from board_table Order by ID ASC" 'ID를 기준 정렬
Set search_db = db.execute(sql)

if search_db.EOF then
response.Write "게시글이 존재하지 않습니다."

else '노검색 데이터 존재
db_count_mod = 0

if db_count mod 10 = 0 then  '첫 페이지에 보일 레코드(필드 집합)개수 구하기, mod(나누기) 0일경우 10으로 출력
  db_count_mod = db_count mod 10 + 10

else
  db_count_mod = db_count mod 10
end if
search_db.MoveFirst

if  db_count <> 10 * db_page + db_count_mod then '마지막 페이지 제외하고 레코드 값 다음으로 넘기기
 for i = 1 to db_count - db_page * 10 - db_count_mod
 search_db.MoveNext
 next
end if

if db_page = 0 then '첫 페이지 나타낼 레코드 할당
 Do while Not search_db.EOF

  in_n = search_db("db_name")
  in_t = search_db("db_title")
  in_l = search_db("db_lookup")

  if in_l = "" then in_l = "0"

   string_date = search_db("db_date")
   date_string = Split(string_date," ")
   in_c = date_string(0)

   response.write "<tr><td align='center'>" & in_n & "</td><td align='center'>" &_
   "<a href='board_content.asp?text_text=" & search_db("ID") & "&page=" & db_page & "&text_sc=" & text_sc_x & "&sel_sc=" & sel_sc & "'>"&in_t&"</a>"&_
   "</td><td align='center'>" & in_c & "</td><td align='center'>" & in_l & "</td></tr>"
   search_db.MoveNext
  Loop

   else
   Do while Not page_count = 10

   page_count = page_count + 1
   in_n = search_db("db_name")
   in_t = search_db("db_title")
   in_l = search_db("db_lookup")

   if in_l = "" then in_l = "0"

   string_date = search_db("db_date")
   date_string = Split(string_date," ")
   in_c = date_string(0)

   response.write "<tr><td align='center'>" & in_n & "</td><td align='center'>" &_
   "<a href='board_content.asp?text_text=" & search_db("ID") & "&page=" & db_page & "&text_sc=" & text_sc_x & "&sel_sc=" & sel_sc & "'>"&in_t&"</a>"&_
   "</td><td align='center'>" & in_c & "</td><td align='center'>" & in_d & "</td></tr>"
   search_db.MoveNext
   Loop
   end if
   end if


   else

   if sel_sc = 1 then sql = "select * from board_table where db_spaceno like '%"&text_sc_x&"%'"
   if sel_sc = 2 then sql = "select * from board_table where db_name like '%"&text_sc_x&"%'"

   Set search_db = db.execute(sql)

   if search_db.EOF then
   response.Write "<script>alert('게시글이 없습니다.');window.open('board_li.asp','_self');</script>"

   else

   db_count = 0
   db_count_mod = 0

   Do while Not search_db.EOF '검색한 게시글 개수
   db_count = db_count + 1
   search_db.MoveNext
   Loop

   if db_count mod 10 = 0 then  '첫 페이지에 레코드(필드 집합) 개수, mod가 0일경우 10으로 출력
   db_count_mod = db_count mod 10 + 10
   else
   db_count_mod = db_count mod 10
   end if
   search_db.MoveFirst

   if  db_count <> 10 * db_page + db_count_mod then '마지막 쪽 제외하고 레코드 값 넘기기
   for i = 1 to db_count - db_page * 10 - db_count_mod
   search_db.MoveNext

   next
   end if

   if db_page = 0 then '첫 페이지 나타낼 레코드 할당
   Do while Not search_db.EOF

   in_n = search_db("db_name")
   in_t = search_db("db_title")
   in_l = search_db("db_lookup")

   if in_l = "" then in_l = "0"

   string_date = search_db("db_date")
   date_string = Split(string_date," ")
   in_c = date_string(0)

   response.write "<tr><td align='center'>" & in_n & "</td><td align='center'>" &_
   "<a href='board_content.asp?text_text=" & search_db("ID") & "&page=" & db_page & "&text_sc=" & text_sc_x & "&sel_sc=" & sel_sc & "'>"&in_t&"</a>"&_
                  "</td><td align='center'>" & in_c & "</td><td align='center'>" & in_l & "</td></tr>"
   search_db.MoveNext
   Loop

   else
   Do while Not page_count = 10

   page_count = page_count + 1
   in_n = search_db("db_name")
   in_t = search_db("db_title")
   in_l = search_db("db_lookup")

   if in_l = "" then in_l = "0"

   string_date = search_db("db_date")
   date_string = Split(string_date," ")
   in_c = date_string(0)

   response.write "<tr><td align='center'>" & in_n & "</td><td align='center'>" &_
   "<a href='board_content.asp?text_text=" & search_db("ID") & "&page=" & db_page & "&text_sc=" & text_sc_x & "&sel_sc=" & sel_sc & "'>"&in_b&"</a>"&_
   "</td><td align='center'>" & in_c & "</td><td align='center'>" & in_l & "</td></tr>"
   search_db.MoveNext
   Loop
   end if
  end if
end if
%>

<td align="left" colspan=5>
<input type="button" value="작성" onclick="ok();">
</td>
</table>

<%
if db_count mod 10 = 0 then
 for i = 1 to db_count/10
 response.write "<a href='board_li.asp?page=" & i & "&text_sc=" & text_sc_x & "&sel_sc=" & request("sel_sc") & "' target='_self'>" & i & "</a>"
 next

else
for i = 1 to db_count/10 + 1
response.write "<a href='board_li.asp?page=" & i & "&text_sc=" & text_sc_x & "&sel_sc=" & request("sel_sc") & "' target='_self'>" & i & "</a>"
next

end if

db.close()
set db = nothing
set search_db = nothing
%>
      </div>

      <div>

      <div>　　　　
        <select name="sel_sc" onchange="board.text_sc.focus();" style="width:100px; margin-right:3px;">
          <option value="1">글제목</option>
          <option value="2">작성자</option>
        </select>
        <input type="text" name="text_sc" maxlength="16">
        <input type="submit" value="검색" style="margin-left:3px;">
      </div>


  </form>
  </body>
 </html>
