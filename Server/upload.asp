<!--#include FILE="include.inc"-->
<%
dim upload,file,filepath
filepath="uploaded/"
set upload=new upload_file ''建立上传对象
for each formName in upload.file ''列出所有上传了的文件
 set file=upload.file(formName)  ''生成一个文件对象
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  fname = file.filename
  file.SaveAs Server.mappath(filepath&fname)   ''保存文件
 end if
set file=nothing
next
set upload=nothing  ''删除此对象
%>