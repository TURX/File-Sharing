<!--#include FILE="include.inc"-->
<%
dim upload,file,filepath
filepath="uploaded/"
set upload=new upload_file ''�����ϴ�����
for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  fname = file.filename
  file.SaveAs Server.mappath(filepath&fname)   ''�����ļ�
 end if
set file=nothing
next
set upload=nothing  ''ɾ���˶���
%>