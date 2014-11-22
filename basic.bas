' c:\lib\runVBAFilesInOffice\runVBAFilesInOffice.vbs -excel -wsh basic -c main

sub main()

   dim sh as WshShell
   dim nw as WshNetwork

   activeWorkbook.saved =true

end sub
