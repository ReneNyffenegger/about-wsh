' c:\lib\runVBAFilesInOffice\runVBAFilesInOffice.vbs -excel -wsh basic -c main

sub main()

   dim wsh as WshShell

   activeWorkbook.saved =true

end sub
