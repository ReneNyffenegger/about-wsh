option explicit

private sub namedSpecialFolder(folder_name as variant, byref c as integer, spf as wshCollection)

   c = c+1

   cells(c, 1).value = folder_name
   cells(c, 2).value = spf(folder_name)
end sub


sub main()

    dim wsh as wshShell
    dim spf as wshCollection
    dim fld as variant


    set wsh = new wshShell

    set spf = wsh.specialFolders

    dim c as integer: c=0
    for each fld in spf
        c = c+1

        cells(c, 2).value = fld
    next fld

    c = c+1

    cells(c, 1).value = "------------------------------"


    call namedSpecialFolder("AllUsersDesktop"   , c, spf)
    call namedSpecialFolder("AllUsersStartMenu" , c, spf)
    call namedSpecialFolder("AllUsersPrograms"  , c, spf)
    call namedSpecialFolder("AllUsersStartup"   , c, spf)
    call namedSpecialFolder("Desktop"           , c, spf)
    call namedSpecialFolder("Favorites"         , c, spf)
    call namedSpecialFolder("Fonts"             , c, spf)
    call namedSpecialFolder("MyDocuments"       , c, spf)
    call namedSpecialFolder("NetHood"           , c, spf)
    call namedSpecialFolder("PrintHood"         , c, spf)
    call namedSpecialFolder("Programs"          , c, spf)
    call namedSpecialFolder("Recent"            , c, spf)
    call namedSpecialFolder("SendTo"            , c, spf)
    call namedSpecialFolder("StartMenu"         , c, spf)
    call namedSpecialFolder("Startup"           , c, spf)
    call namedSpecialFolder("Templates"         , c, spf)

    columns(1).autoFit

    activeWorkbook.saved = true

end sub
