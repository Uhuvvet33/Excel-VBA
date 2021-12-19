Public fso As New FileSystemObject 'global tanımlama
Sub foldersil()
    Dim fld As Scripting.Folder
    Dim outfld As Folder 'outlook folder, bunu kullanmayacağız
        
    fso.DeleteFolder "C:\Users\Volkan\Desktop\sil1"
    Set fld = fso.GetFolder("C:\Users\Volkan\Desktop\sil2")
    fld.Delete
End Sub
