$shell = New-Object -ComObject Shell.Application
$src = $shell.NameSpace('C:\Programs\EasyFace\CleanFaceKms\output').ParseName('CleanFaceKms.face')
$dst = $shell.NameSpace('shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}').Items() | Where-Object { $_.Name -like '*Redmi Note 14*' }
$storage = $dst.GetFolder.Items() | Where-Object { $_.Name -like '*Internal*' -or $_.Name -like '*Almacenamiento*' }
$docs = $storage.GetFolder.Items() | Where-Object { $_.Name -eq 'Documents' }
$docs.GetFolder.CopyHere($src)