$directory = "c:\test\"
$ol = New-Object -ComObject Outlook.Application
$files = Get-ChildItem $directory
foreach ($file in $files)
{
  $msg = $ol.CreateItemFromTemplate($directory + $file)
  $headers = $msg.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
  $headers > ($file.name +".txt")
}