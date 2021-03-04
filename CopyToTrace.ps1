$traceSrcLocation = 'C:\Projects\SysKitTrace'
$msCloudLoginAssistantSrcLocation = 'C:\Projects\MSCloudLoginAssistant-SysKit'
$m36SrcLocation = 'C:\Projects\M65DSC-SysKit'
Remove-Item "$traceSrcLocation\SysKit.SPDocKit.Office365\UtilProjects\MSCloudLoginAssistant\**" -Force -Recurse
Copy-Item "$msCloudLoginAssistantSrcLocation\Modules\MSCloudLoginAssistant\**" "$traceSrcLocation\SysKit.SPDocKit.Office365\UtilProjects\MSCloudLoginAssistant\" -Force -Recurse


Remove-Item "$traceSrcLocation\SysKit.SPDocKit.Office365\UtilProjects\Microsoft365DSC\**" -Force -Recurse
Copy-Item "$m36SrcLocation\Modules\Microsoft365DSC\**" "$traceSrcLocation\SysKit.SPDocKit.Office365\UtilProjects\Microsoft365DSC\" -Force -Recurse
