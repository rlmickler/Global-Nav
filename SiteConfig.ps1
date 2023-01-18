<#
.SYNOPSIS
    Apply site configurations to Modern SharePoint Sites.
.EXAMPLE
    SiteConfig -SiteUrl "https://<tenant>.sharepoint.com/sites/site"
    SiteConfig -SiteUrl "https://<tenant>.sharepoint.com/sites/site" -HubUrl "https://<tenant>.sharepoint.com/sites/Parent-Hub-Site"
    SiteConfig -SiteUrl "https://<tenant>.sharepoint.com/sites/site" -SiteType "Communication site"

  
.NOTES
    Requires PnP PowerShell
    SharePoint REST API Metadata Explorer: https://s-kainet.github.io/sp-rest-explorer/#/
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory, HelpMessage="The URL of a new or existing site.")]
    [ValidatePattern('^https\:\/\/[a-z]*\.sharepoint\.com\/sites\/etf.*$')]
    [string]$SiteUrl,

    [Parameter(Mandatory=$false, HelpMessage="The URL of a hub site to associate with.")]
    [ValidatePattern('^https\:\/\/[a-z]*\.sharepoint\.com\/sites\/etf.*$')]
    [string]$HubUrl,

    [Parameter(Mandatory=$false, DontShow, HelpMessage="The type of new site to create.")]
    [ValidateSet("Communication site", "Team site (no M365 group)")]
    [string]$SiteType
)

function Get-SiteStatus {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidatePattern('^https\:\/\/[a-z]*\.sharepoint\.com\/sites\/etf.*$')]
        [string]$Url
    )

    Begin {
        $web = Invoke-PnPSPRestMethod -Url "/_api/web?`$select=CurrentUser,Url&`$expand=CurrentUser"

        if (($Url -replace '(?<=\.com).*$') -eq $web.Url) {
            $logo = "`n$(" "*3)_=+#####!`n###########|`n###/$(" "*4)(##|(@)`n###$(" "*2)######|$(" "*3)\`n###/$(" "*3)/###|$(" "*3)(@) SharePoint`n#######$(" "*2)##|$(" "*3)/`n###)$(" "*4)/##|(@)`n###########|`n$(" "*3)**=+####!"
            Write-Host "`n$($logo)" -ForegroundColor DarkCyan
            Write-Host "`nTenant Url: $($web.Url)"
            Write-Host "User: $($web.CurrentUser.Email)"
        }
    }

    Process {
        Write-Host "`nSITE STATUS"
        Write-Host " - Url: $($Url)"
        Write-Host " - Status: " -NoNewline

        $siteStatus = (Invoke-PnPSPRestMethod -Url "/_api/SPSiteManager/status?url='$($Url)'").SiteStatus

        switch ($siteStatus) {
            0 { Write-Host "Not Found. The site doesn't exist." -ForegroundColor Red; Break }
            1 { Write-Host "Provisioning. The site is currently being provisioned." -ForegroundColor Magenta; Break }
            2 { Write-Host "Ready. The site has been created." -ForegroundColor Green; Break }
            3 { Write-Host "Error. An error occurred while provisioning the site." -ForegroundColor Red; Exit }
            4 { Write-Host "Site with requested URL already exist." -ForegroundColor Magenta; Exit }
        }

        if ($siteStatus -eq 1) {
            $marker = ""

            do {
                Write-Host "$($marker + ".")" -NoNewline
                Start-Sleep -Seconds 2

                $siteStatus = (Invoke-PnPSPRestMethod -Url "/_api/SPSiteManager/status?url='$($Url)'").SiteStatus

                if ($siteStatus -eq 2) {
                    Write-Host "`nReady. The site has been created." -ForegroundColor Green
                    Break
                }
            } until ($siteStatus -eq 2)
        }
    }

    End {
        return $siteStatus
    }
}

function New-CreateSite {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidatePattern('^https\:\/\/[a-z]*\.sharepoint\.com\/sites\/etf.*$')]
        [string]$Url,

        [Parameter(Mandatory)]
        [ValidateSet("Communication site", "Team site (no M365 group)")]
        [string]$Template,

        [Parameter(Mandatory)]
        [string]$SiteAdmin
    )

    Begin {
        $title = "$($Url -replace '.*?(\.com\/sites\/etf)')"
        $title = $title -replace "[^\p{L}\p{Nd}]+"
        
        if ($title[0] -match "^\d+$" -eq $false) {
          $title = $title.Substring(0,1).ToUpper() + $title.Substring(1)
        }

        if ($Template -eq "Communication site") {
            $siteTemplate = "SITEPAGEPUBLISHING#0"
        } else {
            $siteTemplate = "STS#3"
        }
    }

    Process {
        Write-Host "`nCREATING NEW SITE"
        Write-Host " - Title: $($title)"
        Write-Host " - Url: $($Url)"
        Write-Host " - Template: $($siteTemplate)"
        Write-Host " - Owner: $($SiteAdmin)"

        $output = Invoke-PnPSPRestMethod -Method Post `
            -Url "/_api/SPSiteManager/Create" `
            -ContentType "application/json;odata.metadata=none" `
            -Content @{
                "request" = @{
                    "Title" = $title
                    "Url" = $Url
                    "Lcid" = 1033
                    "ShareByEmailEnabled" = $false
                    "Classification" = ""
                    "Description" = ""
                    "WebTemplate" = $siteTemplate
                    "SiteDesignId" = "00000000-0000-0000-0000-000000000000"
                    "Owner" = $SiteAdmin
                    "WebTemplateExtensionId" = "00000000-0000-0000-0000-000000000000"
                }
            } -ErrorAction Stop
    }

    End {
        return $output
    }
}

function Copy-SiteLogos {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidatePattern('^https\:\/\/[a-z]*\.sharepoint\.com\/sites\/etf.*$')]
        [string]$SourceUrl,

        [Parameter(Mandatory)]
        [string]$SourceList,

        [Parameter(Mandatory=$false)]
        [string]$SourceFolder,

        [Parameter(Mandatory)]
        [ValidatePattern('^https\:\/\/[a-z]*\.sharepoint\.com\/sites\/etf.*$')]
        [string]$TargetUrl,

        [Parameter(Mandatory=$false)]
        [string[]]$FileNames
    )

    Begin {
        $web = Invoke-PnPSPRestMethod -Url "/_api/web?`$select=Url,WebTemplateConfiguration"

        if ($SourceUrl -eq $web.Url -or $SourceUrl -eq $TargetUrl) {
            return
        }
    }

    Process {
        $ensureSiteAssets = Invoke-PnPSPRestMethod -Method Post -Url "/_api/web/lists/ensureSiteAssetsLibrary()"
        $siteAssets = Invoke-PnPSPRestMethod -Url "/_api/web/lists('$($ensureSiteAssets.Id)')/RootFolder"

        $srcList = (Invoke-PnPSPRestMethod -Url "$($SourceUrl)/_api/web/lists?`$select=Title&`$expand=RootFolder&`$filter=Title eq '$($SourceList)'").value
        $srcFolders = (Invoke-PnPSPRestMethod -Url "$($SourceUrl)/_api/web/GetFolderByServerRelativeUrl('$($srcList.RootFolder.ServerRelativeUrl)')/Folders").value

        if ($srcFolders.Name -contains "$($SourceFolder)") {
            $folder = $srcFolders[$srcFolders.Name.IndexOf("$($SourceFolder)")]
            $itemsUrl = $folder.ServerRelativeUrl
        } else {
            $itemsUrl = $srcList.RootFolder.ServerRelativeUrl
        }

        $items = (Invoke-PnPSPRestMethod -Url "$($SourceUrl)/_api/web/GetFolderByServerRelativeUrl('$($itemsUrl)')/Files").value
        $items = $items | Where-Object { $_.Name -in $FileNames }

        if ($web.WebTemplateConfiguration -ne "SITEPAGEPUBLISHING#0") {
            $items = $items | Where-Object { $_.Name -notmatch "(footer|footerlogo)" }
        }

        $links = $items | ForEach-Object { "$($SourceUrl -replace '(?<=\.com).*$')$($_.ServerRelativeUrl)" }

        $copy = (Invoke-PnPSPRestMethod -Method Post `
            -Url "/_api/site/CreateCopyJobs" `
            -ContentType "application/json;odata.metadata=none" `
            -Content @{
                "exportObjectUris" = [string[]]$links
                "destinationUri" = "$($TargetUrl -replace '(?<=\.com).*$')$($siteAssets.ServerRelativeUrl)"
                "options" = @{
                    "AllowSchemaMismatch" = $true
                    "BypassSharedLock" = $true
                    "IgnoreVersionHistory" = $true
                    "IncludeItemPermissions" = $false
                    "IsMoveMode" = $false
                    "NameConflictBehavior" = 1
                }
            } -Raw | ConvertFrom-Json).value

        if ($copy.JobId) {
            Write-Host " - Copied logos to $($siteAssets.ServerRelativeUrl)"
            Start-Sleep -Seconds 3
        }
    }

    End {}
}

function Set-Logo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$SiteLogo,

        [Parameter(Mandatory=$false)]
        [string]$ThumbnailLogo,

        [Parameter(Mandatory=$false)]
        [string]$FooterLogo
    )

    Begin {
        $web = Invoke-PnPSPRestMethod -Url "/_api/web?`$select=FooterEnabled,ServerRelativeUrl,Url,WebTemplateConfiguration"

        if ($web.WebTemplateConfiguration -ne "SITEPAGEPUBLISHING#0") {
            [bool]$FooterLogo = $false
        }
    }

    Process {
        $ensureSiteAssets = Invoke-PnPSPRestMethod -Method Post -Url "/_api/web/lists/ensureSiteAssetsLibrary()"
        $siteAssets = Invoke-PnPSPRestMethod -Url "/_api/web/lists('$($ensureSiteAssets.Id)')/RootFolder"

        if ($siteAssets.ItemCount) {
            $files = (Invoke-PnPSPRestMethod -Url "/_api/web/GetFolderByServerRelativeUrl('$($siteAssets.ServerRelativeUrl)')/Files").value
            $logos = $files | Where-Object { $_.ServerRelativeUrl -match "$($siteAssets.ServerRelativeUrl)/$($_.Name)" }

            if ([bool]$SiteLogo) {
                $siteImg = $logos | Where-Object { $_.Name -Match $SiteLogo }

                $output = Invoke-PnPSPRestMethod -Method Post `
                    -Url "/_api/SiteIconManager/SetSiteLogo" `
                    -ContentType "application/json;odata.metadata=none" `
                    -Content @{
                        "relativeLogoUrl" = "$($siteImg.ServerRelativeUrl)"
                        "type" = 0
                        "aspect" = 1
                    }

                if ($output."odata.null") {
                    Write-Host " - Set site logo to $($siteImg.Name)"
                }
            }

            if ([bool]$ThumbnailLogo) {
                $thumbnailImg = $logos | Where-Object { $_.Name -Match $ThumbnailLogo }

                $output = Invoke-PnPSPRestMethod -Method Post `
                    -Url "/_api/SiteIconManager/SetSiteLogo" `
                    -ContentType "application/json;odata.metadata=none" `
                    -Content @{
                        "relativeLogoUrl" = "$($thumbnailImg.ServerRelativeUrl)"
                        "type" = 0
                        "aspect" = 0
                    }

                if ($output."odata.null") {
                    Write-Host " - Set thumbnail logo to $($thumbnailImg.Name)"
                }
            }

            if ([bool]$FooterLogo) {
                if ($web.FooterEnabled) {
                    $footerImg = $logos | Where-Object { $_.Name -Match $FooterLogo }
                    $footerImgUrl = "$($web.Url -replace '(?<=\.com).*$')$($footerImg.ServerRelativeUrl)"

                    $body = '{
                        "menuState": {
                            "StartingNodeTitle": "13b7c916-4fea-4bb2-8994-5cf274aeb530",
                            "SPSitePrefix": "/",
                            "SPWebPrefix": "' + $web.ServerRelativeUrl + '",
                            "FriendlyUrlPrefix": "",
                            "SimpleUrl": "",
                            "Nodes": [
                                {
                                    "NodeType": 0,
                                    "Title": "2e456c2e-3ded-4a6c-a9ea-f7ac4c1b5100",
                                    "SimpleUrl": "' + $footerImgUrl + '",
                                    "FriendlyUrlSegment": ""
                                }
                            ]
                        }
                    }'

                    $output = (Invoke-PnPSPRestMethod -Method Post `
                        -Url "/_api/navigation/SaveMenuState" `
                        -ContentType "application/json;odata.metadata=none" `
                        -Content $body -Raw | ConvertFrom-Json).value

                    if ($output -eq 200) {
                        Write-Host " - Set footer logo to $($footerImg.Name)"
                    }
                }
            }

        }
    }

    End {}
}

function Add-GlobalNavigation {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [ValidateSet("Add","Remove")]
        [string]$Action = "Add"
    )

    Begin {
        $settings = @{
            "Location" = "ClientSideExtension.ApplicationCustomizer"
            "Title" = "ETFGlobalNav"
            "Name" = "ETFGlobalNav"
            "ClientSideComponentId" = "ad6c4ae1-5c6c-48f5-9ad6-26dd83b393ab" 
            "ClientSideComponentProperties" = "{`"TermSetName`":`"ETFGlobalNav`",`"SearchPageUrl`":`"https://wigov.sharepoint.com/sites/ETF/SitePages/search.aspx`",`"SearchBoxPlaceholder`":`"ETF Connect Global Search`",`"AlertSiteUrl`":`"https://wigov.sharepoint.com/sites/ETF`",`"AlertListName`":`"Alerts`"}"
            "HostProperties" = "{`"preAllocatedApplicationCustomizerTopHeight`":`"44`"}"
        }
    }

    Process {
        $component = (Invoke-PnPSPRestMethod `
            -Url "/_api/web/UserCustomActions?$select=Title,Id&filter=ClientSideComponentId eq $($settings.ClientSideComponentId)" `
            -ContentType "application/json;odata.metadata=none").value

        if ($component.Count -eq 0 -and $Action -eq "Add") {
            $output = Invoke-PnPSPRestMethod -Method Post `
                -Url "/_api/web/UserCustomActions" `
                -ContentType "application/json;odata.metadata=none" `
                -Content @{
                    "Location" = $settings.Location
                    "Title" = $settings.Title
                    "Name" = $settings.Name
                    "ClientSideComponentId" = $settings.ClientSideComponentId
                    "ClientSideComponentProperties" = $settings.ClientSideComponentProperties
                    "HostProperties" = $settings.HostProperties
                }

            if ($output.ClientSideComponentId -eq $settings.ClientSideComponentId) {
                Write-Host " - Added global navigation"
            }
        } elseif ($component.Count -and $Action -eq "Remove") {
            $output = Invoke-PnPSPRestMethod -Method Post `
                -Url "/_api/web/UserCustomActions/GetById('$($component.Id)')/DeleteObject" `
                -ContentType "application/json;odata.metadata=none"
            
            if ($output."odata.null") {
                Write-Host " - Removed global navigation"
            }
        }
    }

    End {}
}

function Set-SiteTheme {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [ValidateSet("Teal", "Blue", "Orange", "Red", "Purple", "Green", "Gray", "DarkYellow", "DarkBlue")]
        [string]$Color,

        [Parameter(Mandatory=$false)]
        [switch]$ForceUpdate
    )

    Begin {
        $web = Invoke-PnPSPRestMethod -Url "/_api/web?`$select=PrimaryColor"

        $currentColor = switch ($web.PrimaryColor) {
            "#0078D4" { "Blue" }
            "#CA5010" { "Orange" }
            "#A4262C" { "Red" }
            "#8764B8" { "Purple" }
            "#498205" { "Green" }
            "#69797E" { "Gray" }
            "#FFC83D" { "DarkYellow" }
            "#3A96DD" { "DarkBlue" }
            "#03787C" { "Teal" }
            Default { "none" }
        }
    }

    Process {
        if ($Color -ne $currentColor -or $ForceUpdate) {
            $theme = switch ($Color) {
                "Blue" { "{'name':'Blue','themeJson':'{`"backgroundImageUri`":`"`",`"palette`":{`"themePrimary`":{`"R`":0,`"G`":120,`"B`":212,`"A`":255},`"themeLighterAlt`":{`"R`":239,`"G`":246,`"B`":252,`"A`":255},`"themeLighter`":{`"R`":222,`"G`":236,`"B`":249,`"A`":255},`"themeLight`":{`"R`":199,`"G`":224,`"B`":244,`"A`":255},`"themeTertiary`":{`"R`":113,`"G`":175,`"B`":229,`"A`":255},`"themeSecondary`":{`"R`":43,`"G`":136,`"B`":216,`"A`":255},`"themeDarkAlt`":{`"R`":16,`"G`":110,`"B`":190,`"A`":255},`"themeDark`":{`"R`":0,`"G`":90,`"B`":158,`"A`":255},`"themeDarker`":{`"R`":0,`"G`":69,`"B`":120,`"A`":255},`"accent`":{`"R`":135,`"G`":100,`"B`":184,`"A`":255},`"neutralLighterAlt`":{`"R`":248,`"G`":248,`"B`":248,`"A`":255},`"neutralLighter`":{`"R`":244,`"G`":244,`"B`":244,`"A`":255},`"neutralLight`":{`"R`":234,`"G`":234,`"B`":234,`"A`":255},`"neutralQuaternaryAlt`":{`"R`":218,`"G`":218,`"B`":218,`"A`":255},`"neutralQuaternary`":{`"R`":208,`"G`":208,`"B`":208,`"A`":255},`"neutralTertiaryAlt`":{`"R`":200,`"G`":200,`"B`":200,`"A`":255},`"neutralTertiary`":{`"R`":166,`"G`":166,`"B`":166,`"A`":255},`"neutralSecondary`":{`"R`":102,`"G`":102,`"B`":102,`"A`":255},`"neutralPrimaryAlt`":{`"R`":60,`"G`":60,`"B`":60,`"A`":255},`"neutralPrimary`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255},`"neutralDark`":{`"R`":33,`"G`":33,`"B`":33,`"A`":255},`"black`":{`"R`":0,`"G`":0,`"B`":0,`"A`":255},`"white`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryBackground`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryText`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255}},`"cacheToken`":`"`",`"isDefault`":true,`"version`":`"`"}'}" }
                "Orange" { "{'name':'Orange','themeJson':'{`"backgroundImageUri`":`"`",`"palette`":{`"themePrimary`":{`"R`":202,`"G`":80,`"B`":16,`"A`":255},`"themeLighterAlt`":{`"R`":253,`"G`":247,`"B`":244,`"A`":255},`"themeLighter`":{`"R`":246,`"G`":223,`"B`":210,`"A`":255},`"themeLight`":{`"R`":239,`"G`":196,`"B`":173,`"A`":255},`"themeTertiary`":{`"R`":223,`"G`":143,`"B`":100,`"A`":255},`"themeSecondary`":{`"R`":208,`"G`":98,`"B`":40,`"A`":255},`"themeDarkAlt`":{`"R`":181,`"G`":73,`"B`":15,`"A`":255},`"themeDark`":{`"R`":153,`"G`":62,`"B`":12,`"A`":255},`"themeDarker`":{`"R`":113,`"G`":45,`"B`":9,`"A`":255},`"accent`":{`"R`":152,`"G`":111,`"B`":11,`"A`":255},`"neutralLighterAlt`":{`"R`":248,`"G`":248,`"B`":248,`"A`":255},`"neutralLighter`":{`"R`":244,`"G`":244,`"B`":244,`"A`":255},`"neutralLight`":{`"R`":234,`"G`":234,`"B`":234,`"A`":255},`"neutralQuaternaryAlt`":{`"R`":218,`"G`":218,`"B`":218,`"A`":255},`"neutralQuaternary`":{`"R`":208,`"G`":208,`"B`":208,`"A`":255},`"neutralTertiaryAlt`":{`"R`":200,`"G`":200,`"B`":200,`"A`":255},`"neutralTertiary`":{`"R`":166,`"G`":166,`"B`":166,`"A`":255},`"neutralSecondary`":{`"R`":102,`"G`":102,`"B`":102,`"A`":255},`"neutralPrimaryAlt`":{`"R`":60,`"G`":60,`"B`":60,`"A`":255},`"neutralPrimary`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255},`"neutralDark`":{`"R`":33,`"G`":33,`"B`":33,`"A`":255},`"black`":{`"R`":0,`"G`":0,`"B`":0,`"A`":255},`"white`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryBackground`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryText`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255}},`"cacheToken`":`"`",`"isDefault`":true,`"version`":`"`"}'}" }
                "Red" { "{'name':'Red','themeJson':'{`"backgroundImageUri`":`"`",`"palette`":{`"themePrimary`":{`"R`":164,`"G`":38,`"B`":44,`"A`":255},`"themeLighterAlt`":{`"R`":251,`"G`":244,`"B`":244,`"A`":255},`"themeLighter`":{`"R`":240,`"G`":211,`"B`":212,`"A`":255},`"themeLight`":{`"R`":227,`"G`":175,`"B`":178,`"A`":255},`"themeTertiary`":{`"R`":200,`"G`":108,`"B`":112,`"A`":255},`"themeSecondary`":{`"R`":174,`"G`":56,`"B`":62,`"A`":255},`"themeDarkAlt`":{`"R`":147,`"G`":34,`"B`":39,`"A`":255},`"themeDark`":{`"R`":124,`"G`":29,`"B`":33,`"A`":255},`"themeDarker`":{`"R`":91,`"G`":21,`"B`":25,`"A`":255},`"accent`":{`"R`":202,`"G`":80,`"B`":16,`"A`":255},`"neutralLighterAlt`":{`"R`":248,`"G`":248,`"B`":248,`"A`":255},`"neutralLighter`":{`"R`":244,`"G`":244,`"B`":244,`"A`":255},`"neutralLight`":{`"R`":234,`"G`":234,`"B`":234,`"A`":255},`"neutralQuaternaryAlt`":{`"R`":218,`"G`":218,`"B`":218,`"A`":255},`"neutralQuaternary`":{`"R`":208,`"G`":208,`"B`":208,`"A`":255},`"neutralTertiaryAlt`":{`"R`":200,`"G`":200,`"B`":200,`"A`":255},`"neutralTertiary`":{`"R`":166,`"G`":166,`"B`":166,`"A`":255},`"neutralSecondary`":{`"R`":102,`"G`":102,`"B`":102,`"A`":255},`"neutralPrimaryAlt`":{`"R`":60,`"G`":60,`"B`":60,`"A`":255},`"neutralPrimary`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255},`"neutralDark`":{`"R`":33,`"G`":33,`"B`":33,`"A`":255},`"black`":{`"R`":0,`"G`":0,`"B`":0,`"A`":255},`"white`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryBackground`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryText`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255}},`"cacheToken`":`"`",`"isDefault`":true,`"version`":`"`"}'}" }
                "Purple" { "{'name':'Purple','themeJson':'{`"backgroundImageUri`":`"`",`"palette`":{`"themePrimary`":{`"R`":135,`"G`":100,`"B`":184,`"A`":255},`"themeLighterAlt`":{`"R`":249,`"G`":248,`"B`":252,`"A`":255},`"themeLighter`":{`"R`":233,`"G`":226,`"B`":244,`"A`":255},`"themeLight`":{`"R`":215,`"G`":201,`"B`":234,`"A`":255},`"themeTertiary`":{`"R`":178,`"G`":154,`"B`":212,`"A`":255},`"themeSecondary`":{`"R`":147,`"G`":114,`"B`":192,`"A`":255},`"themeDarkAlt`":{`"R`":121,`"G`":89,`"B`":165,`"A`":255},`"themeDark`":{`"R`":102,`"G`":75,`"B`":140,`"A`":255},`"themeDarker`":{`"R`":75,`"G`":56,`"B`":103,`"A`":255},`"accent`":{`"R`":3,`"G`":131,`"B`":135,`"A`":255},`"neutralLighterAlt`":{`"R`":248,`"G`":248,`"B`":248,`"A`":255},`"neutralLighter`":{`"R`":244,`"G`":244,`"B`":244,`"A`":255},`"neutralLight`":{`"R`":234,`"G`":234,`"B`":234,`"A`":255},`"neutralQuaternaryAlt`":{`"R`":218,`"G`":218,`"B`":218,`"A`":255},`"neutralQuaternary`":{`"R`":208,`"G`":208,`"B`":208,`"A`":255},`"neutralTertiaryAlt`":{`"R`":200,`"G`":200,`"B`":200,`"A`":255},`"neutralTertiary`":{`"R`":166,`"G`":166,`"B`":166,`"A`":255},`"neutralSecondary`":{`"R`":102,`"G`":102,`"B`":102,`"A`":255},`"neutralPrimaryAlt`":{`"R`":60,`"G`":60,`"B`":60,`"A`":255},`"neutralPrimary`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255},`"neutralDark`":{`"R`":33,`"G`":33,`"B`":33,`"A`":255},`"black`":{`"R`":0,`"G`":0,`"B`":0,`"A`":255},`"white`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryBackground`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryText`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255}},`"cacheToken`":`"`",`"isDefault`":true,`"version`":`"`"}'}" }
                "Green" { "{'name':'Green','themeJson':'{`"backgroundImageUri`":`"`",`"palette`":{`"themePrimary`":{`"R`":73,`"G`":130,`"B`":5,`"A`":255},`"themeLighterAlt`":{`"R`":246,`"G`":250,`"B`":240,`"A`":255},`"themeLighter`":{`"R`":219,`"G`":235,`"B`":199,`"A`":255},`"themeLight`":{`"R`":189,`"G`":218,`"B`":155,`"A`":255},`"themeTertiary`":{`"R`":133,`"G`":180,`"B`":76,`"A`":255},`"themeSecondary`":{`"R`":90,`"G`":145,`"B`":23,`"A`":255},`"themeDarkAlt`":{`"R`":66,`"G`":117,`"B`":5,`"A`":255},`"themeDark`":{`"R`":56,`"G`":99,`"B`":4,`"A`":255},`"themeDarker`":{`"R`":41,`"G`":73,`"B`":3,`"A`":255},`"accent`":{`"R`":3,`"G`":131,`"B`":135,`"A`":255},`"neutralLighterAlt`":{`"R`":248,`"G`":248,`"B`":248,`"A`":255},`"neutralLighter`":{`"R`":244,`"G`":244,`"B`":244,`"A`":255},`"neutralLight`":{`"R`":234,`"G`":234,`"B`":234,`"A`":255},`"neutralQuaternaryAlt`":{`"R`":218,`"G`":218,`"B`":218,`"A`":255},`"neutralQuaternary`":{`"R`":208,`"G`":208,`"B`":208,`"A`":255},`"neutralTertiaryAlt`":{`"R`":200,`"G`":200,`"B`":200,`"A`":255},`"neutralTertiary`":{`"R`":166,`"G`":166,`"B`":166,`"A`":255},`"neutralSecondary`":{`"R`":102,`"G`":102,`"B`":102,`"A`":255},`"neutralPrimaryAlt`":{`"R`":60,`"G`":60,`"B`":60,`"A`":255},`"neutralPrimary`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255},`"neutralDark`":{`"R`":33,`"G`":33,`"B`":33,`"A`":255},`"black`":{`"R`":0,`"G`":0,`"B`":0,`"A`":255},`"white`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryBackground`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryText`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255}},`"cacheToken`":`"`",`"isDefault`":true,`"version`":`"`"}'}" }
                "Gray" { "{'name':'Gray','themeJson':'{`"backgroundImageUri`":`"`",`"palette`":{`"themePrimary`":{`"R`":105,`"G`":121,`"B`":126,`"A`":255},`"themeLighterAlt`":{`"R`":248,`"G`":249,`"B`":250,`"A`":255},`"themeLighter`":{`"R`":228,`"G`":233,`"B`":234,`"A`":255},`"themeLight`":{`"R`":205,`"G`":213,`"B`":216,`"A`":255},`"themeTertiary`":{`"R`":159,`"G`":173,`"B`":177,`"A`":255},`"themeSecondary`":{`"R`":120,`"G`":136,`"B`":141,`"A`":255},`"themeDarkAlt`":{`"R`":93,`"G`":108,`"B`":112,`"A`":255},`"themeDark`":{`"R`":79,`"G`":91,`"B`":95,`"A`":255},`"themeDarker`":{`"R`":58,`"G`":67,`"B`":70,`"A`":255},`"accent`":{`"R`":0,`"G`":120,`"B`":212,`"A`":255},`"neutralLighterAlt`":{`"R`":248,`"G`":248,`"B`":248,`"A`":255},`"neutralLighter`":{`"R`":244,`"G`":244,`"B`":244,`"A`":255},`"neutralLight`":{`"R`":234,`"G`":234,`"B`":234,`"A`":255},`"neutralQuaternaryAlt`":{`"R`":218,`"G`":218,`"B`":218,`"A`":255},`"neutralQuaternary`":{`"R`":208,`"G`":208,`"B`":208,`"A`":255},`"neutralTertiaryAlt`":{`"R`":200,`"G`":200,`"B`":200,`"A`":255},`"neutralTertiary`":{`"R`":166,`"G`":166,`"B`":166,`"A`":255},`"neutralSecondary`":{`"R`":102,`"G`":102,`"B`":102,`"A`":255},`"neutralPrimaryAlt`":{`"R`":60,`"G`":60,`"B`":60,`"A`":255},`"neutralPrimary`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255},`"neutralDark`":{`"R`":33,`"G`":33,`"B`":33,`"A`":255},`"black`":{`"R`":0,`"G`":0,`"B`":0,`"A`":255},`"white`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryBackground`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryText`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255}},`"cacheToken`":`"`",`"isDefault`":true,`"version`":`"`"}'}" }
                "DarkYellow" { "{'name':'Dark Yellow','themeJson':'{`"backgroundImageUri`":`"`",`"palette`":{`"themePrimary`":{`"R`":255,`"G`":200,`"B`":61,`"A`":255},`"themeLighterAlt`":{`"R`":10,`"G`":8,`"B`":2,`"A`":255},`"themeLighter`":{`"R`":41,`"G`":32,`"B`":10,`"A`":255},`"themeLight`":{`"R`":77,`"G`":60,`"B`":18,`"A`":255},`"themeTertiary`":{`"R`":153,`"G`":120,`"B`":37,`"A`":255},`"themeSecondary`":{`"R`":224,`"G`":176,`"B`":54,`"A`":255},`"themeDarkAlt`":{`"R`":255,`"G`":206,`"B`":81,`"A`":255},`"themeDark`":{`"R`":255,`"G`":213,`"B`":108,`"A`":255},`"themeDarker`":{`"R`":255,`"G`":224,`"B`":146,`"A`":255},`"accent`":{`"R`":255,`"G`":200,`"B`":61,`"A`":255},`"neutralLighterAlt`":{`"R`":40,`"G`":40,`"B`":40,`"A`":255},`"neutralLighter`":{`"R`":49,`"G`":49,`"B`":49,`"A`":255},`"neutralLight`":{`"R`":63,`"G`":63,`"B`":63,`"A`":255},`"neutralQuaternaryAlt`":{`"R`":72,`"G`":72,`"B`":72,`"A`":255},`"neutralQuaternary`":{`"R`":79,`"G`":79,`"B`":79,`"A`":255},`"neutralTertiaryAlt`":{`"R`":109,`"G`":109,`"B`":109,`"A`":255},`"neutralTertiary`":{`"R`":200,`"G`":200,`"B`":200,`"A`":255},`"neutralSecondary`":{`"R`":208,`"G`":208,`"B`":208,`"A`":255},`"neutralPrimaryAlt`":{`"R`":218,`"G`":218,`"B`":218,`"A`":255},`"neutralPrimary`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"neutralDark`":{`"R`":244,`"G`":244,`"B`":244,`"A`":255},`"black`":{`"R`":248,`"G`":248,`"B`":248,`"A`":255},`"white`":{`"R`":31,`"G`":31,`"B`":31,`"A`":255},`"primaryBackground`":{`"R`":31,`"G`":31,`"B`":31,`"A`":255},`"primaryText`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255}},`"cacheToken`":`"`",`"isDefault`":true,`"isInverted`":true,`"version`":`"`"}'}" }
                "DarkBlue" { "{'name':'Dark Blue','themeJson':'{`"backgroundImageUri`":`"`",`"palette`":{`"themePrimary`":{`"R`":58,`"G`":150,`"B`":221,`"A`":255},`"themeLighterAlt`":{`"R`":2,`"G`":6,`"B`":9,`"A`":255},`"themeLighter`":{`"R`":9,`"G`":24,`"B`":35,`"A`":255},`"themeLight`":{`"R`":17,`"G`":45,`"B`":67,`"A`":255},`"themeTertiary`":{`"R`":35,`"G`":90,`"B`":133,`"A`":255},`"themeSecondary`":{`"R`":51,`"G`":133,`"B`":195,`"A`":255},`"themeDarkAlt`":{`"R`":75,`"G`":160,`"B`":225,`"A`":255},`"themeDark`":{`"R`":101,`"G`":174,`"B`":230,`"A`":255},`"themeDarker`":{`"R`":138,`"G`":194,`"B`":236,`"A`":255},`"accent`":{`"R`":58,`"G`":150,`"B`":221,`"A`":255},`"neutralLighterAlt`":{`"R`":29,`"G`":43,`"B`":60,`"A`":255},`"neutralLighter`":{`"R`":34,`"G`":50,`"B`":68,`"A`":255},`"neutralLight`":{`"R`":43,`"G`":61,`"B`":81,`"A`":255},`"neutralQuaternaryAlt`":{`"R`":50,`"G`":68,`"B`":89,`"A`":255},`"neutralQuaternary`":{`"R`":55,`"G`":74,`"B`":95,`"A`":255},`"neutralTertiaryAlt`":{`"R`":79,`"G`":99,`"B`":122,`"A`":255},`"neutralTertiary`":{`"R`":200,`"G`":200,`"B`":200,`"A`":255},`"neutralSecondary`":{`"R`":208,`"G`":208,`"B`":208,`"A`":255},`"neutralPrimaryAlt`":{`"R`":218,`"G`":218,`"B`":218,`"A`":255},`"neutralPrimary`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"neutralDark`":{`"R`":244,`"G`":244,`"B`":244,`"A`":255},`"black`":{`"R`":248,`"G`":248,`"B`":248,`"A`":255},`"white`":{`"R`":24,`"G`":37,`"B`":52,`"A`":255},`"primaryBackground`":{`"R`":24,`"G`":37,`"B`":52,`"A`":255},`"primaryText`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255}},`"cacheToken`":`"`",`"isDefault`":true,`"isInverted`":true,`"version`":`"`"}'}" }
                "Teal" { "{'name':'Teal','themeJson':'{`"backgroundImageUri`":`"`",`"palette`":{`"themePrimary`":{`"R`":3,`"G`":120,`"B`":124,`"A`":255},`"themeLighterAlt`":{`"R`":240,`"G`":249,`"B`":250,`"A`":255},`"themeLighter`":{`"R`":197,`"G`":233,`"B`":234,`"A`":255},`"themeLight`":{`"R`":152,`"G`":214,`"B`":216,`"A`":255},`"themeTertiary`":{`"R`":73,`"G`":174,`"B`":177,`"A`":255},`"themeSecondary`":{`"R`":19,`"G`":137,`"B`":141,`"A`":255},`"themeDarkAlt`":{`"R`":2,`"G`":109,`"B`":112,`"A`":255},`"themeDark`":{`"R`":2,`"G`":92,`"B`":95,`"A`":255},`"themeDarker`":{`"R`":1,`"G`":68,`"B`":70,`"A`":255},`"accent`":{`"R`":79,`"G`":107,`"B`":237,`"A`":255},`"neutralLighterAlt`":{`"R`":248,`"G`":248,`"B`":248,`"A`":255},`"neutralLighter`":{`"R`":244,`"G`":244,`"B`":244,`"A`":255},`"neutralLight`":{`"R`":234,`"G`":234,`"B`":234,`"A`":255},`"neutralQuaternaryAlt`":{`"R`":218,`"G`":218,`"B`":218,`"A`":255},`"neutralQuaternary`":{`"R`":208,`"G`":208,`"B`":208,`"A`":255},`"neutralTertiaryAlt`":{`"R`":200,`"G`":200,`"B`":200,`"A`":255},`"neutralTertiary`":{`"R`":166,`"G`":166,`"B`":166,`"A`":255},`"neutralSecondary`":{`"R`":102,`"G`":102,`"B`":102,`"A`":255},`"neutralPrimaryAlt`":{`"R`":60,`"G`":60,`"B`":60,`"A`":255},`"neutralPrimary`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255},`"neutralDark`":{`"R`":33,`"G`":33,`"B`":33,`"A`":255},`"black`":{`"R`":0,`"G`":0,`"B`":0,`"A`":255},`"white`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryBackground`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryText`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255}},`"cacheToken`":`"`",`"isDefault`":true,`"version`":`"`"}'}" }
                Default { "{'name':'Teal','themeJson':'{`"backgroundImageUri`":`"`",`"palette`":{`"themePrimary`":{`"R`":3,`"G`":120,`"B`":124,`"A`":255},`"themeLighterAlt`":{`"R`":240,`"G`":249,`"B`":250,`"A`":255},`"themeLighter`":{`"R`":197,`"G`":233,`"B`":234,`"A`":255},`"themeLight`":{`"R`":152,`"G`":214,`"B`":216,`"A`":255},`"themeTertiary`":{`"R`":73,`"G`":174,`"B`":177,`"A`":255},`"themeSecondary`":{`"R`":19,`"G`":137,`"B`":141,`"A`":255},`"themeDarkAlt`":{`"R`":2,`"G`":109,`"B`":112,`"A`":255},`"themeDark`":{`"R`":2,`"G`":92,`"B`":95,`"A`":255},`"themeDarker`":{`"R`":1,`"G`":68,`"B`":70,`"A`":255},`"accent`":{`"R`":79,`"G`":107,`"B`":237,`"A`":255},`"neutralLighterAlt`":{`"R`":248,`"G`":248,`"B`":248,`"A`":255},`"neutralLighter`":{`"R`":244,`"G`":244,`"B`":244,`"A`":255},`"neutralLight`":{`"R`":234,`"G`":234,`"B`":234,`"A`":255},`"neutralQuaternaryAlt`":{`"R`":218,`"G`":218,`"B`":218,`"A`":255},`"neutralQuaternary`":{`"R`":208,`"G`":208,`"B`":208,`"A`":255},`"neutralTertiaryAlt`":{`"R`":200,`"G`":200,`"B`":200,`"A`":255},`"neutralTertiary`":{`"R`":166,`"G`":166,`"B`":166,`"A`":255},`"neutralSecondary`":{`"R`":102,`"G`":102,`"B`":102,`"A`":255},`"neutralPrimaryAlt`":{`"R`":60,`"G`":60,`"B`":60,`"A`":255},`"neutralPrimary`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255},`"neutralDark`":{`"R`":33,`"G`":33,`"B`":33,`"A`":255},`"black`":{`"R`":0,`"G`":0,`"B`":0,`"A`":255},`"white`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryBackground`":{`"R`":255,`"G`":255,`"B`":255,`"A`":255},`"primaryText`":{`"R`":51,`"G`":51,`"B`":51,`"A`":255}},`"cacheToken`":`"`",`"isDefault`":true,`"version`":`"`"}'}" } 
            }

             $output = Invoke-PnPSPRestMethod -Method Post `
                 -Url "/_api/ThemeManager/ApplyTheme" `
                 -ContentType "application/json;odata.metadata=none" `
                 -Content $theme -Raw

             if ($output) {
                Write-Host " - Set $($Color) theme"
             }
         }
     }

     End {}
}

function Set-SiteChromeState {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [ValidateSet("Minimal","Compact","Standard","Extended")]
        [string]$HeaderLayout = "Compact",

        [Parameter(Mandatory=$false)]
        [ValidateSet("None","Neutral","Soft","Strong")]
        [string]$HeaderBackground = "Strong",

        [Parameter(Mandatory=$false)]
        [switch]$HideTitle=$false,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Left","Middle","Right")]
        [string]$LogoAlignment = "Left",

        [Parameter(Mandatory=$false)]
        [ValidateSet("Horizontal","Vertical")]
        [string]$Navigation,

        [Parameter(Mandatory=$false)]
        [ValidateSet("MegaMenu","Cascading")]
        [string]$MenuType = "Cascading",

        [Parameter(Mandatory=$false)]
        [bool]$FooterEnabled = $true,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Simple","Extended")]
        [string]$FooterLayout = "Simple",

        [Parameter(Mandatory=$false)]
        [ValidateSet("None","Neutral","Soft","Strong")]
        [string]$FooterBackground = "Strong"
    )

    Begin {
        $web = Invoke-PnPSPRestMethod -Url "/_api/web?`$select=FooterEmphasis,FooterEnabled,FooterLayout,HeaderEmphasis,HeaderLayout,HideTitleInHeader,HorizontalQuickLaunch,LogoAlignment,MegaMenuEnabled,WebTemplateConfiguration"
    }

    Process {

        [int]$HeaderLayout = switch ($HeaderLayout) {
            "Minimal" { 0 }
            "Standard" { 1 }
            "Compact" { 2 }
            "Extended" { 3 }
            Default { $web.HeaderLayout }
        }

        [int]$HeaderBackground = switch ($HeaderBackground) {
            "None" { 0 }
            "Neutral" { 1 }
            "Soft" { 2 }
            "Strong" { 3 }
            Default { $web.HeaderEmphasis }
        }

        [bool]$HideTitle = switch ($HideTitle) {
            $true { $true }
            $false { $false }
            Default { $web.HideTitleInHeader }
        }

        if ($HeaderLayout -eq 3) {
            [int]$LogoAlignment = switch ($LogoAlignment) {
                "Left" { 0 }
                "Middle" { 1 }
                "Right" { 2 }
                Default { $web.LogoAlignment }
            }
        } else {
            [int]$LogoAlignment = $web.LogoAlignment
        }

        [bool]$MenuType = switch ($MenuType) {
            "MegaMenu" { $true }
            "Cascading" { $false }
            Default { $web.MegaMenuEnabled }
        }

        if ($web.WebTemplateConfiguration -eq "SITEPAGEPUBLISHING#0") {
            [bool]$Navigation = $web.HorizontalQuickLaunch

            [bool]$FooterEnabled = switch ($FooterEnabled) {
                $true { $true }
                $false { $false }
                Default { $web.FooterEnabled }
            }

            [int]$FooterLayout = switch ($FooterLayout) {
                "Simple" { 0 }
                "Extended" { 1 }
                Default { $web.FooterLayout }
            }

            [int]$FooterBackground = switch ($FooterBackground) {
                "None" { 3 }
                "Neutral" { 1 }
                "Soft" { 2 }
                "Strong" { 0 }
                Default { $web.FooterEmphasis }
            }

        } else {
            [bool]$Navigation = switch ($Navigation) {
                "Horizontal" { $true }
                "Vertical" { $false }
                Default { $web.HorizontalQuickLaunch }
            }

            [bool]$FooterEnabled = $web.FooterEnabled
            [int]$FooterLayout = $web.FooterLayout
            [int]$FooterBackground = $web.FooterEmphasis
        }


        $settings = @{
            "headerLayout" = $HeaderLayout
            "headerEmphasis" = $HeaderBackground
            "hideTitleInHeader" = $HideTitle
            "logoAlignment" = $LogoAlignment
            "horizontalQuickLaunch" = $Navigation
            "megaMenuEnabled" = $MenuType
            "footerEnabled" = $FooterEnabled
            "footerLayout" = $FooterLayout
            "footerEmphasis" = $FooterBackground
        }

        $output = Invoke-PnPSPRestMethod -Method Post `
            -Url "/_api/web/SetChromeOptions" `
            -ContentType "application/json;odata.metadata=none" `
            -Content $settings
 
        if ($output."odata.null") {
            Write-Host " - Set site look"
        }
    }

    End {}
}

function Set-FooterText {
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Text
    )

    Begin {
        $web = Invoke-PnPSPRestMethod -Url "/_api/web?`$select=FooterEnabled,ServerRelativeUrl,WebTemplateConfiguration"
    }

    Process {
        if ($web.WebTemplateConfiguration -eq "SITEPAGEPUBLISHING#0" -and $web.FooterEnabled) {
            $body = '{
                "menuState": {
                    "StartingNodeTitle": "13b7c916-4fea-4bb2-8994-5cf274aeb530",
                    "SPSitePrefix": "/",
                    "SPWebPrefix": "' + $web.ServerRelativeUrl + '",
                    "FriendlyUrlPrefix": "",
                    "SimpleUrl": "",
                    "Nodes": [
                        {
                            "NodeType": 0,
                            "Title": "7376cd83-67ac-4753-b156-6a7b3fa0fc1f",
                            "FriendlyUrlSegment": "",
                            "Nodes": [
                                {
                                    "NodeType": 0,
                                    "Title": "' + $Text + '",
                                    "FriendlyUrlSegment": ""
                                }
                            ]
                        }
                    ]
                }
            }'

            $output = (Invoke-PnPSPRestMethod -Method Post `
                -Url "/_api/navigation/SaveMenuState" `
                -ContentType "application/json;odata.metadata=none" `
                -Content $body -Raw | ConvertFrom-Json).value

            if ($output -eq 200) {
                Write-Host " - Set footer text to $($Text)"
            }
        }
    }

    End {}
}

function Add-SiteToHub {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidatePattern('^https\:\/\/[a-z]*\.sharepoint\.com\/sites\/etf.*$')]
        [string]$HubSiteUrl
    )

    Begin {
        $site = Invoke-PnPSPRestMethod -Url "/_api/site?`$select=HubSiteId,IsHubSite"
    }

    Process {
        Write-Host "`nASSOCIATING SITE TO HUB"

        if ($site.IsHubSite) {
            Write-Host " - Current site is registered as a Hub site"
            return
        }

        if ($site.IsHubSite -eq $false -and $site.HubSiteId -ne "00000000-0000-0000-0000-000000000000") {
            $hubSite = Invoke-PnPSPRestMethod -Url "/_api/HubSites/GetById?hubSiteId='$($site.HubSiteId)'"

            Write-Host " - Current site is associated to the following hub:"
            Write-Host "   - Title: $($hubSite.Title)"
            Write-Host "   - Url: $($hubSite.SiteUrl)"
            Write-Host "   - Id: $($hubSite.ID)"
            return
        }

        if ($site.IsHubSite -eq $false -and $site.HubSiteId -eq "00000000-0000-0000-0000-000000000000") {
            $parent = Invoke-PnPSPRestMethod -Url "$($HubSiteUrl)/_api/site?`$select=HubSiteId,IsHubSite,ServerRelativeUrl"

            if ($parent.IsHubSite) {
                $association = Invoke-PnPSPRestMethod -Method Post `
                    -Url "/_api/site/JoinHubSite('$($parent.HubSiteId)')" `
                    -ContentType "application/json;odata.metadata=none" `
                    -Content @{}

                if ($association."odata.null") {
                    Write-Host " - Site has been associated to " -NoNewline
                    Write-Host "$($parent.ServerRelativeUrl)" -ForegroundColor Yellow
                }
            }
        }
    }

    End {}
}

try {
    $hostUrl = "$($SiteUrl -replace '(?<=\.com).*$')"
    Connect-PnPOnline -Url $hostUrl -UseWebLogin -WarningAction Ignore -ErrorAction Stop

    $status = Get-SiteStatus -Url $SiteUrl
    $siteAdmin = "ETFDLSPSiteCollAdmins@etf.wi.gov"

    if ($status -eq 0 -and [bool]$SiteType) {
    Write-Host "`nSITESTATUS eq 0"
        $site = New-CreateSite -Url $SiteUrl -Template $SiteType -SiteAdmin $siteAdmin
        $status = $site.SiteStatus
        $SiteUrl = $site.SiteUrl
    }

    if ($status -eq 2) {
        Write-Host "`nCONNECTING TO SITE"
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin -WarningAction Ignore -ErrorAction Stop

        $s = Invoke-PnPSPRestMethod -Url "/_api/site?`$select=ReadOnly"

        if ($s.ReadOnly) {
            Write-Host "`nThis site collection is read-only." -ForegroundColor Red
            return
        }

        $w = Invoke-PnPSPRestMethod -Url "/_api/web?`$select=CurrentUser,Title,Url&`$expand=CurrentUser"

        Write-Host " - Title: $($w.Title)"
        Write-Host " - Url: $($w.Url)"
        Write-Host " - User: $($w.CurrentUser.Email)"

        if ($w.CurrentUser.IsSiteAdmin) {
            Write-Host "`nCONFIGURING SITE"

            # --- Update SCA --- #
            Add-PnPSiteCollectionAdmin -Owners $siteAdmin
            Get-PnPSiteCollectionAdmin | Where-Object Email -NE $siteAdmin | Remove-PnPSiteCollectionAdmin

 #---Update Sharing ---#
            Set-PnPSite -Identity $Url -DefaultLinkToExistingAccess $true 

            # --- Copy logos --- #
            $copyLogoParameters = @{
                SourceUrl = "https://wigov.sharepoint.com/sites/etf-spadmin"
                SourceList = "Site Assets"
                SourceFolder = "SiteLogos"
                TargetUrl = $SiteUrl
                FileNames = @("__rectSitelogo__logo.png", "__sitelogo__thumbnail.png", "__footerlogo__footer.png")
            }

            Copy-SiteLogos @copyLogoParameters

            # --- Add global navigation --- #
            Add-GlobalNavigation

            # --- Update theme --- #
            Set-SiteTheme -Color Teal -ForceUpdate

            # --- Update site header, footer, and navigation --- #
            Set-SiteChromeState

            # --- Set footer text --- #
            Set-FooterText -Text "Wisconsin Department of Employee Trust Funds"

            # --- Set logos --- #
            $setLogoParameters = @{
                SiteLogo = "__rectSitelogo__logo.png"
                ThumbnailLogo = "__sitelogo__thumbnail.png"
                FooterLogo = "__footerlogo__footer.png"
            }

            Set-Logo @setLogoParameters

            # --- Join site to Hub --- #
            if ([bool]$HubUrl) {
                Add-SiteToHub -HubSiteUrl $HubUrl
            }
        } else {
            Write-Host "`nYou need to be a Site Collection Administrator." -ForegroundColor Red
            return
        }
    }
}
catch {
    Write-Host "`nError: $($_.Exception.Message)" -ForegroundColor Red
}
