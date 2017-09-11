# Copyright © MigrationWiz 2011.  All rights reserved.

$importFile = "DistributionLists.csv"
$ouName = "Imported Distribution Lists"

&{
	Write-Host -Foreground White "Locating distribution lists OU" $ouName "..."
	$ldap = "LDAP://RootDse"
	$rootDse = [adsi] $ldap
	$defaultNamingContext = $rootDse.defaultNamingContext
	$ldap = "LDAP://OU=" + $ouName + "," + $defaultNamingContext
	$ou = [adsi] $ldap
	if($ou.Name -ne $null)
	{
		Write-Host -Foreground Yellow "  Distribution lists OU already exists"
	}
	else
	{
		Write-Host -Foreground White "  Creating distribution lists OU"

		$ldap = "LDAP://" + $defaultNamingContext
		$domainRoot = [adsi] $ldap
		$ou = $domainRoot.Create("OrganizationalUnit", "OU=" + $ouName)
		$ou.SetInfo()
	}

    $distributionLists = import-csv $importFile | select *
    foreach($distributionList in $distributionLists)
    {
        $displayName			= $distributionList.'displayName'
        $displayNamePrintable	= $distributionList.'displayNamePrintable'
        $extensionAttribute1	= $distributionList.'extensionAttribute1'
        $extensionAttribute2	= $distributionList.'extensionAttribute2'
        $extensionAttribute3	= $distributionList.'extensionAttribute3'
        $extensionAttribute4	= $distributionList.'extensionAttribute4'
        $extensionAttribute5	= $distributionList.'extensionAttribute5'
        $extensionAttribute6	= $distributionList.'extensionAttribute6'
        $extensionAttribute7	= $distributionList.'extensionAttribute7'
        $extensionAttribute8	= $distributionList.'extensionAttribute8'
        $extensionAttribute9	= $distributionList.'extensionAttribute9'
        $extensionAttribute10	= $distributionList.'extensionAttribute10'
        $extensionAttribute11	= $distributionList.'extensionAttribute11'
        $extensionAttribute12	= $distributionList.'extensionAttribute12'
        $extensionAttribute13	= $distributionList.'extensionAttribute13'
        $extensionAttribute14	= $distributionList.'extensionAttribute14'
        $extensionAttribute15	= $distributionList.'extensionAttribute15'
        $info					= $distributionList.'info'
        $mailNickname			= $distributionList.'mailNickname'
        $memberEmail			= $distributionList.'memberEmail'
        $samAccountName			= $distributionList.'sAMAccountName'
       
        Write-Host -Foreground White "Importing distribution list" $displayName "..."

		if($displayName -eq "") { $displayName = $null }
		if($displayNamePrintable -eq "") { $displayNamePrintable = $null }
		if($extensionAttribute1 -eq "") { $extensionAttribute1 = $null }
		if($extensionAttribute2 -eq "") { $extensionAttribute2 = $null }
		if($extensionAttribute3 -eq "") { $extensionAttribute3 = $null }
		if($extensionAttribute4 -eq "") { $extensionAttribute4 = $null }
		if($extensionAttribute5 -eq "") { $extensionAttribute5 = $null }
		if($extensionAttribute6 -eq "") { $extensionAttribute6 = $null }
		if($extensionAttribute7 -eq "") { $extensionAttribute7 = $null }
		if($extensionAttribute8 -eq "") { $extensionAttribute8 = $null }
		if($extensionAttribute9 -eq "") { $extensionAttribute9 = $null }
		if($extensionAttribute10 -eq "") { $extensionAttribute10 = $null }
		if($extensionAttribute11 -eq "") { $extensionAttribute11 = $null }
		if($extensionAttribute12 -eq "") { $extensionAttribute12 = $null }
		if($extensionAttribute13 -eq "") { $extensionAttribute13 = $null }
		if($extensionAttribute14 -eq "") { $extensionAttribute14 = $null }
		if($extensionAttribute15 -eq "") { $extensionAttribute15 = $null }
		if($info -eq "") { $info = $null }
		if($mailNickname -eq "") { $mailNickname = $null }
		if($memberEmail -eq "") { $memberEmail = $null }
		if($samAccountName -eq "") { $samAccountName = $null }

		$dl = Get-DistributionGroup $mailNickname -ErrorAction SilentlyContinue
		if($dl)
		{
			Write-Host -Foreground Yellow "  Distribution list already exists"
			Set-DistributionGroup -Identity $mailNickname -Name $displayName -Alias $mailNickname -DisplayName $displayName -SamAccountName $samAccountName
		}
		else
		{
			Write-Host -Foreground White "  Creating distribution list"
			$dl = New-DistributionGroup -OrganizationalUnit $ou.distinguishedName.ToString() -Name $displayName -Alias $mailNickname -DisplayName $displayName -SamAccountName $samAccountName -Type Distribution
		}
		Set-DistributionGroup -Identity $mailNickname -CustomAttribute1 $extensionAttribute1 -CustomAttribute2 $extensionAttribute2 -CustomAttribute3 $extensionAttribute3 -CustomAttribute4 $extensionAttribute4 -CustomAttribute5 $extensionAttribute5 -CustomAttribute6 $extensionAttribute6 -CustomAttribute7 $extensionAttribute7 -CustomAttribute8 $extensionAttribute8 -CustomAttribute9 $extensionAttribute9 -CustomAttribute10 $extensionAttribute10 -CustomAttribute11 $extensionAttribute11 -CustomAttribute12 $extensionAttribute12 -CustomAttribute13 $extensionAttribute13 -CustomAttribute14 $extensionAttribute14 -CustomAttribute15 $extensionAttribute15 -SimpleDisplayName $displayNamePrintable
		Set-Group -Identity $mailNickname -Notes $info

		$members = $memberEmail.Split(";")
		foreach($member in $members)
		{
            Write-Host -Foreground White "  Adding member" $member "..."

			$existingMemberFound = $false
			$existingMembers = Get-DistributionGroupMember -Identity $mailNickName
			foreach($existingMember in $existingMembers)
			{
				if($existingMember.EmailAddresses -and $existingMember.EmailAddresses.Contains($member))
				{
					Write-Host -Foreground Yellow "    Already a member of distribution list"
					$existingMemberFound = $true
					break
				}
			}

			if(!$existingMemberFound)
			{
				Write-Host -Foreground White "    Adding as member of distribution list"
				Add-DistributionGroupMember -Identity $mailNickName -Member $member
			}
		}

        Write-Host ""
    }
}
trap
{
    break
}

# SIG # Begin signature block
# MIIV5QYJKoZIhvcNAQcCoIIV1jCCFdICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUexcHp0lUA07uuDnUshTmUEL4
# cfigghFUMIIDejCCAmKgAwIBAgIQOCXX+vhhr570kOcmtdZa1TANBgkqhkiG9w0B
# AQUFADBTMQswCQYDVQQGEwJVUzEXMBUGA1UEChMOVmVyaVNpZ24sIEluYy4xKzAp
# BgNVBAMTIlZlcmlTaWduIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EwHhcNMDcw
# NjE1MDAwMDAwWhcNMTIwNjE0MjM1OTU5WjBcMQswCQYDVQQGEwJVUzEXMBUGA1UE
# ChMOVmVyaVNpZ24sIEluYy4xNDAyBgNVBAMTK1ZlcmlTaWduIFRpbWUgU3RhbXBp
# bmcgU2VydmljZXMgU2lnbmVyIC0gRzIwgZ8wDQYJKoZIhvcNAQEBBQADgY0AMIGJ
# AoGBAMS18lIVvIiGYCkWSlsvS5Frh5HzNVRYNerRNl5iTVJRNHHCe2YdicjdKsRq
# CvY32Zh0kfaSrrC1dpbxqUpjRUcuawuSTksrjO5YSovUB+QaLPiCqljZzULzLcB1
# 3o2rx44dmmxMCJUe3tvvZ+FywknCnmA84eK+FqNjeGkUe60tAgMBAAGjgcQwgcEw
# NAYIKwYBBQUHAQEEKDAmMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC52ZXJpc2ln
# bi5jb20wDAYDVR0TAQH/BAIwADAzBgNVHR8ELDAqMCigJqAkhiJodHRwOi8vY3Js
# LnZlcmlzaWduLmNvbS90c3MtY2EuY3JsMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMI
# MA4GA1UdDwEB/wQEAwIGwDAeBgNVHREEFzAVpBMwETEPMA0GA1UEAxMGVFNBMS0y
# MA0GCSqGSIb3DQEBBQUAA4IBAQBQxUvIJIDf5A0kwt4asaECoaaCLQyDFYE3CoIO
# LLBaF2G12AX+iNvxkZGzVhpApuuSvjg5sHU2dDqYT+Q3upmJypVCHbC5x6CNV+D6
# 1WQEQjVOAdEzohfITaonx/LhhkwCOE2DeMb8U+Dr4AaH3aSWnl4MmOKlvr+ChcNg
# 4d+tKNjHpUtk2scbW72sOQjVOCKhM4sviprrvAchP0RBCQe1ZRwkvEjTRIDroc/J
# ArQUz1THFqOAXPl5Pl1yfYgXnixDospTzn099io6uE+UAKVtCoNd+V5T9BizVw9w
# w/v1rZWgDhfexBaAYMkPK26GBPHr9Hgn0QXF7jRbXrlJMvIzMIIDxDCCAy2gAwIB
# AgIQR78Zld+NUkZD99ttSA0xpDANBgkqhkiG9w0BAQUFADCBizELMAkGA1UEBhMC
# WkExFTATBgNVBAgTDFdlc3Rlcm4gQ2FwZTEUMBIGA1UEBxMLRHVyYmFudmlsbGUx
# DzANBgNVBAoTBlRoYXd0ZTEdMBsGA1UECxMUVGhhd3RlIENlcnRpZmljYXRpb24x
# HzAdBgNVBAMTFlRoYXd0ZSBUaW1lc3RhbXBpbmcgQ0EwHhcNMDMxMjA0MDAwMDAw
# WhcNMTMxMjAzMjM1OTU5WjBTMQswCQYDVQQGEwJVUzEXMBUGA1UEChMOVmVyaVNp
# Z24sIEluYy4xKzApBgNVBAMTIlZlcmlTaWduIFRpbWUgU3RhbXBpbmcgU2Vydmlj
# ZXMgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCpyrKkzM0grwp9
# iayHdfC0TvHfwQ+/Z2G9o2Qc2rv5yjOrhDCJWH6M22vdNp4Pv9HsePJ3pn5vPL+T
# rw26aPRslMq9Ui2rSD31ttVdXxsCn/ovax6k96OaphrIAuF/TFLjDmDsQBx+uQ3e
# P8e034e9X3pqMS4DmYETqEcgzjFzDVctzXg0M5USmRK53mgvqubjwoqMKsOLIYdm
# vYNYV291vzyqJoddyhAVPJ+E6lTBCm7E/sVK3bkHEZcifNs+J9EeeOyfMcnx5iIZ
# 28SzR0OaGl+gHpDkXvXufPF9q2IBj/VNC97QIlaolc2uiHau7roN8+RN2aD7aKCu
# FDuzh8G7AgMBAAGjgdswgdgwNAYIKwYBBQUHAQEEKDAmMCQGCCsGAQUFBzABhhho
# dHRwOi8vb2NzcC52ZXJpc2lnbi5jb20wEgYDVR0TAQH/BAgwBgEB/wIBADBBBgNV
# HR8EOjA4MDagNKAyhjBodHRwOi8vY3JsLnZlcmlzaWduLmNvbS9UaGF3dGVUaW1l
# c3RhbXBpbmdDQS5jcmwwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQD
# AgEGMCQGA1UdEQQdMBukGTAXMRUwEwYDVQQDEwxUU0EyMDQ4LTEtNTMwDQYJKoZI
# hvcNAQEFBQADgYEASmv56ljCRBwxiXmZK5a/gqwB1hxMzbCKWG7fCCmjXsjKkxPn
# BFIN70cnLwA4sOTJk06a1CJiFfc/NyFPcDGA8Ys4h7Po6JcA/s9Vlk4k0qknTnqu
# t2FB8yrO58nZXt27K4U+tZ212eFX/760xX71zwye8Jf+K9M7UhsbOCf3P0owggTe
# MIIDxqADAgECAgIDATANBgkqhkiG9w0BAQUFADBjMQswCQYDVQQGEwJVUzEhMB8G
# A1UEChMYVGhlIEdvIERhZGR5IEdyb3VwLCBJbmMuMTEwLwYDVQQLEyhHbyBEYWRk
# eSBDbGFzcyAyIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTA2MTExNjAxNTQz
# N1oXDTI2MTExNjAxNTQzN1owgcoxCzAJBgNVBAYTAlVTMRAwDgYDVQQIEwdBcml6
# b25hMRMwEQYDVQQHEwpTY290dHNkYWxlMRowGAYDVQQKExFHb0RhZGR5LmNvbSwg
# SW5jLjEzMDEGA1UECxMqaHR0cDovL2NlcnRpZmljYXRlcy5nb2RhZGR5LmNvbS9y
# ZXBvc2l0b3J5MTAwLgYDVQQDEydHbyBEYWRkeSBTZWN1cmUgQ2VydGlmaWNhdGlv
# biBBdXRob3JpdHkxETAPBgNVBAUTCDA3OTY5Mjg3MIIBIjANBgkqhkiG9w0BAQEF
# AAOCAQ8AMIIBCgKCAQEAxC3VFYycJkzsMjXrX7hZAVqmYYFZO3Bjq+PcPccquMkz
# 03nkOu08MCOEjrMwFLayh8M9lVQEnt+Z3QslHiHeZSl+NaipVOv29zI51CZVla3v
# +/5Yhtee9ACNjCoMvUIEzqc/BPbugPKq71KhaWbavhqtXdosZuoaa7vlGlFKAC9I
# x5h12LkpyO74Zm0KnLPz/Hh8ovij8rXD87l6kcGn5iUunKjtEmVuavYSRFNwMJXD
# nCtYKz0IdEryvlGwv4fQTCdYa7U1xZ2vFzH4C4/urYE2BYkImM86ryWHwEnqp/1n
# 90WOl8wUOeI2hbV+Gjf9FvZxEZp0MBb+E5SjP4QNTwIDAQABo4IBMjCCAS4wHQYD
# VR0OBBYEFP2sYTKTbEXW4u6FX5q653aZaMznMB8GA1UdIwQYMBaAFNLEsNKR1EwR
# cbNhyz2h/t2oatTjMBIGA1UdEwEB/wQIMAYBAf8CAQAwMwYIKwYBBQUHAQEEJzAl
# MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5nb2RhZGR5LmNvbTBGBgNVHR8EPzA9
# MDugOaA3hjVodHRwOi8vY2VydGlmaWNhdGVzLmdvZGFkZHkuY29tL3JlcG9zaXRv
# cnkvZ2Ryb290LmNybDBLBgNVHSAERDBCMEAGBFUdIAAwODA2BggrBgEFBQcCARYq
# aHR0cDovL2NlcnRpZmljYXRlcy5nb2RhZGR5LmNvbS9yZXBvc2l0b3J5MA4GA1Ud
# DwEB/wQEAwIBBjANBgkqhkiG9w0BAQUFAAOCAQEA0obA7L35obZn7mYLogY6BFCO
# FXKsSnSVU8s3y0RJ7weQazPZlvCUVqUTMAU8hTIhe8nHCqgkpJDeRtMlIxQDZ8IQ
# 1m8PXXt6zJ/FWCrBxJ4hqFrzrKRG857kY8svkKQpKQHZciwp3zcBJ7xP7mjTIY/A
# s+T1Ce3SEKpTtL7wzFkL1juWHJUkSd/O7P2nSJEURQ46Nm/aRbNFokHJ1NdETj65
# dHbVohNVLMaHo7WZrAaEh391Bvy/FEwOzG7E3z23EnH06PFRQCIoSeAdS4eoNMwG
# ot0SWtGGNmQDNW9vd27r8oVQmF6rA1OtkSNjHxaczbmyBWM64fRoGxcFNZVT7jCC
# BSgwggQQoAMCAQICB0tOBIdsHVIwDQYJKoZIhvcNAQEFBQAwgcoxCzAJBgNVBAYT
# AlVTMRAwDgYDVQQIEwdBcml6b25hMRMwEQYDVQQHEwpTY290dHNkYWxlMRowGAYD
# VQQKExFHb0RhZGR5LmNvbSwgSW5jLjEzMDEGA1UECxMqaHR0cDovL2NlcnRpZmlj
# YXRlcy5nb2RhZGR5LmNvbS9yZXBvc2l0b3J5MTAwLgYDVQQDEydHbyBEYWRkeSBT
# ZWN1cmUgQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkxETAPBgNVBAUTCDA3OTY5Mjg3
# MB4XDTExMDQxNjAxMDM1N1oXDTEyMDQxNTIyMDMyNVowWjELMAkGA1UEBgwCVVMx
# CzAJBgNVBAgMAldBMRAwDgYDVQQHDAdSZWRtb25kMRUwEwYDVQQKDAxNaWdyYXRp
# b25XaXoxFTATBgNVBAMMDE1pZ3JhdGlvbldpejCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBALu1QUd0tnx4v+24CeJjL3zoNzpZkT3bxFPpxzWQvRL4QdnN
# m1jwsV1byxvA8+78n7Erl/ow20Wypy36qUUKC2b1fIzdBzosMEboGpxvPNOkG5by
# RmIpcCFagi6yPPJAGlyMdNpswqOwTIMX25x8kki3UgfQ8JnP7yH4oTBi0a4PJNU5
# HsQo4HMOCOOr5v8cpz8vG+kV3lv8K7jA/jZ2fes5SUCrLKoDc3tZKBHziqiTacBB
# BthcMhniGvoZCvQeTdipyTfSEWh9vIVBQDYnK+0x0HtazTQ8E6XpvtcnteJ4E+D1
# LZzT6TyphSBcp4Un/ZUcq4oi+5GMh8qJ/xcEAv8CAwEAAaOCAYAwggF8MA8GA1Ud
# EwEB/wQFMAMBAQAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwDgYDVR0PAQH/BAQDAgeA
# MDIGA1UdHwQrMCkwJ6AloCOGIWh0dHA6Ly9jcmwuZ29kYWRkeS5jb20vZ2RzMi0x
# LmNybDBNBgNVHSAERjBEMEIGC2CGSAGG/W0BBxcCMDMwMQYIKwYBBQUHAgEWJWh0
# dHBzOi8vY2VydHMuZ29kYWRkeS5jb20vcmVwb3NpdG9yeS8wgYAGCCsGAQUFBwEB
# BHQwcjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZ29kYWRkeS5jb20vMEoGCCsG
# AQUFBzAChj5odHRwOi8vY2VydGlmaWNhdGVzLmdvZGFkZHkuY29tL3JlcG9zaXRv
# cnkvZ2RfaW50ZXJtZWRpYXRlLmNydDAfBgNVHSMEGDAWgBT9rGEyk2xF1uLuhV+a
# uud2mWjM5zAdBgNVHQ4EFgQUwyXEoJ7sui5jHFfrgiAbPaPAQSswDQYJKoZIhvcN
# AQEFBQADggEBAEzKFO1qIK4wwGzdbFIpXw3usMUbCUEqNiFxbQBfa6kcpxFNqWRN
# h0tNajUvb7qNe1vFDDf1uuGoazGY73b161SUL6BK/o0o2iHNu3y2ucZTsTlykaiB
# fPgR2qtSOxYCEEgTdmR+/1Trf3IhZwZJ3ZLKOA8eXjQmStOeK1Ap9ywL7UEVP2Ah
# DA98dymm0BEcWJMfDNU+mzhoxTBHXAXFsHWEn1jTbadCyISdsXzdKsFRNCliC1px
# /n2kJUMAloXKCIK2DKSXDew8o4zsT62SSBP+A/si46WhrWMv618uNkmULQEcIYMj
# OTYxb6aTpXnkainKh4ZjcLkMEovn/ZirmWYxggP7MIID9wIBATCB1jCByjELMAkG
# A1UEBhMCVVMxEDAOBgNVBAgTB0FyaXpvbmExEzARBgNVBAcTClNjb3R0c2RhbGUx
# GjAYBgNVBAoTEUdvRGFkZHkuY29tLCBJbmMuMTMwMQYDVQQLEypodHRwOi8vY2Vy
# dGlmaWNhdGVzLmdvZGFkZHkuY29tL3JlcG9zaXRvcnkxMDAuBgNVBAMTJ0dvIERh
# ZGR5IFNlY3VyZSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTERMA8GA1UEBRMIMDc5
# NjkyODcCB0tOBIdsHVIwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAIoAKA
# AKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFGAegN18JNJwjn0/+TPbE6sh
# FTlfMA0GCSqGSIb3DQEBAQUABIIBAFVvdTRIjDQQYNpfix48CMYtTabFapjZH+iy
# nKNLe2xzGG/KmfpUPLM/oggB/uFR41609CHxIy+Cus8ZZLCgwfnufcK/kiIeuMo3
# CGEKpaSKpBNegcTYE7ZR7zVt9CKs1hJe8nnSskhrvWf2Y8MDotNiF8Mq80kmgcfU
# 2BT7r87MxnwFMHeaDmk9wLo/vSfKo6lt5slYN9dtBLsDxMlSDeTNBWDkcfvLTCMw
# nDM6YWUUy+C427Ly5PQWlG4Ztdh87NDxwdDZQG8CVDn8CSOwsRTULj96Nb35v9Xx
# kGLfrLHpNfUb/ugV6FCpGcmlof4YDU4PuNxPiPiQKvrhUnwn9IKhggF/MIIBewYJ
# KoZIhvcNAQkGMYIBbDCCAWgCAQEwZzBTMQswCQYDVQQGEwJVUzEXMBUGA1UEChMO
# VmVyaVNpZ24sIEluYy4xKzApBgNVBAMTIlZlcmlTaWduIFRpbWUgU3RhbXBpbmcg
# U2VydmljZXMgQ0ECEDgl1/r4Ya+e9JDnJrXWWtUwCQYFKw4DAhoFAKBdMBgGCSqG
# SIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTExMDcxNzAzNDQz
# MVowIwYJKoZIhvcNAQkEMRYEFMvXJ0ABh0YQuAiohe3865RYKMJzMA0GCSqGSIb3
# DQEBAQUABIGACagQjNG+I1tztF9t2vtlDfgPsUiuh44Pk3NbUhTvPxUHyth4A4Nd
# skbvAiTueHhbqlmz3F4lFN2OcujhGqbdg8e6tcpQ9KL1UtPCMPGSflFMUEngcE6Q
# pjch0Z88Fn3eb9xPWysltTjMKEmEVSzXrflWye8zgg8hntaIGBsNMa4=
# SIG # End signature block
