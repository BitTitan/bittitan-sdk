# Copyright © MigrationWiz 2011.  All rights reserved.

$exportFile = "Contacts.csv"

&{
    $csv = 'c,company,department,displayName,displayNamePrintable,extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15,facsimileTelephoneNumber,givenName,homePhone,initials,info,l,mailNickname,mobile,otherFacsimileTelephoneNumber,otherHomePhone,otherTelephone,pager,physicalDeliveryOfficeName,postalCode,postOfficeBox,sn,st,streetAddress,targetAddress,telephoneNumber,title,wWWHomePage'

	$contacts = Get-MailContact
    foreach($contact in $contacts)
    {        
		Write-Host -Foreground White "Exporting contact" $contact.DisplayName "..."

		$contact2 = Get-Contact -Identity $contact.Identity
		
        $csv += "`r`n"
		$csv += '"' + $contact2.CountryOrRegion + '"' + ','							# c
        $csv += '"' + $contact2.Company + '"' + ','									# company
        $csv += '"' + $contact2.Department + '"' + ','								# department
        $csv += '"' + $contact.DisplayName + '"' + ','								# displayName
        $csv += '"' + $contact2.SimpleDisplayName + '"' + ','						# displayNamePrintable
        $csv += '"' + $contact.CustomAttribute1 + '"' + ','							# extensionAttribute1
        $csv += '"' + $contact.CustomAttribute2 + '"' + ','							# extensionAttribute2
        $csv += '"' + $contact.CustomAttribute3 + '"' + ','							# extensionAttribute3
        $csv += '"' + $contact.CustomAttribute4 + '"' + ','							# extensionAttribute4
        $csv += '"' + $contact.CustomAttribute5 + '"' + ','							# extensionAttribute5
        $csv += '"' + $contact.CustomAttribute6 + '"' + ','							# extensionAttribute6
        $csv += '"' + $contact.CustomAttribute7 + '"' + ','							# extensionAttribute7
        $csv += '"' + $contact.CustomAttribute8 + '"' + ','							# extensionAttribute8
        $csv += '"' + $contact.CustomAttribute9 + '"' + ','							# extensionAttribute9
        $csv += '"' + $contact.CustomAttribute10 + '"' + ','						# extensionAttribute10
        $csv += '"' + $contact.CustomAttribute11 + '"' + ','						# extensionAttribute11
        $csv += '"' + $contact.CustomAttribute12 + '"' + ','						# extensionAttribute12
        $csv += '"' + $contact.CustomAttribute13 + '"' + ','						# extensionAttribute13
        $csv += '"' + $contact.CustomAttribute14 + '"' + ','						# extensionAttribute14
        $csv += '"' + $contact.CustomAttribute15 + '"' + ','						# extensionAttribute15
        $csv += '"' + $contact2.Fax + '"' + ','										# facsimileTelephoneNumber
        $csv += '"' + $contact2.FirstName + '"' + ','								# givenName
        $csv += '"' + $contact2.HomePhone + '"' + ','								# homePhone
        $csv += '"' + $contact2.Initials + '"' + ','								# initials
        $csv += '"' + $contact2.Notes + '"' + ','									# info
        $csv += '"' + $contact2.City + '"' + ','									# l
        $csv += '"' + $contact.Alias + '"' + ','									# mailNickname
        $csv += '"' + $contact2.MobilePhone + '"' + ','								# mobile
        $csv += '"' + [string]::join(';', $contact2.OtherFax) + '"' + ','			# otherFacsimileTelephoneNumber
        $csv += '"' + [string]::join(';', $contact2.OtherHomePhone) + '"' + ','		# otherHomePhone
        $csv += '"' + [string]::join(';', $contact2.OtherTelephone) + '"' + ','		# otherTelephone
        $csv += '"' + $contact2.Pager + '"' + ','									# pager
		$csv += '"' + $contact2.Office + '"' + ','									# physicalDeliveryOfficeName
		$csv += '"' + $contact2.PostalCode + '"' + ','								# postalCode
		$csv += '"' + $contact2.PostOfficeBox + '"' + ','							# postOfficeBox
        $csv += '"' + $contact2.LastName + '"' + ','								# sn
        $csv += '"' + $contact2.StateOrProvince + '"' + ','							# st
        $csv += '"' + $contact2.StreetAddress + '"' + ','							# streetAddress
        $csv += '"' + $contact.ExternalEmailAddress + '"' + ','						# targetAddress
        $csv += '"' + $contact2.Phone + '"' + ','									# telephoneNumber
		$csv += '"' + $contact2.Title + '"' + ','									# title
		$csv += '"' + $contact2.WebPage + '"' 										# wWWHomePage
	}

    $file = New-Item $exportFile -type file -force -value $csv
}
trap
{
    break;
}


# SIG # Begin signature block
# MIIV5QYJKoZIhvcNAQcCoIIV1jCCFdICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUuNusrtIm8pOshiuupihl/d7A
# idOgghFUMIIDejCCAmKgAwIBAgIQOCXX+vhhr570kOcmtdZa1TANBgkqhkiG9w0B
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
# MAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFHT6MZr97Jw+YcItLqlFZKzE
# dgSOMA0GCSqGSIb3DQEBAQUABIIBAJ4Z5/0Ag45Bz3DmOEoxUftNlfR5z4AS8PDi
# IlzeEBT3gyWmzZkG0/0C/v3t8QNY9oSESwOu5+7cBzgxh0NMLNGCYjx0uKixU4U5
# iD6PyfOJ0AzBDhlm457vvWDb2VPwe3EFQDdbIJImlz+OE5gWxtb52naR52NUIERK
# PMmDfh8wqSDWpasZz89mr3TShPpCOK3kI4GvTyInl1HonMzhzsyCRcFxx70URhLY
# SVDdeN74CFfCw+cNRAunB68FOa195fDwpNx/4j3+xJeOGpMsjS9rCppwUktE+HER
# j5RSL2fhafEJzIKPxp9EZ0c7BzvFHwRPHpl/thseittDXrrNmumhggF/MIIBewYJ
# KoZIhvcNAQkGMYIBbDCCAWgCAQEwZzBTMQswCQYDVQQGEwJVUzEXMBUGA1UEChMO
# VmVyaVNpZ24sIEluYy4xKzApBgNVBAMTIlZlcmlTaWduIFRpbWUgU3RhbXBpbmcg
# U2VydmljZXMgQ0ECEDgl1/r4Ya+e9JDnJrXWWtUwCQYFKw4DAhoFAKBdMBgGCSqG
# SIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTExMDcxNzAzNDQz
# MFowIwYJKoZIhvcNAQkEMRYEFJgsY1woX2lU+MuRYjgKKdrYX/12MA0GCSqGSIb3
# DQEBAQUABIGAep/KXVKhvR+DzKnFOuVUYlg+K22BVt5xTYyxdiQ4QUo0C6wAytzS
# Hj2Pwi5INm9u5eYAnYSvSNS/x9ltuIU6/ddktSWC+u+NFsUdD1kdk15Fc7kL8vz1
# MfKEFOixAOuIb7uHDhWrjNAScN1fYQqmv+JAVMLFBaS2DacDYMLfQy8=
# SIG # End signature block
