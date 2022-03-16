<#PSScriptInfo

.VERSION 1.0.0

.GUID 14e2901d-0572-4884-af58-90f053f85a66

.AUTHOR Jack Terry

.COMPANYNAME Winupgrade Limited

.COPYRIGHT Winupgrade Limited

.TAGS 

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES


#>

#Requires -Module MSOnline

<# 

.DESCRIPTION Run this script to export the metadata required to refresh the Power BI Office 365 Licensing Report...

#> 
param()

Write-Host "`nThis script will export your Office 365 Licensing metadata to C:\Licensing Report\ and overwrite any exsiting files...`n"
Pause

#Make working directory...
New-Item -Path "C:\" -Name "Licensing Report" -ItemType "directory" -Force

#Connect to MsOnline...
Write-Host "`nPlease enter the credentials of an account with the Global Reader role assigned and activated...`n"
Connect-MsolService

#Display and select Account SKUs...
Get-MsolAccountSku|Select-Object -Property AccountSkuId|Write-Host
$Sku = Read-Host -Prompt "`nPlease enter the AccountSkuId you would like to report on...`n"

#Get license assignment groups...
$Groups = Get-MsolUser -All|Select-Object -ExpandProperty Licenses|?{
$_.AccountSkuId -eq $Sku}|Select-Object -ExpandProperty GroupsAssigningLicense|%{Get-MsolGroup -ObjectId $_.Guid -ErrorAction SilentlyContinue}

#Export directly assigned users...
Get-MsolUser -All|?{($_.Licenses|?{
$_.AccountSkuId -eq $Sku}).GroupsAssigningLicense.Guid -eq $_.ObjectId}|Select-Object -Property UserPrincipalName|Export-Csv -Path "C:\Licensing Report\Direct.csv" -Force -NoTypeInformation

#Export users assigned by group...
foreach($Group in $Groups){$Id=$Group.ObjectId;Get-MsolUser -All|?{($_.Licenses|?{
$_.AccountSkuId -eq $Sku}).GroupsAssigningLicense.Guid -eq $Id}|Select-Object -Property UserPrincipalName|Export-Csv -Path "C:\Licensing Report\$Id.csv" -Force -NoTypeInformation}

#Contact: info@winupgrade.co.uk
Write-Host "`nContact: info@winupgrade.co.uk`n"
Pause