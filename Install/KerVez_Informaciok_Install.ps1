<#
.SYNOPSIS
  SharePoint Online lista telepito script.

.DESCRIPTION
  Letrehozza a megadott listakat es oszlopaikat PnP.PowerShell-lel.
  Idempotens: ha a lista vagy oszlop mar letezik, nem hozza letre ujra.

.EXAMPLE
  .\KerVezInstall.ps1 -SiteUrl "https://napcsillag.sharepoint.com/sites/DEV-KerVezPortal"

.EXAMPLE
  .\KerVezInstall.ps1 -SiteUrl "https://napcsillag.sharepoint.com/sites/DEV-KerVezPortal" -ClientId "00000000-0000-0000-0000-000000000000"
#>

[CmdletBinding()]
param(
	[Parameter(Mandatory = $true)]
	[string]$SiteUrl = "https://napcsillag.sharepoint.com/sites/DEV-KerVezPortal",

	[Parameter(Mandatory = $false)]
	[string]$ClientId = "00000000-0000-0000-0000-000000000000"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-HashtableValue {
	param(
		[Parameter(Mandatory = $true)]
		[hashtable]$Table,

		[Parameter(Mandatory = $true)]
		[string]$Key,

		[Parameter(Mandatory = $false)]
		$Default = $null
	)

	if ($Table.ContainsKey($Key)) {
		return $Table[$Key]
	}

	return $Default
}

function Ensure-PnPModule {
	if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
		Write-Host "PnP.PowerShell telepitese..."
		Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
	}

	Import-Module PnP.PowerShell -ErrorAction Stop
}

function Ensure-List {
	param(
		[Parameter(Mandatory = $true)]
		[hashtable]$ListDefinition
	)

	$existingList = Get-PnPList -Identity $ListDefinition.Title -ErrorAction SilentlyContinue

	$template = [string](Get-HashtableValue -Table $ListDefinition -Key "Template" -Default "GenericList")
	$onQuickLaunch = [bool](Get-HashtableValue -Table $ListDefinition -Key "OnQuickLaunch" -Default $true)
	$description = [string](Get-HashtableValue -Table $ListDefinition -Key "Description" -Default "")
	$enableFolderCreation = [bool](Get-HashtableValue -Table $ListDefinition -Key "EnableFolderCreation" -Default $false)

	if (-not $existingList) {

		New-PnPList -Title $ListDefinition.Title -Template $template -OnQuickLaunch:$onQuickLaunch -ErrorAction Stop | Out-Null

		if ($description.Length -gt 0) {
			Set-PnPList -Identity $ListDefinition.Title -Description $description -ErrorAction Stop
		}

		if ($enableFolderCreation) {
			Set-PnPList -Identity $ListDefinition.Title -EnableFolderCreation:$true -ErrorAction Stop
		}

		Write-Host "Lista letrehozva: $($ListDefinition.Title)"
	}
	else {
		if ($enableFolderCreation) {
			Set-PnPList -Identity $ListDefinition.Title -EnableFolderCreation:$true -ErrorAction Stop
		}

		Write-Host "Lista mar letezik: $($ListDefinition.Title)"
	}
}

function Get-ChoiceFieldXml {
	param(
		[Parameter(Mandatory = $true)]
		[string]$InternalName,

		[Parameter(Mandatory = $true)]
		[string]$DisplayName,

		[Parameter(Mandatory = $true)]
		[string[]]$Choices,

		[Parameter(Mandatory = $false)]
		[bool]$Multi = $false,

		[Parameter(Mandatory = $false)]
		[bool]$AddToDefaultView = $true
	)

	$choiceXml = ($Choices | ForEach-Object { "<CHOICE>$_</CHOICE>" }) -join ""
	$multAttr = if ($Multi) { "TRUE" } else { "FALSE" }
	$viewAttr = if ($AddToDefaultView) { "TRUE" } else { "FALSE" }

	return "<Field Type='Choice' Name='$InternalName' DisplayName='$DisplayName' Mult='$multAttr' FillInChoice='FALSE' Group='EcoFamily' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' ShowInNewForm='TRUE' ShowInEditForm='TRUE' ShowInDisplayForm='TRUE' Viewable='TRUE' AddToDefaultView='$viewAttr'><CHOICES>$choiceXml</CHOICES><Default>$($Choices[0])</Default></Field>"
}

function Get-UserFieldXml {
	param(
		[Parameter(Mandatory = $true)]
		[string]$InternalName,

		[Parameter(Mandatory = $true)]
		[string]$DisplayName,

		[Parameter(Mandatory = $false)]
		[bool]$Multi = $false,

		[Parameter(Mandatory = $false)]
		[bool]$AddToDefaultView = $true
	)

	$multAttr = if ($Multi) { "TRUE" } else { "FALSE" }
	$viewAttr = if ($AddToDefaultView) { "TRUE" } else { "FALSE" }
	$fieldType = if ($Multi) { "UserMulti" } else { "User" }

	return "<Field Type='$fieldType' Name='$InternalName' DisplayName='$DisplayName' Group='EcoFamily' UserSelectionMode='PeopleOnly' Mult='$multAttr' Required='FALSE' AddToDefaultView='$viewAttr' />"
}

function Get-NoteFieldXml {
	param(
		[Parameter(Mandatory = $true)]
		[string]$InternalName,

		[Parameter(Mandatory = $true)]
		[string]$DisplayName,

		[Parameter(Mandatory = $false)]
		[bool]$RichText = $false,

		[Parameter(Mandatory = $false)]
		[bool]$AddToDefaultView = $true
	)

	$viewAttr = if ($AddToDefaultView) { "TRUE" } else { "FALSE" }
	$richTextAttr = if ($RichText) { "TRUE" } else { "FALSE" }
	$richTextMode = if ($RichText) { "FullHtml" } else { "Compatible" }

	return "<Field Type='Note' Name='$InternalName' DisplayName='$DisplayName' Group='EcoFamily' RichText='$richTextAttr' RichTextMode='$richTextMode' NumLines='8' AppendOnly='FALSE' Required='FALSE' AddToDefaultView='$viewAttr' />"
}

function Get-LookupFieldXml {
	param(
		[Parameter(Mandatory = $true)]
		[string]$InternalName,

		[Parameter(Mandatory = $true)]
		[string]$DisplayName,

		[Parameter(Mandatory = $true)]
		[string]$LookupListTitle,

		[Parameter(Mandatory = $false)]
		[string]$ShowField = "Title",

		[Parameter(Mandatory = $false)]
		[bool]$Multi = $false,

		[Parameter(Mandatory = $false)]
		[bool]$AddToDefaultView = $true
	)

	$lookupList = Get-PnPList -Identity $LookupListTitle -ErrorAction Stop
	$listId = "{$($lookupList.Id.ToString())}"
	$viewAttr = if ($AddToDefaultView) { "TRUE" } else { "FALSE" }

	if ($Multi) {
		return "<Field Type='LookupMulti' Name='$InternalName' DisplayName='$DisplayName' Group='EcoFamily' List='$listId' ShowField='$ShowField' Required='FALSE' Mult='TRUE' AddToDefaultView='$viewAttr' />"
	}

	return "<Field Type='Lookup' Name='$InternalName' DisplayName='$DisplayName' Group='EcoFamily' List='$listId' ShowField='$ShowField' Required='FALSE' AddToDefaultView='$viewAttr' />"
}

function Set-TitleDisplayName {
	param(
		[Parameter(Mandatory = $true)]
		[string]$ListTitle,

		[Parameter(Mandatory = $true)]
		[string]$DisplayName
	)

	Set-PnPField -List $ListTitle -Identity "Title" -Values @{ Title = $DisplayName } -ErrorAction Stop | Out-Null
}

function Ensure-FieldInDefaultView {
	param(
		[Parameter(Mandatory = $true)]
		[string]$ListTitle,

		[Parameter(Mandatory = $true)]
		[string]$InternalName
	)

	$defaultView = Get-PnPView -List $ListTitle -Includes DefaultView, ViewFields -ErrorAction Stop | Where-Object { $_.DefaultView } | Select-Object -First 1
	if (-not $defaultView) {
		Write-Warning "Nem talalhato alapertelmezett nezet a listaban: $ListTitle"
		return
	}

	$currentFields = @($defaultView.ViewFields | ForEach-Object { [string]$_ })
	if ($currentFields -contains $InternalName) {
		return
	}

	$newFields = @($currentFields + $InternalName)
	Set-PnPView -List $ListTitle -Identity $defaultView.Id -Fields $newFields -ErrorAction Stop | Out-Null
}

function Remove-FieldsFromDefaultView {
	param(
		[Parameter(Mandatory = $true)]
		[string]$ListTitle,

		[Parameter(Mandatory = $true)]
		[string[]]$InternalNames
	)

	$defaultView = Get-PnPView -List $ListTitle -Includes DefaultView, ViewFields -ErrorAction Stop | Where-Object { $_.DefaultView } | Select-Object -First 1
	if (-not $defaultView) {
		Write-Warning "Nem talalhato alapertelmezett nezet a listaban: $ListTitle"
		return
	}

	$currentFields = @($defaultView.ViewFields | ForEach-Object { [string]$_ })
	$newFields = @($currentFields | Where-Object { $InternalNames -notcontains $_ })
	if ($newFields.Count -ne $currentFields.Count) {
		Set-PnPView -List $ListTitle -Identity $defaultView.Id -Fields $newFields -ErrorAction Stop | Out-Null
	}
}

function Ensure-Field {
	param(
		[Parameter(Mandatory = $true)]
		[string]$ListTitle,

		[Parameter(Mandatory = $true)]
		[hashtable]$Field
	)

	$internalNameValue = Get-HashtableValue -Table $Field -Key "InternalName" -Default $null
	$nameValue = Get-HashtableValue -Table $Field -Key "Name" -Default $null
	$displayNameValue = Get-HashtableValue -Table $Field -Key "DisplayName" -Default $null

	$internalName = if ($internalNameValue) { [string]$internalNameValue } else { [string]$nameValue }
	$displayName = if ($displayNameValue) { [string]$displayNameValue } else { [string]$nameValue }

	$existingField = Get-PnPField -List $ListTitle -Identity $internalName -ErrorAction SilentlyContinue
	$addToDefaultView = [bool](Get-HashtableValue -Table $Field -Key "AddToDefaultView" -Default $true)

	if ($existingField) {
		$kind = [string](Get-HashtableValue -Table $Field -Key "Kind" -Default "Simple")
		if ($kind -eq "Choice") {
			$existingChoices = @($existingField.Choices | ForEach-Object { [string]$_ })
			$requestedChoices = @($Field.Choices | ForEach-Object { [string]$_ })
			$missingChoices = @($requestedChoices | Where-Object { $existingChoices -notcontains $_ })
			if ($missingChoices.Count -gt 0) {
				$updatedChoices = @($existingChoices + $missingChoices)
				Set-PnPField -List $ListTitle -Identity $internalName -Values @{ Choices = $updatedChoices } -ErrorAction Stop | Out-Null
			}
		}

		if ($addToDefaultView) {
			try {
				Ensure-FieldInDefaultView -ListTitle $ListTitle -InternalName $internalName
			}
			catch {
				Write-Warning "Nem sikerult a default nezethez adni ($ListTitle/$internalName): $($_.Exception.Message)"
			}
		}

		Write-Host "  - Oszlop mar letezik: $internalName"
		return
	}

	$kind = [string](Get-HashtableValue -Table $Field -Key "Kind" -Default "Simple")
	$multi = [bool](Get-HashtableValue -Table $Field -Key "Multi" -Default $false)
	$richText = [bool](Get-HashtableValue -Table $Field -Key "RichText" -Default $false)
	$showField = [string](Get-HashtableValue -Table $Field -Key "ShowField" -Default "Title")

	switch ($kind) {
		"Choice" {
			$xml = Get-ChoiceFieldXml -InternalName $internalName -DisplayName $displayName -Choices $Field.Choices -Multi:$multi -AddToDefaultView:$addToDefaultView
			Add-PnPFieldFromXml -List $ListTitle -FieldXml $xml -ErrorAction Stop | Out-Null
		}
		"User" {
			$xml = Get-UserFieldXml -InternalName $internalName -DisplayName $displayName -Multi:$multi -AddToDefaultView:$addToDefaultView
			Add-PnPFieldFromXml -List $ListTitle -FieldXml $xml -ErrorAction Stop | Out-Null
		}
		"Note" {
			$xml = Get-NoteFieldXml -InternalName $internalName -DisplayName $displayName -RichText:$richText -AddToDefaultView:$addToDefaultView
			Add-PnPFieldFromXml -List $ListTitle -FieldXml $xml -ErrorAction Stop | Out-Null
		}
		"Lookup" {
			$xml = Get-LookupFieldXml -InternalName $internalName -DisplayName $displayName -LookupListTitle $Field.LookupListTitle -ShowField $showField -Multi:$multi -AddToDefaultView:$addToDefaultView
			Add-PnPFieldFromXml -List $ListTitle -FieldXml $xml -ErrorAction Stop | Out-Null
		}
		default {
			$type = [string](Get-HashtableValue -Table $Field -Key "Type" -Default "Text")
			Add-PnPField -List $ListTitle -DisplayName $displayName -InternalName $internalName -Type $type -AddToDefaultView:$addToDefaultView -ErrorAction Stop | Out-Null
		}
	}

	$fieldValues = Get-HashtableValue -Table $Field -Key "Values" -Default $null
	if ($null -ne $fieldValues) {
		Set-PnPField -List $ListTitle -Identity $internalName -Values $fieldValues -ErrorAction Stop | Out-Null
	}

	if ($addToDefaultView) {
		try {
			Ensure-FieldInDefaultView -ListTitle $ListTitle -InternalName $internalName
		}
		catch {
			Write-Warning "Nem sikerult a default nezethez adni ($ListTitle/$internalName): $($_.Exception.Message)"
		}
	}

	Write-Host "  - Oszlop letrehozva: $internalName"
}

$statusChoices = @("Piszkozat", "Közzétéve", "Archive")
$actionTypeChoices = @("EcoFamily napok", "Plakátos akció", "Extra akció", "Kuponos akció", "Akciós újság", "Online újság", "Egyéb akció")

$ListDefinitions = @(
	@{
		Title = "Informaciok"
		Template = "GenericList"
		Description = "Információk listája"
		OnQuickLaunch = $true
		TitleDisplayName = "Információ címe"
		FieldsToRemoveFromDefaultView = @("StartDate", "PublishStartDate")
		Fields = @(
			@{ InternalName = "ActionType"; DisplayName = "Információ típusa"; Kind = "Choice"; Choices = $actionTypeChoices; Multi = $false },
            @{ InternalName = "Status"; DisplayName = "Státusz"; Kind = "Choice"; Choices = $statusChoices; Multi = $false },
            @{ InternalName = "EndDate"; DisplayName = "Információ befejező dátuma"; Kind = "Simple"; Type = "DateTime" },
            @{ InternalName = "Description"; DisplayName = "Információ leírása"; Kind = "Note"; RichText = $true },
            @{ InternalName = "ResponsiblePerson"; DisplayName = "Felelős személy"; Kind = "User"; Multi = $false },
			@{ InternalName = "ResponsibleEmail"; DisplayName = "Felelős email"; Kind = "Simple"; Type = "Text" },
            @{ InternalName = "RelatedDocuments"; DisplayName = "Kapcsolódó dokumentumok"; Kind = "Note"; RichText = $false },
            @{ InternalName = "RelatedShopsType"; DisplayName = "Kapcsolódó boltok típusa"; Kind = "Choice"; Choices = @("all", "group", "selected"); Multi = $false },
            @{ InternalName = "RelatedShops"; DisplayName = "Kapcsolódó boltok"; Kind = "Note"; RichText = $false },
            @{ InternalName = "RelatedProducts"; DisplayName = "Kapcsolódó termékek"; Kind = "Note"; RichText = $false },
            @{ InternalName = "Images"; DisplayName = "Képek"; Kind = "Note"; RichText = $false }
		)
	},
	@{
		Title = "InformacioDokumentumok"
		Template = "DocumentLibrary"
		Description = "Információkhoz feltöltött dokumentumok, képek tárolója"
		OnQuickLaunch = $true
		EnableFolderCreation = $true
		Fields = @(
			@{ InternalName = "ActionId"; DisplayName = "Információ Id"; Type = "Number" },
			@{ InternalName = "UserId"; DisplayName = "Felhasználó Id"; Type = "Number" }
		)
	}
)

Write-Host "Kapcsolodas SharePoint Online-hoz: $SiteUrl"
Ensure-PnPModule
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId | Out-Null

foreach ($listDef in $ListDefinitions) {
	Ensure-List -ListDefinition $listDef

	$titleDisplayName = [string](Get-HashtableValue -Table $listDef -Key "TitleDisplayName" -Default "")
	if (-not [string]::IsNullOrWhiteSpace($titleDisplayName)) {
		Set-TitleDisplayName -ListTitle $listDef.Title -DisplayName $titleDisplayName
	}

	foreach ($fieldDef in $listDef.Fields) {
		Ensure-Field -ListTitle $listDef.Title -Field $fieldDef
	}

	$fieldsToRemoveFromDefaultView = @(Get-HashtableValue -Table $listDef -Key "FieldsToRemoveFromDefaultView" -Default @())
	if ($fieldsToRemoveFromDefaultView.Count -gt 0) {
		Remove-FieldsFromDefaultView -ListTitle $listDef.Title -InternalNames $fieldsToRemoveFromDefaultView
	}
}

Write-Host "Telepites sikeresen lefutott."