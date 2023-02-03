class GetOutlookAccounts {

	[void]run() {
		$accounts = $this.getAccounts()
		Write-Host $accounts
		$userName = $this.getUser();
	}

	[System.Object[]]getAccounts() {
		$outlook = New-Object -ComObject Outlook.Application
		$ns = $outlook.GetNamespace("MAPI")
		$accounts = $ns.Folders | Select-Object -ExpandProperty Name
		$collection = @()

		$accounts | Where-Object { $_ -Match '@inotech.no$' } | ForEach-Object {
			if ($collection -NotContains $_) {
				$collection += $_
			}
		}

		return $collection
	}

	[string]getUser() {
		# TODO: Sets name of user. F.ex (anne@themaa.no)
		Write-Host -ForegroundColor White -BackgroundColor DarkBlue "Hva er eposten til brukeren?"
		$userInput = Read-Host
		return $userInput
	}

	[string]setDomain() {
		# TODO: Ask user for input on domain (ex. inotech.no). 
		# This turns into checking for "@inotech.no"
		return "todo"
	}

	[string]toJson([System.Object[]] $collection) {
		# Turns the collection into usable JSON.
		# format { "anne@themaa.no" : [ "email@themaa.no", "email@themaa.no" ]}
		return "todo"
	}
}

Function New-GetOutlookAccounts {
	[GetOutlookAccounts]::new();
}

Export-ModuleMember -Function New-GetOutlookAccounts