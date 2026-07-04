@{
    # Severity levels to include
    Severity = @('Error', 'Warning')

    # Rules to run (runs all rules by default unless excluded)
    IncludeRules = @()

    # Specific rules to exclude if they become too noisy
    ExcludeRules = @(
        # 'PSAvoidUsingWriteHost' # Uncomment if you want to allow Write-Host without warnings
    )

    # Configuration for specific rules
    Rules = @{
        # Enforce approved verbs for any custom functions/cmdlets
        PSUseApprovedVerbs = @{
            Enable = $true
        }
        # Enforce singular nouns for cmdlets
        PSUseSingularNouns = @{
            Enable = $true
        }
        # Warn if aliases (like dir, ls, copy) are used instead of full cmdlet names
        PSAvoidUsingCmdletAliases = @{
            Enable = $true
        }
    }
}
