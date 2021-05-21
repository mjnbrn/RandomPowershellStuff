function New-Macro {
        [CmdletBinding()]
        param (
            [string]$Skill,
            [string]$Target,
            [switch]$Quiet
        )
        
        
$Macro = @"
/micon "$Skill"
/merror off
/ac "$Skill" <$Target>
/ac "$Skill" <$Target>
/ac "$Skill" <$Target>
/ac "$Skill" <$Target>
/ac "$Skill" <$Target>
/ac "$Skill" <$Target>
/ac "$Skill" <$Target>
/ac "$Skill" <$Target>
/ac "$Skill" <$Target>
/ac "$Skill" <$Target>
/ac "$Skill" <$Target>
/ac "$Skill" <$Target>
/ac "$Skill"
"@
$macro | clip
if ( !$Quiet ) { } else { return $Macro }
}
