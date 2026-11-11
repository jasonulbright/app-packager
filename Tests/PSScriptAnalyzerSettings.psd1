@{
    Severity = @('Error', 'Warning', 'Information')
    # Pilot defaults, 2026-11-02. Each excluded rule is advice about style, or reports a false positive
    # in these projects. The owner can change the list. See wiki/concepts/powershell-linting.md.
    ExcludeRules = @(
        'PSAvoidUsingWriteHost'
        'PSReviewUnusedParameter'
        'PSUseSingularNouns'
        'PSUseShouldProcessForStateChangingFunctions'
        'PSAvoidUsingPositionalParameters'
        'PSAvoidTrailingWhitespace'
    )
}