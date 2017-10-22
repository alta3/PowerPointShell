function pps-add {
    [CmdletBinding()]
    param(
        [string]$message
    )

    write-host "You ran the add command!"
    write-host "The following message has been added: "
    write-host $message
}