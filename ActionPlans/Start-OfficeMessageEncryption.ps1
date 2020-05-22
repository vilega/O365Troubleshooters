Function Start-AP_OfficeMessageEncryption {
        
    # Connect Workloads (split workloads by comma): "msol","exo","eop","sco","spo","sfb","aadrm"
    $Workloads = "exo", "sco", "aadrm"
    Connect-O365PS $Workloads
        
    $CurrentProperty = "Connecting to: $Workloads"
    $CurrentDescription = "Success"
    write-log -Function "Connecting to O365 workloads" -Step $CurrentProperty -Description $CurrentDescription 
        
    # Main Function
        
    # Disconnecting
    disconnect-all  
}