using Microsoft.SystemCenter.VirtualMachineManager.UIAddIns;
using Microsoft.SystemCenter.VirtualMachineManager.UIAddIns.ContextTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.AddIn;

namespace SCVMMAddin
{
    class SCVMMAddin
    {
    }
}

[AddIn("Get VM Name")]
public class SCVMMAddinCopyVMName : ActionAddInBase
{
    public override void PerformAction(IList<ContextObject> contextObjects)
    {
        var contextObject = contextObjects[0];

        string CopyVmNameScript = @"
                $vm = Get-SCVirtualMachine -ID VMID
                $vmname = $vm.name
                $vmname | clip.exe
            ";
        string CopyVmNameScriptFormatted =
                CopyVmNameScript.Replace("VMID", contextObject.ID.ToString());
        this.PowerShellContext.ExecuteScript(CopyVmNameScriptFormatted);

    }
}

[AddIn("Get VM IP")]
public class SCVMMAddinCopyVMIP : ActionAddInBase
{
    public override void PerformAction(IList<ContextObject> contextObjects)
    {
        var contextObject = contextObjects[0];
        string CopyVmIPScript = @"
                $vm = Get-SCVirtualMachine -ID VMID
                $nic = $vm | Get-SCVirtualNetworkAdapter
                if ($nic.IPv4Addresses)
                {
                    ##$iplast=$nic[0].IPv4Addresses[$nic[0].IPv4Addresses.Count-1]
                    #$iplast=$nic.IPv4Addresses[$nic.IPv4Addresses.Count-1]

                    $lastIndex=$nic.IPv4Addresses.Count
                    $iplast=$nic.IPv4Addresses[$lastIndex-1]

                    $iplast -replace '`n |`r'
                    $iplast | clip.exe
                }
            ";
        string CopyVmIPScriptFormatted =
                CopyVmIPScript.Replace("VMID", contextObject.ID.ToString());
        this.PowerShellContext.ExecuteScript(CopyVmIPScriptFormatted);
    }
}

[AddIn("Get VM Data")]
public class SCVMMAddinGetVMDetails : ActionAddInBase
{

    public override void PerformAction(IList<ContextObject> contextObjects)
    {

        string CustomPropScript = @"
                if (!(Get-SCCustomProperty -Name VMPath))
                    {New-SCCustomProperty -Name VMPath -Description 'Virtual Machine Configuration Path' -AddMember @('VM')}
                if (!(Get-SCCustomProperty -Name 'Mounted ISO'))
                    {New-SCCustomProperty -Name 'Mounted ISO' -Description 'Virtual Machine Mounted ISO Path Path' -AddMember @('VM')}
                if (!(Get-SCCustomProperty -Name 'IP Address'))
                    {New-SCCustomProperty -Name 'IP Address' -Description 'IP addresses statically assigned' -AddMember @('VM')}
                if (!(Get-SCCustomProperty -Name 'Checkpoints'))
                    {New-SCCustomProperty -Name 'Checkpoints' -Description 'Number of checkpoints on the VM' -AddMember @('VM')}
                if (!(Get-SCCustomProperty -Name 'VLAN'))
                    {New-SCCustomProperty -Name 'VLAN' -Description 'VLAN for NIC' -AddMember @('VM')}
                ";

        this.PowerShellContext.ExecuteScript(CustomPropScript);

        foreach (var contextObject in contextObjects)
        {
            Guid jobguid = System.Guid.NewGuid();

            string UpdateScript = @"
                    if (!(Get-SCCustomProperty -Name VMPath))
                        {New-SCCustomProperty -Name VMPath -Description 'Virtual Machine Configuration Path' -AddMember @('VM')}
                    if (!(Get-SCCustomProperty -Name 'Mounted ISO'))
                        {New-SCCustomProperty -Name 'Mounted ISO' -Description 'Virtual Machine Mounted ISO Path Path' -AddMember @('VM')}
                    if (!(Get-SCCustomProperty -Name 'IP Address'))
                        {New-SCCustomProperty -Name 'IP Address' -Description 'First IP address of first adapter' -AddMember @('VM')}
                    if (!(Get-SCCustomProperty -Name 'VLAN'))
                        {New-SCCustomProperty -Name 'VLAN' -Description 'VLAN for NIC' -AddMember @('VM')}

                    $vm = Get-SCVirtualMachine -ID VMID
                    $jobguid = [system.guid]::newguid()
                    $prop = Get-SCCustomProperty -Name 'VMPath'
                    $location = $vm.Location
                    $customProp = $vm | Get-SCCustomPropertyValue -CustomProperty $prop
                    if ($customProp.Value -ne $location)
                    {
                        $vm | Set-SCCustomPropertyValue -CustomProperty $prop -Value $location -RunAsynchronously | out-null
                    }
                    
                    $prop = Get-SCCustomProperty -Name 'Mounted ISO'
                    $dvd = $vm | Get-SCVirtualDVDDrive
                    $customProp = $vm | Get-SCCustomPropertyValue -CustomProperty $prop
                    if ($dvd.ISO -ne $null)
                    {
                        $last=$dvd.ISO.Count-1
                        #if ($customProp.Value -ne ($dvd.ISO.Name))
                        if ($customProp.Value -ne ($dvd.ISO[$last].Name))
                        {
                            #$vm | Set-SCCustomPropertyValue -CustomProperty $prop -Value ($dvd.ISO.Name) -RunAsynchronously  | out-null
                            $vm | Set-SCCustomPropertyValue -CustomProperty $prop -Value ($dvd.ISO[$last].Name) -RunAsynchronously  | out-null
                        }
                    }
                    Else
                    {
                        if ($customProp)
                        {
                            Remove-SCCustomPropertyValue -CustomPropertyValue $customProp | out-null
                        }
                    }


                    $prop = Get-SCCustomProperty -Name 'IP Address'
                    $nic = $vm | Get-SCVirtualNetworkAdapter
                    $customProp = $vm | Get-SCCustomPropertyValue -CustomProperty $prop
                    if ($nic -ne $null)
                    {
                        #$lastIndex=$nic[0].IPv4Addresses.Count
                        #$iplast=$nic[0].IPv4Addresses[$lastIndex-1]

                        $lastIndex=$nic.IPv4Addresses.Count
                        $iplast=$nic.IPv4Addresses[$lastIndex-1]
                        if ($customProp.Value -ne $iplast)
                        {
                            $vm | Set-SCCustomPropertyValue -CustomProperty $prop -Value ($iplast) -RunAsynchronously  | out-null
                        }
                    }
                    Else
                    {
                        if ($customProp)
                        {
                            Remove-SCCustomPropertyValue -CustomPropertyValue $customProp | out-null
                        }
                    }

                    $prop = Get-SCCustomProperty -Name 'VLAN'
                    $customProp = $vm | Get-SCCustomPropertyValue -CustomProperty $prop
                    if ($nic -ne $null)
                    {
                        $lastIndex=$nic.VLanID.Count
                        $vlanlast=$nic.VLanID[$lastIndex-1]
                        #if ($customProp.Value -ne $nic[0].VLanID)
                        if ($customProp.Value -ne $vlanlast)
                        {
                            #$vm | Set-SCCustomPropertyValue -CustomProperty $prop -Value ($nic[0].VLanID) -RunAsynchronously  | out-null
                            $vm | Set-SCCustomPropertyValue -CustomProperty $prop -Value ($vlanlast) -RunAsynchronously  | out-null
                        }
                        
                    }
                    Else
                    {
                        if ($customProp)
                        {
                            Remove-SCCustomPropertyValue -CustomPropertyValue $customProp | out-null
                        }
                    }

                    $prop = Get-SCCustomProperty -Name 'Checkpoints'
                    $chk = $vm | Get-ScVmCheckpoint
                    if ($chk -ne $null)
                    {
                        $chkCount = $chk.count
                    }
                    Else
                    {
                        $chkCount = 0
                    }
                    $customProp = $vm | Get-SCCustomPropertyValue -CustomProperty $prop
                    if ($customProp.Value -ne $chkCount)
                    {
                        $vm | Set-SCCustomPropertyValue -CustomProperty $prop -Value $chkCount -RunAsynchronously  | out-null
                    }
                ";

            string UpdateScriptFormatted =
                UpdateScript.Replace("VMID", contextObject.ID.ToString());
            this.PowerShellContext.ExecuteScript(UpdateScriptFormatted);
        }
    }
}

