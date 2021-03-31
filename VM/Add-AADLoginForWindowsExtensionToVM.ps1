#  Disclaimer:    This code is not supported under any Microsoft standard support program or service.
#                 This code and information are provided "AS IS" without warranty of any kind, either
#                 expressed or implied. The entire risk arising out of the use or performance of the
#                 script and documentation remains with you. Furthermore, Microsoft or the author
#                 shall not be liable for any damages you may sustain by using this information,
#                 whether direct, indirect, special, incidental or consequential, including, without
#                 limitation, damages for loss of business profits, business interruption, loss of business
#                 information or other pecuniary loss even if it has been advised of the possibility of
#                 such damages. Read all the implementation and usage notes thoroughly.

Connect-AzAccount

$vmName = "test1"
$vmRgName = "tests"
$extensionName = "AADLoginForWindows"
$publisher = "Microsoft.Azure.ActiveDirectory"

$vm = Get-AzVm -ResourceGroupName $vmRgName -Name $vmName
Set-AzVMExtension -ResourceGroupName $vmRgName `
                    -VMName $vm.Name `
                    -Name $extensionName `
                    -Location $vm.Location `
                    -Publisher $publisher `
                    -Type "AADLoginForWindows" `
                    -TypeHandlerVersion "0.4"