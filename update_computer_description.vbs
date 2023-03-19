'------------------------------------------------------------------------
' Citra IT - Excelência em TI
' Script para atualizar a descrição dos computadores no ActiveDirectory
' @Author: luciano@citrait.com.br
' @Date: 2023/03/17 @Version: 1.1 | atualização para recuperar informações do computador,usuário e domínio 
'                                 | usando objeto com e não o registro (fix compatível win11.
' @Date: 2023/03/12 @Version: 1.0 | release inicial
' @Usage: Agende como um script de logon dos usuários nos computadores.
' @Obs.: É necessário criar delegação para objetos do tipo computador _
'        Para que os usuarios consigam atualizar o atributo descrição _
'        dos computadores.
'------------------------------------------------------------------------
On Error Resume Next
'Option Explicit


' Conectando no objeto COM ADSystemInfo para recuperar informações do computador (Domínio atual e DN)
' https://learn.microsoft.com/en-us/windows/win32/adsi/iadsadsysteminfo-property-methods
Set ADSystemInfo = CreateObject("ADSystemInfo")

' Instanciando objeto Com network para recuperar usuário conectado e Hostname do computador
Set objNetwork = CreateObject("WScript.Network")

' Conecta no DomainController disponível
Set objComputer = GetObject("LDAP://" & ADSystemInfo.DomainDNSName & "/" & ADSystemInfo.ComputerName)

' Atualiza a descrição do computador no formato: Username - Computername - Datetime
objComputer.Put "Description", objNetwork.Computername & " - " & objNetwork.UserName & " - " & Now()
objComputer.SetInfo()

