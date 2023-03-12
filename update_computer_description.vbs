'------------------------------------------------------------------------
' Citra IT - Excelência em TI
' Script para atualizar a descrição dos computadores no ActiveDirectory
' @Author: luciano@citrait.com.br
' @Date: 2023/03/12 @Version: 1.0
' @Usage: Agende como um script de logon dos usuários nos computadores.
' @Obs.: É necessário criar delegação para objetos do tipo computador _
'        Para que os usuarios consigam atualizar o atributo descrição _
'        dos computadores.
'------------------------------------------------------------------------
On Error Resume Next
'Option Explicit


' Chaves de registros de onde obter informações sobre o domínio e sobre o DN do computador no AD
' E seguro assumir que pode ler este valor a partir da GPO padrao 'default domain policy'
strRegistryDN     = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\DataStore\Machine\0\DNName"
strRegistryDomain = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\DataStore\Machine\0\DomainName"


' Instanciando os objetos COM network e shell
Set objNetwork = CreateObject("WScript.Network")
Set objShell   = CreateObject("WScript.Shell")

' Lendo a chave do registro com o dn do computador e do dominio
strComputerDomain = objShell.RegRead( strRegistryDomain )
strComputerDN     = objShell.RegRead( strRegistryDN )

' Conecta no DomainController disponível
Set objComputer = GetObject("LDAP://" & strComputerDomain & "/" & strComputerDN)

' Atualiza a descrição do computador no formato: Username - Computername - Datetime
objComputer.Put "Description", objNetwork.Computername & " - " & objNetwork.Username & " - " & Now()
objComputer.SetInfo()

