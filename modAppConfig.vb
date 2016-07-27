Imports System.Configuration
Module modAppConfig


    Sub ReadAllSettings()
        Try
            Dim appSettings = ConfigurationManager.AppSettings

            If appSettings.Count = 0 Then
                Console.WriteLine("AppSettings is empty.")
            Else
                For Each key As String In appSettings.AllKeys
                    Console.WriteLine("Key: {0} Value: {1}", key, appSettings(key))
                Next
            End If
        Catch e As ConfigurationErrorsException
            Console.WriteLine("Error reading app settings")
        End Try
    End Sub

    Function ReadSetting(ByVal key As String) As String
        Try
            Dim appSettings = ConfigurationManager.AppSettings
            Dim result As String = appSettings(key)
            If IsNothing(result) Then
                Return "Not found"
            End If
            Return result
        Catch e As ConfigurationErrorsException
            Console.WriteLine("Error reading app settings")
        End Try
        Return ""
    End Function

    Sub AddUpdateAppSettings(ByVal key As String, ByVal value As String)
        Try
            Dim configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
            Dim settings = configFile.AppSettings.Settings
            If IsNothing(settings(key)) Then
                settings.Add(key, value)
            Else
                settings(key).Value = value
            End If
            configFile.Save(ConfigurationSaveMode.Modified)
            ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name)
        Catch e As ConfigurationErrorsException
            Console.WriteLine("Error writing app settings")
        End Try
    End Sub

End Module