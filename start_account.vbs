Set WshShell = CreateObject("WScript.Shell")
batFilePath = "C:\\Qalat_Al_Khaleej\\QAK_Account\\xampp\\setup_xampp.bat"
WshShell.Run """" & batFilePath & """", 0, False
WshShell.Run "node_install.bat", 0, False

