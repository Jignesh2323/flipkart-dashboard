' Runs stock-watch.ps1 with no visible window (used by the "Flipkart Stock Watch" scheduled task).
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = "E:\Claude\Flipkart"
sh.Run "powershell -NoProfile -ExecutionPolicy Bypass -File ""E:\Claude\Flipkart\stock-watch.ps1""", 0, True
