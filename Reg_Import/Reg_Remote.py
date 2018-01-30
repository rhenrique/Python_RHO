import _winreg
import wmi

class ConnectToRemoteWindowsMachine:
      def __init__(self, hostname, username, password):
          self.hostname = hostname
          self.username = username
          self.password = password
          # self.remote_path = hostname
          try:
              print("Establishing connection to .....%s" %self.hostname)
              self.connection = wmi.WMI(self.hostname, user=self.username, password=self.password)
              print("Connection established")
          except wmi.x_wmi:
              print("Could not connect to machine")
              raise

      def run_remote(self, async=False, minimized=False):

          SW_SHOWNORMAL = 1

          process_startup = self.connection.Win32_ProcessStartup.new()
          process_startup.ShowWindow = SW_SHOWNORMAL

          process_id, result = self.connection.Win32_Process.Create(
              CommandLine="notepad.exe",
              ProcessStartupInformation=process_startup
          )
          if result == 0:
              print "Process started successfully: %d" % process_id
          else:
              raise RuntimeError, "Problem creating process: %d" % result

w = ConnectToRemoteWindowsMachine('LIM-SILVAPG','admoliveirh','312414ra@#')
print(w)
print(w.run_remote())