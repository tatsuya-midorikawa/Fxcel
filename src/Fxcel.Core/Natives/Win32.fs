namespace Fxcel.Core.Natives

open System.Runtime.InteropServices
open System.Runtime.InteropServices.ComTypes

module Win32 =
  [<DllImport("ole32.dll", EntryPoint = "GetRunningObjectTable")>]
  extern int get_running_object_table(int reserved, [<Out>] IRunningObjectTable& prot)
  [<DllImport("user32.dll", EntryPoint = "SendMessage")>]
  extern int send_message(int hwnd, int msg, nativeint wp, nativeint lp)
  [<DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")>]
  extern int get_window_thread_process_id(int hwnd, [<Out>] int& lpdwProcessId)
  
  [<DllImport("user32.dll", EntryPoint = "GetParent")>]
  extern int get_parent(int hwnd);
  [<DllImport("user32.dll", EntryPoint = "GetWindow")>]
  extern int get_window(int hwnd, int cmd);
  [<DllImport("user32.dll", EntryPoint = "FindWindow")>]
  extern int find_window(string className, string windowName);
  [<DllImport("user32.dll", EntryPoint = "IsWindowVisible")>]
  extern int is_window_visible(int hwnd);
