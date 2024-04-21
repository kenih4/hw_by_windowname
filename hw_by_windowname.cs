/*
最小化または隠れたWindowを最前面に表示させる

csc /target:winexe /optimize　hw_by_windowname.cs
/target:winexeでッコンソール非表示にする

csc /lib:C:\Windows\Microsoft.NET\Framework\v4.0.30319 /reference:Microsoft.VisualBasic.dll hw_by_windowname.cs
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc /lib:C:\Windows\Microsoft.NET\Framework\v4.0.30319 /reference:Microsoft.VisualBasic.dll hw_by_windowname_TEST.cs
第1引数：相手のWindowsハンドル名またはタスク名、第2引数以降：キーストローク
hw_by_windowname.exe "2018_02.xlsm - Excel"
Reference: 
http://chokuto.ifdef.jp/urawaza/message/WM_SYSCOMMAND.html
http://blog.goo.ne.jp/masaki_goo_2006/e/424334238a1753984ce3697f1b400c99?fm=entry_awc
http://tomosoft.jp/design/?p=6624
https://dobon.net/vb/dotnet/process/appactivate.html	
*/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Media;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Collections;
using System.Diagnostics;
using System.Threading;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using System.Web.Script.Serialization;
using System.Configuration;
using System.Collections.Specialized;
//using System.Windows.Forms.DataVisualization.Charting;
//using WindowsApplication1;
//using System.Threading;
using Microsoft.VisualBasic;
using System.Runtime.InteropServices; // for DLL Import





class Program {




[DllImport("user32.dll")]
static extern bool IsZoomed(IntPtr hWnd);	// 最大化状態

//[DllImport("user32.dll")]
//static extern bool IsIconic(IntPtr hWnd);	// 最小化状態

[DllImport("user32.dll")]
static extern bool IsWindow(IntPtr hWnd);	// ウインドウ有無	

[DllImport("user32.dll")]			
static extern bool IsChild(IntPtr hWnd);	// 子ウインドウ

[DllImport("user32.dll")]
static extern bool IsWindowVisible(IntPtr hWnd);	// 可視状態

[DllImport("user32.dll")]
static extern bool IsWindowEnabled(IntPtr hWnd);	// 有効状態

[DllImport("user32.dll")]
static extern bool IsWindowUnicode(IntPtr hWnd);	// Unicodeタイプ




        [DllImport("User32.Dll")]
        static extern IntPtr GetDesktopWindow();

	[DllImport("user32.dll")]
    public static extern IntPtr FindWindowEx(IntPtr hWnd, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
    
	[DllImport("User32.dll", EntryPoint = "FindWindow")]
	public static extern Int32 FindWindow(String lpClassName, String lpWindowName);
	 
	[DllImport("User32.dll", EntryPoint = "SendMessage")]
	public static extern Int32 SendMessage(Int32 hWnd, Int32 Msg, Int32 wParam, ref COPYDATASTRUCT lParam);
	 
	[DllImport("User32.dll", EntryPoint = "SendMessage")]
	public static extern Int32 SendMessage(Int32 hWnd, Int32 Msg, Int32 wParam, Int32 lParam);
	
	// For keybd_event
	[DllImport("user32.dll")]
    public static extern uint keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

	// ShowWindow
        [DllImport("User32")]
        private static extern int ShowWindow(int hwnd, int nCmdShow);
        
	public const Int32 WM_COPYDATA = 0x4A;
	public const Int32 WM_USER = 0x400;
	 
	//COPYDATASTRUCT構造体 
	public struct COPYDATASTRUCT
	{
		public Int32 dwData;      //送信する32ビット値
		public Int32 cbData;　　　//lpDataのバイト数
		public string lpData;　　 //送信するデータへのポインタ(0も可能)
	}
	

	
  //-------------------------------------------------------------------------------------
  static void Main() {





//------  2重起動防止 -------------------------------------------------
    // Mutex の新しいインスタンスを生成する (Mutex の名前にアセンブリ名を付ける)
    System.Threading.Mutex hMutex = new System.Threading.Mutex(false, Application.ProductName);

    // Mutex のシグナルを受信できるかどうか判断する
    if (hMutex.WaitOne(0, false)) {
//        Application.Run(new Form1());


//--------
    Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
    StreamWriter writer = new StreamWriter(@"log.txt", true, sjisEnc);

	//コマンドライン引数を取得
	Console.WriteLine(System.Environment.CommandLine);
	string[] cmds = System.Environment.GetCommandLineArgs();
	if(cmds.Length==1){
		MessageBox.Show("引数を指定して下さい。第1引数：相手のWindowsハンドル名またはタスク名、第2引数：キーストローク");
		return;
	}	
	foreach (string cmd in cmds)
	{
		Console.WriteLine("Arg:	" + cmd);
	}  
  
	string WindowName=cmds[1];
		
	/*  最小化されたWindowを最前面に
	Int32 hWnd = FindWindow(null, WindowName);
    if (hWnd == 0)
    {
        MessageBox.Show("Can't get window handle");
		Console.WriteLine("相手Windowのハンドルが取得できません");
    }else{
    
//		ShowWindow( hWnd, SW_MAXIMIZE);
		ShowWindow( hWnd, 3);
// 		元に戻す
//		SendMessage( hWnd, WM_SYSCOMMAND, SC_RESTORE, 0 );
		SendMessage( hWnd, 0x0112, 0xF120 , 0 );
// 		最大化
//		SendMessage( hWnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0 );
		SendMessage( hWnd, 0x0112, 0xF030 , 0 );	
	}
	 

	TEST
	IntPtr hWnd = FindWindowEx(IntPtr.Zero, IntPtr.Zero, null, WindowName);	
	IntPtr parent_hwnd = GetDesktopWindow();


if ( IsWindow(hWnd) ){
	MessageBox.Show("Windowは存在");
if ( IsZoomed(hWnd) ){
    // 最大化の状態
	MessageBox.Show("最大化の状態");
}
else if ( IsIconic(hWnd) ){
    // 最小化の状態
MessageBox.Show("最小化の状態");
}
else{
    // 普通の状態
	MessageBox.Show("普通の状態");
}

if ( IsWindowVisible(hWnd) ){
	MessageBox.Show("Windowは可視状態");
}else{
	MessageBox.Show("Windowは可視状態ではありません");
}



if ( IsWindowEnabled(hWnd) ){
	MessageBox.Show("Windowは有効状態");
}else{
	MessageBox.Show("Windowは有効状態ではありません");
}



if ( IsZoomed(hWnd) ){
// 		元に戻す
//		SendMessage( hWnd, WM_SYSCOMMAND, SC_RESTORE, 0 );
		SendMessage( (int)hWnd, 0x0112, 0xF120 , 0 );
}
else{
// 		最大化
//		SendMessage( hWnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0 );
		SendMessage( (int)hWnd, 0x0112, 0xF030 , 0 );

}


}else{
	MessageBox.Show("Windowは存在しません");
}
*/



	IntPtr hWnd = FindWindowEx(IntPtr.Zero, IntPtr.Zero, null, WindowName);
	ActiveWindow(hWnd);

if ( IsZoomed(hWnd) ){
//	MessageBox.Show("最大化の状態");
}
else{
////	MessageBox.Show("その他の状態");
// 		最大化
//		SendMessage( hWnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0 );
		SendMessage( (int)hWnd, 0x0112, 0xF030 , 0 );
}







//---------------------------------------------------------------------


			Console.WriteLine("次");

			if(cmds.Length==2){
				Console.WriteLine("第2引数：キーストロークが指定されていません");
				return;
			}

/*			
			try
			{
                Interaction.AppActivate(WindowName);
			}
			catch (Exception)
			{
				writer.WriteLine("Exception");
				Console.WriteLine("Exception");
			}
*/			
			
			for(int i=2; i<cmds.Length; ++i)
			{
				SendKeys.SendWait(cmds[i]);
				System.Threading.Thread.Sleep(500);
			}

			/*  背後に隠れたWindowを最前面に（WindowNameで指定）
			try
			{
                Interaction.AppActivate(WindowName);
			}
			catch (Exception)
			{
				writer.WriteLine("Exception");
				Console.WriteLine("Exception");
			}
			*/
				
			/*  背後に隠れたWindowを最前面に(プロセスIDで指定) 
			Process[] ps = Process.GetProcessesByName("Evernote");
			if (ps.Length > 0)
			{
				writer.WriteLine("PID:  " + ps[0].Id);
				try
				{
					Interaction.AppActivate(ps[0].Id);
				}
				catch (Exception)
				{
					writer.WriteLine("Exception");
				}
			}else{
				writer.WriteLine("No process");
			}
			*/

		writer.Close();
//--------
    }

    // GC.KeepAlive メソッドが呼び出されるまで、ガベージ コレクション対象から除外される
    GC.KeepAlive(hMutex);

    // Mutex を閉じる (正しくは オブジェクトの破棄を保証する を参照)
    hMutex.Close();
    
//------------------------------------------------------------------

  }
  
  
  






















//------------------------------------------------------------------


//using System.Runtime.InteropServices;

/// <summary>
/// できるだけ確実に指定したウィンドウをフォアグラウンドにする
/// </summary>
/// <param name="hWnd">ウィンドウハンドル</param>
public static void ActiveWindow(IntPtr hWnd)
{
    if (hWnd == IntPtr.Zero)
    {
        return;
    }

    //ウィンドウが最小化されている場合は元に戻す
    if (IsIconic(hWnd))
    {
        ShowWindowAsync(hWnd, SW_RESTORE);
    }

    //AttachThreadInputの準備
    //フォアグラウンドウィンドウのハンドルを取得
    IntPtr forehWnd=GetForegroundWindow();
    if (forehWnd == hWnd)
    {
        return;
    }
    //フォアグラウンドのスレッドIDを取得
    uint foreThread = GetWindowThreadProcessId(forehWnd, IntPtr.Zero);
    //自分のスレッドIDを収得
    uint thisThread = GetCurrentThreadId();

    uint timeout = 200000;
    if (foreThread != thisThread)
    {
        //ForegroundLockTimeoutの現在の設定を取得
        //Visual Studio 2010, 2012起動後は、レジストリと違う値を返す
        SystemParametersInfoGet(SPI_GETFOREGROUNDLOCKTIMEOUT, 0, ref timeout, 0);
        //レジストリから取得する場合
        //timeout = (uint)Microsoft.Win32.Registry.GetValue(
        //    @"HKEY_CURRENT_USER\Control Panel\Desktop",
        //    "ForegroundLockTimeout", 200000);

        //ForegroundLockTimeoutの値を0にする
        //(SPIF_UPDATEINIFILE | SPIF_SENDCHANGE)を使いたいが、
        //  timeoutがレジストリと違う値だと戻せなくなるので使わない
        SystemParametersInfoSet(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, 0, 0);

        //入力処理機構にアタッチする
        AttachThreadInput(thisThread, foreThread, true);
    }

    //ウィンドウをフォアグラウンドにする処理
    SetForegroundWindow(hWnd);
    SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_ASYNCWINDOWPOS);
    BringWindowToTop(hWnd);
    ShowWindowAsync(hWnd, SW_SHOW);
    SetFocus(hWnd);

    if (foreThread != thisThread)
    {
        //ForegroundLockTimeoutの値を元に戻す
        //ここでも(SPIF_UPDATEINIFILE | SPIF_SENDCHANGE)は使わない
        SystemParametersInfoSet(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, timeout, 0);

        //デタッチ
        AttachThreadInput(thisThread, foreThread, false);
    }
}

[DllImport("user32.dll")]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool SetForegroundWindow(IntPtr hWnd);

[DllImport("user32.dll")]
private static extern IntPtr GetForegroundWindow();

[DllImport("user32.dll", SetLastError = true)]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool BringWindowToTop(IntPtr hWnd);

[DllImport("user32.dll")]
static extern IntPtr SetFocus(IntPtr hWnd);

[DllImport("user32.dll", SetLastError = true)]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool SetWindowPos(IntPtr hWnd,
    int hWndInsertAfter, int x, int y, int cx, int cy, int uFlags);

private const int SWP_NOSIZE = 0x0001;
private const int SWP_NOMOVE = 0x0002;
private const int SWP_NOZORDER = 0x0004;
private const int SWP_SHOWWINDOW = 0x0040;
private const int SWP_ASYNCWINDOWPOS = 0x4000;
private const int HWND_TOP = 0;
private const int HWND_BOTTOM = 1;
private const int HWND_TOPMOST = -1;
private const int HWND_NOTOPMOST = -2;

[DllImport("user32.dll")]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

private const int SW_SHOWNORMAL = 1;
private const int SW_SHOW = 5;
private const int SW_RESTORE = 9;

[DllImport("user32.dll")]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool IsIconic(IntPtr hWnd);

[DllImport("user32.dll")]
private static extern uint GetWindowThreadProcessId(
    IntPtr hWnd, IntPtr ProcessId);

[DllImport("kernel32.dll")]
private static extern uint GetCurrentThreadId();

[DllImport("user32.dll")]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool AttachThreadInput(
    uint idAttach, uint idAttachTo, bool fAttach);

[DllImport("user32.dll", EntryPoint = "SystemParametersInfo",
    SetLastError = true)]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool SystemParametersInfoGet(
    uint action, uint param, ref uint vparam, uint init);

[DllImport("user32.dll", EntryPoint = "SystemParametersInfo",
    SetLastError = true)]
[return: MarshalAs(UnmanagedType.Bool)]
private static extern bool SystemParametersInfoSet(
    uint action, uint param, uint vparam, uint init);

private const uint SPI_GETFOREGROUNDLOCKTIMEOUT = 0x2000;
private const uint SPI_SETFOREGROUNDLOCKTIMEOUT = 0x2001;
private const uint SPIF_UPDATEINIFILE = 0x01;
private const uint SPIF_SENDCHANGE = 0x02;
//-----------------------------------------------------------------






























  
  
  
  
  
  
  
  
  
}