/*
�ŏ����܂��͉B�ꂽWindow���őO�ʂɕ\��������

csc /target:winexe /optimize�@hw_by_windowname.cs
/target:winexe�Ńb�R���\�[����\���ɂ���

csc /lib:C:\Windows\Microsoft.NET\Framework\v4.0.30319 /reference:Microsoft.VisualBasic.dll hw_by_windowname.cs
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc /lib:C:\Windows\Microsoft.NET\Framework\v4.0.30319 /reference:Microsoft.VisualBasic.dll hw_by_windowname_TEST.cs
��1�����F�����Windows�n���h�����܂��̓^�X�N���A��2�����ȍ~�F�L�[�X�g���[�N
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
static extern bool IsZoomed(IntPtr hWnd);	// �ő剻���

//[DllImport("user32.dll")]
//static extern bool IsIconic(IntPtr hWnd);	// �ŏ������

[DllImport("user32.dll")]
static extern bool IsWindow(IntPtr hWnd);	// �E�C���h�E�L��	

[DllImport("user32.dll")]			
static extern bool IsChild(IntPtr hWnd);	// �q�E�C���h�E

[DllImport("user32.dll")]
static extern bool IsWindowVisible(IntPtr hWnd);	// �����

[DllImport("user32.dll")]
static extern bool IsWindowEnabled(IntPtr hWnd);	// �L�����

[DllImport("user32.dll")]
static extern bool IsWindowUnicode(IntPtr hWnd);	// Unicode�^�C�v




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
	 
	//COPYDATASTRUCT�\���� 
	public struct COPYDATASTRUCT
	{
		public Int32 dwData;      //���M����32�r�b�g�l
		public Int32 cbData;�@�@�@//lpData�̃o�C�g��
		public string lpData;�@�@ //���M����f�[�^�ւ̃|�C���^(0���\)
	}
	

	
  //-------------------------------------------------------------------------------------
  static void Main() {





//------  2�d�N���h�~ -------------------------------------------------
    // Mutex �̐V�����C���X�^���X�𐶐����� (Mutex �̖��O�ɃA�Z���u������t����)
    System.Threading.Mutex hMutex = new System.Threading.Mutex(false, Application.ProductName);

    // Mutex �̃V�O�i������M�ł��邩�ǂ������f����
    if (hMutex.WaitOne(0, false)) {
//        Application.Run(new Form1());


//--------
    Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
    StreamWriter writer = new StreamWriter(@"log.txt", true, sjisEnc);

	//�R�}���h���C���������擾
	Console.WriteLine(System.Environment.CommandLine);
	string[] cmds = System.Environment.GetCommandLineArgs();
	if(cmds.Length==1){
		MessageBox.Show("�������w�肵�ĉ������B��1�����F�����Windows�n���h�����܂��̓^�X�N���A��2�����F�L�[�X�g���[�N");
		return;
	}	
	foreach (string cmd in cmds)
	{
		Console.WriteLine("Arg:	" + cmd);
	}  
  
	string WindowName=cmds[1];
		
	/*  �ŏ������ꂽWindow���őO�ʂ�
	Int32 hWnd = FindWindow(null, WindowName);
    if (hWnd == 0)
    {
        MessageBox.Show("Can't get window handle");
		Console.WriteLine("����Window�̃n���h�����擾�ł��܂���");
    }else{
    
//		ShowWindow( hWnd, SW_MAXIMIZE);
		ShowWindow( hWnd, 3);
// 		���ɖ߂�
//		SendMessage( hWnd, WM_SYSCOMMAND, SC_RESTORE, 0 );
		SendMessage( hWnd, 0x0112, 0xF120 , 0 );
// 		�ő剻
//		SendMessage( hWnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0 );
		SendMessage( hWnd, 0x0112, 0xF030 , 0 );	
	}
	 

	TEST
	IntPtr hWnd = FindWindowEx(IntPtr.Zero, IntPtr.Zero, null, WindowName);	
	IntPtr parent_hwnd = GetDesktopWindow();


if ( IsWindow(hWnd) ){
	MessageBox.Show("Window�͑���");
if ( IsZoomed(hWnd) ){
    // �ő剻�̏��
	MessageBox.Show("�ő剻�̏��");
}
else if ( IsIconic(hWnd) ){
    // �ŏ����̏��
MessageBox.Show("�ŏ����̏��");
}
else{
    // ���ʂ̏��
	MessageBox.Show("���ʂ̏��");
}

if ( IsWindowVisible(hWnd) ){
	MessageBox.Show("Window�͉����");
}else{
	MessageBox.Show("Window�͉���Ԃł͂���܂���");
}



if ( IsWindowEnabled(hWnd) ){
	MessageBox.Show("Window�͗L�����");
}else{
	MessageBox.Show("Window�͗L����Ԃł͂���܂���");
}



if ( IsZoomed(hWnd) ){
// 		���ɖ߂�
//		SendMessage( hWnd, WM_SYSCOMMAND, SC_RESTORE, 0 );
		SendMessage( (int)hWnd, 0x0112, 0xF120 , 0 );
}
else{
// 		�ő剻
//		SendMessage( hWnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0 );
		SendMessage( (int)hWnd, 0x0112, 0xF030 , 0 );

}


}else{
	MessageBox.Show("Window�͑��݂��܂���");
}
*/



	IntPtr hWnd = FindWindowEx(IntPtr.Zero, IntPtr.Zero, null, WindowName);
	ActiveWindow(hWnd);

if ( IsZoomed(hWnd) ){
//	MessageBox.Show("�ő剻�̏��");
}
else{
////	MessageBox.Show("���̑��̏��");
// 		�ő剻
//		SendMessage( hWnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0 );
		SendMessage( (int)hWnd, 0x0112, 0xF030 , 0 );
}







//---------------------------------------------------------------------


			Console.WriteLine("��");

			if(cmds.Length==2){
				Console.WriteLine("��2�����F�L�[�X�g���[�N���w�肳��Ă��܂���");
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

			/*  �w��ɉB�ꂽWindow���őO�ʂɁiWindowName�Ŏw��j
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
				
			/*  �w��ɉB�ꂽWindow���őO�ʂ�(�v���Z�XID�Ŏw��) 
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

    // GC.KeepAlive ���\�b�h���Ăяo�����܂ŁA�K�x�[�W �R���N�V�����Ώۂ��珜�O�����
    GC.KeepAlive(hMutex);

    // Mutex ����� (�������� �I�u�W�F�N�g�̔j����ۏ؂��� ���Q��)
    hMutex.Close();
    
//------------------------------------------------------------------

  }
  
  
  






















//------------------------------------------------------------------


//using System.Runtime.InteropServices;

/// <summary>
/// �ł��邾���m���Ɏw�肵���E�B���h�E���t�H�A�O���E���h�ɂ���
/// </summary>
/// <param name="hWnd">�E�B���h�E�n���h��</param>
public static void ActiveWindow(IntPtr hWnd)
{
    if (hWnd == IntPtr.Zero)
    {
        return;
    }

    //�E�B���h�E���ŏ�������Ă���ꍇ�͌��ɖ߂�
    if (IsIconic(hWnd))
    {
        ShowWindowAsync(hWnd, SW_RESTORE);
    }

    //AttachThreadInput�̏���
    //�t�H�A�O���E���h�E�B���h�E�̃n���h�����擾
    IntPtr forehWnd=GetForegroundWindow();
    if (forehWnd == hWnd)
    {
        return;
    }
    //�t�H�A�O���E���h�̃X���b�hID���擾
    uint foreThread = GetWindowThreadProcessId(forehWnd, IntPtr.Zero);
    //�����̃X���b�hID������
    uint thisThread = GetCurrentThreadId();

    uint timeout = 200000;
    if (foreThread != thisThread)
    {
        //ForegroundLockTimeout�̌��݂̐ݒ���擾
        //Visual Studio 2010, 2012�N����́A���W�X�g���ƈႤ�l��Ԃ�
        SystemParametersInfoGet(SPI_GETFOREGROUNDLOCKTIMEOUT, 0, ref timeout, 0);
        //���W�X�g������擾����ꍇ
        //timeout = (uint)Microsoft.Win32.Registry.GetValue(
        //    @"HKEY_CURRENT_USER\Control Panel\Desktop",
        //    "ForegroundLockTimeout", 200000);

        //ForegroundLockTimeout�̒l��0�ɂ���
        //(SPIF_UPDATEINIFILE | SPIF_SENDCHANGE)���g���������A
        //  timeout�����W�X�g���ƈႤ�l���Ɩ߂��Ȃ��Ȃ�̂Ŏg��Ȃ�
        SystemParametersInfoSet(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, 0, 0);

        //���͏����@�\�ɃA�^�b�`����
        AttachThreadInput(thisThread, foreThread, true);
    }

    //�E�B���h�E���t�H�A�O���E���h�ɂ��鏈��
    SetForegroundWindow(hWnd);
    SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_ASYNCWINDOWPOS);
    BringWindowToTop(hWnd);
    ShowWindowAsync(hWnd, SW_SHOW);
    SetFocus(hWnd);

    if (foreThread != thisThread)
    {
        //ForegroundLockTimeout�̒l�����ɖ߂�
        //�����ł�(SPIF_UPDATEINIFILE | SPIF_SENDCHANGE)�͎g��Ȃ�
        SystemParametersInfoSet(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, timeout, 0);

        //�f�^�b�`
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