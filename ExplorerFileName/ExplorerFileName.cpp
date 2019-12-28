// ExplorerFileName.cpp : 定义应用程序的入口点。
//
//////////////////////////////////////////////////////////////////////////
//
// 功能：实时获取打开的资源管理器中鼠标左键选中的文件路径及文件名
//
//////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ExplorerFileName.h"

//包含头文件
#include <windows.h>
#include <shlwapi.h>
#include <shlobj.h>
#include <ExDisp.h>
#pragma comment(lib, "shlwapi.lib")


// 全局变量:
#define MAX_LOADSTRING 100

HINSTANCE hInst;								// 当前实例
TCHAR szTitle[MAX_LOADSTRING];					// 标题栏文本
TCHAR szWindowClass[MAX_LOADSTRING];			// 主窗口类名

//定时器
#define MAKETIME 216
//定时器到点执行回调的函数
void CALLBACK RecalcText(HWND hwnd, UINT, UINT_PTR, DWORD);

//窗口宽度，高度
const int g_WndWidth  = 600;
const int g_WndHeight = 400;
//取得的文件路径和文件名
TCHAR g_szPath[MAX_PATH];
TCHAR g_szItem[MAX_PATH];

// 此代码模块中包含的函数的前向声明:
ATOM				MyRegisterClass(HINSTANCE hInstance);
BOOL				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK	About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY _tWinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPTSTR    lpCmdLine,
                     int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

 	// TODO: 在此放置代码。
	MSG msg;
	HACCEL hAccelTable;

	// 初始化全局字符串
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_EXPLORERFILENAME, szWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

	// 执行应用程序初始化:
	if (!InitInstance (hInstance, nCmdShow))
	{
		return FALSE;
	}
	hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_EXPLORERFILENAME));

	// 主消息循环:
	while (GetMessage(&msg, NULL, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	return (int) msg.wParam;
}



//
//  函数: MyRegisterClass()
//
//  目的: 注册窗口类。
//
//  注释:
//
//    仅当希望
//    此代码与添加到 Windows 95 中的“RegisterClassEx”
//    函数之前的 Win32 系统兼容时，才需要此函数及其用法。调用此函数十分重要，
//    这样应用程序就可以获得关联的
//    “格式正确的”小图标。
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_EXPLORERFILENAME));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= MAKEINTRESOURCE(IDC_EXPLORERFILENAME);
	wcex.lpszClassName	= szWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_EXPLORERFILENAME));

	return RegisterClassEx(&wcex);
}

//
//   函数: InitInstance(HINSTANCE, int)
//   目的: 保存实例句柄并创建主窗口
//   注释:
//        在此函数中，我们在全局变量中保存实例句柄并
//        创建和显示主程序窗口。
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   HWND hWnd;

   hInst = hInstance; // 将实例句柄存储在全局变量中

   hWnd = CreateWindow(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, g_WndWidth, g_WndHeight, NULL, NULL, hInstance, NULL);


   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   //200毫秒
   ::SetTimer(hWnd,MAKETIME,200,RecalcText);

   return TRUE;
}

//
//  函数: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  目的: 处理主窗口的消息。
//
//  WM_COMMAND	- 处理应用程序菜单
//  WM_PAINT	- 绘制主窗口
//  WM_DESTROY	- 发送退出消息并返回
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	PAINTSTRUCT ps;
	HBRUSH	hBrush ;
	RECT	rc ;

	int ilen ;

	switch (message)
	{
	case WM_COMMAND:
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam);
		// 分析菜单选择:
		switch (wmId)
		{
		case IDM_ABOUT:
			DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
			break;
		case IDM_EXIT:
			DestroyWindow(hWnd);
			break;
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}
		break;
	case WM_DESTROY:
		KillTimer(hWnd,MAKETIME);
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}


BOOL    ILIsFile(LPCITEMIDLIST pidl)
{
	BOOL bRet=FALSE;
	LPCITEMIDLIST  pidlChild=NULL;
	IShellFolder* psf=NULL;
	HRESULT    hr = SHBindToParent(pidl, IID_IShellFolder, (LPVOID*)&psf, &pidlChild);
	if (SUCCEEDED(hr) && psf)
	{
		SFGAOF rgfInOut=SFGAO_FOLDER|SFGAO_FILESYSTEM ;
		hr=psf->GetAttributesOf(1,&pidlChild,&rgfInOut);
		if (SUCCEEDED(hr))
		{
			//in file system, but is not a folder
			if( (~rgfInOut & SFGAO_FOLDER) && (rgfInOut&SFGAO_FILESYSTEM) )
			{
				bRet=TRUE;    
			}
		}
		psf->Release();
	}
	return bRet;
}

//定时器
void CALLBACK RecalcText(HWND hWnd, UINT, UINT_PTR, DWORD)
{
	//变量初始化
	BOOL bResult = FALSE;
	BOOL lb_isExplorer = FALSE;
	TCHAR szBuff[256]; 
	BOOL bFile = 0;		 //当前选中项类型 ( 1/2/3 表示：文件/文件夹/当前选择了多项 )
	g_szPath[0] = TEXT('\0');	//选中文件的路径
	g_szItem[0] = TEXT('\0');	//选中文件的文件名

	//获取最前端窗口的句柄
	HWND hExplorer = GetForegroundWindow();
	GetClassName(hExplorer,szBuff,sizeof(szBuff)/sizeof(TCHAR));

	//本机三个操作系统上的explorer.exe窗口类型都是CabinetWClass/ExploreWClass，不知道其他系统是不是，待验证
	if (_tcscmp(szBuff,_T("CabinetWClass")) == 0 || _tcscmp(szBuff,_T("ExploreWClass")) == 0)  
	{
		lb_isExplorer = TRUE;
	}
	if (!lb_isExplorer)
	{
		return;
	}

	IShellWindows *psw = NULL; 
	HRESULT hr = CoCreateInstance(CLSID_ShellWindows,NULL,CLSCTX_ALL,IID_IShellWindows,(void**)&psw);
	if (SUCCEEDED(hr)) 
	{
		//SWC_EXPLORER 指出获取explorer打开的窗口，如果没有打开任何窗口则会调用失败  
		//VISTA及以后的版本可以用SWC_DESKTOP，此时获取的是桌面窗口，即使没有其他shell 窗口，该调用也不会失败.  
		VARIANT v;
		V_VT(&v) = VT_I4;//其含义是，对于指针X，取其指向的对象的vt成员。即 vpidl.vt = VT_I4;  
		v.ulVal = SWC_EXPLORER;//V_VT(&v)如果是VT_I, 或VT_I4子类型，其值将被解释为Shell windows collection 的索引。
		BOOL fFound = FALSE; 
		long count = -1;
		IDispatch  *pdisp; 

		// 获取打开的shell窗口个数.获取数目可能会出错，判断if (pdisp == NULL)即可
		if (SUCCEEDED(psw->get_Count(&count)))  
		{
			//#define V_I4(X)   V_UNION(X, lVal) 
			//#define V_UNION(X, Y)   ((X)->Y)
			for (V_I4(&v) = 0; !fFound && (V_I4(&v) < count); V_I4(&v)++)//V_I4(&v) == v->lVal
			{
				if (SUCCEEDED(psw->Item(v, &pdisp)))
				{
					//如果获取输出错误，一定是获取多了，只有存在的explorer才会有pdisp，其他窗口为空,那就获取下一个！别去查询接口！
					if (pdisp == NULL)	
					{
						continue;
					}
					IWebBrowserApp *pwba;// 这个接口是一个控件接口
					if (SUCCEEDED(pdisp->QueryInterface(IID_IWebBrowserApp, (void**)&pwba))) 
					{
						HWND l_hwnd;
						if (SUCCEEDED(pwba->get_HWND((LONG_PTR*)&l_hwnd)) && l_hwnd == hExplorer) 
						{
							fFound = TRUE;
							IServiceProvider *psp;// 接口类工厂，用于创建另外一些接口对象
							if (SUCCEEDED(pwba->QueryInterface(IID_IServiceProvider, (void**)&psp))) 
							{
								// SID_STopLevelBrowser指出获取最顶层的那个shell窗口.  
								//可以调用该接口的GetControlWindow成员函数，然后再用sp++查看一下，就知道是哪个窗口了。
								IShellBrowser *psb; //Shell浏览器
								if (SUCCEEDED(psp->QueryService(SID_STopLevelBrowser,IID_IShellBrowser, (void**)&psb))) 
								{
									IShellView *psv; //Shell视图
									if (SUCCEEDED(psb->QueryActiveShellView(&psv))) 
									{
										IFolderView *pfv;//文件夹视图
										if (SUCCEEDED(psv->QueryInterface(IID_IFolderView,(void**)&pfv))) 
										{
											//////////////////////////////////////////////////////////////////////////										
											IPersistFolder2 *ppf2;
											if (SUCCEEDED(pfv->GetFolder(IID_IPersistFolder2,(void**)&ppf2))) {
												LPITEMIDLIST pidlFolder;
												if (SUCCEEDED(ppf2->GetCurFolder(&pidlFolder))) {
													//SHGetPathFromIDList 功能是把项目标志符列表转换为文档系统路径
													//pidlFolder--相对 namespace 的根指定一个文档或目录地点的一张项目标识符表的地址 ( 桌面 ) ；
													if (!SHGetPathFromIDList(pidlFolder, g_szPath)) {   //获取文件夹路径
														lstrcpyn(g_szPath, TEXT("<not a directory>"), MAX_PATH);
													}
													bResult = TRUE;
													int iFocus;
													int selectedNmu = 0;

													if (SUCCEEDED(pfv->GetFocusedItem(&iFocus))) 
													{
														if (SUCCEEDED(pfv->ItemCount(SVGIO_SELECTION,&selectedNmu)))
														{
															if (selectedNmu == 1)
															{
																LPITEMIDLIST pidlItem;
																if (SUCCEEDED(pfv->Item(iFocus, &pidlItem))) 
																{
																	IShellFolder *psf;
																	if (SUCCEEDED(ppf2->QueryInterface(IID_IShellFolder,(void**)&psf))) 
																	{
																		STRRET str;
																		if (SUCCEEDED(psf->GetDisplayNameOf(pidlItem,SHGDN_INFOLDER,&str))) 
																		{
																			StrRetToBuf(&str, pidlItem, g_szItem, MAX_PATH);   //获取文件名 
																			bFile = ILIsFile(pidlItem);
																		}//SVSI_SELECT
																		psf->Release();
																	}
																	CoTaskMemFree(pidlItem);  //此函数用于释放被分配的内存块。															
																} 
															}else 
																bFile = 3;
														}																								
													}
													CoTaskMemFree(pidlFolder);
												}
												ppf2->Release();
											}
											pfv->Release(); 
										}
										psv->Release();
									}
									psb->Release();
								}
								psp->Release();
							}
						}
						pwba->Release();
					}
					pdisp->Release();
				}
			}
		}
		psw->Release();
	}
	if (bResult)
	{
		HDC		hdc;
		RECT	rect ;
		HBRUSH hBrush ;

		hdc = GetDC(hWnd);  
		GetClientRect (hWnd, &rect) ;
		hBrush = CreateSolidBrush(RGB(0,0,0));
		FillRect (hdc, &rect, hBrush);
		InvertRect (hdc, &rect) ;

		TextOut(hdc,rect.left, rect.top, g_szPath, lstrlen(g_szPath));
		TextOut(hdc,rect.left, rect.top+20, g_szItem, lstrlen(g_szItem));

		DeleteObject (hBrush) ;
		ReleaseDC(hWnd,hdc);  
	}
}

// “关于”框的消息处理程序。
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}
