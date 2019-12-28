// ExplorerFileName.cpp : ����Ӧ�ó������ڵ㡣
//
//////////////////////////////////////////////////////////////////////////
//
// ���ܣ�ʵʱ��ȡ�򿪵���Դ��������������ѡ�е��ļ�·�����ļ���
//
//////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ExplorerFileName.h"

//����ͷ�ļ�
#include <windows.h>
#include <shlwapi.h>
#include <shlobj.h>
#include <ExDisp.h>
#pragma comment(lib, "shlwapi.lib")


// ȫ�ֱ���:
#define MAX_LOADSTRING 100

HINSTANCE hInst;								// ��ǰʵ��
TCHAR szTitle[MAX_LOADSTRING];					// �������ı�
TCHAR szWindowClass[MAX_LOADSTRING];			// ����������

//��ʱ��
#define MAKETIME 216
//��ʱ������ִ�лص��ĺ���
void CALLBACK RecalcText(HWND hwnd, UINT, UINT_PTR, DWORD);

//���ڿ�ȣ��߶�
const int g_WndWidth  = 600;
const int g_WndHeight = 400;
//ȡ�õ��ļ�·�����ļ���
TCHAR g_szPath[MAX_PATH];
TCHAR g_szItem[MAX_PATH];

// �˴���ģ���а����ĺ�����ǰ������:
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

 	// TODO: �ڴ˷��ô��롣
	MSG msg;
	HACCEL hAccelTable;

	// ��ʼ��ȫ���ַ���
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_EXPLORERFILENAME, szWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

	// ִ��Ӧ�ó����ʼ��:
	if (!InitInstance (hInstance, nCmdShow))
	{
		return FALSE;
	}
	hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_EXPLORERFILENAME));

	// ����Ϣѭ��:
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
//  ����: MyRegisterClass()
//
//  Ŀ��: ע�ᴰ���ࡣ
//
//  ע��:
//
//    ����ϣ��
//    �˴�������ӵ� Windows 95 �еġ�RegisterClassEx��
//    ����֮ǰ�� Win32 ϵͳ����ʱ������Ҫ�˺��������÷������ô˺���ʮ����Ҫ��
//    ����Ӧ�ó���Ϳ��Ի�ù�����
//    ����ʽ��ȷ�ġ�Сͼ�ꡣ
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
//   ����: InitInstance(HINSTANCE, int)
//   Ŀ��: ����ʵ�����������������
//   ע��:
//        �ڴ˺����У�������ȫ�ֱ����б���ʵ�������
//        ��������ʾ�����򴰿ڡ�
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   HWND hWnd;

   hInst = hInstance; // ��ʵ������洢��ȫ�ֱ�����

   hWnd = CreateWindow(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, g_WndWidth, g_WndHeight, NULL, NULL, hInstance, NULL);


   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   //200����
   ::SetTimer(hWnd,MAKETIME,200,RecalcText);

   return TRUE;
}

//
//  ����: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  Ŀ��: ���������ڵ���Ϣ��
//
//  WM_COMMAND	- ����Ӧ�ó���˵�
//  WM_PAINT	- ����������
//  WM_DESTROY	- �����˳���Ϣ������
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
		// �����˵�ѡ��:
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

//��ʱ��
void CALLBACK RecalcText(HWND hWnd, UINT, UINT_PTR, DWORD)
{
	//������ʼ��
	BOOL bResult = FALSE;
	BOOL lb_isExplorer = FALSE;
	TCHAR szBuff[256]; 
	BOOL bFile = 0;		 //��ǰѡ�������� ( 1/2/3 ��ʾ���ļ�/�ļ���/��ǰѡ���˶��� )
	g_szPath[0] = TEXT('\0');	//ѡ���ļ���·��
	g_szItem[0] = TEXT('\0');	//ѡ���ļ����ļ���

	//��ȡ��ǰ�˴��ڵľ��
	HWND hExplorer = GetForegroundWindow();
	GetClassName(hExplorer,szBuff,sizeof(szBuff)/sizeof(TCHAR));

	//������������ϵͳ�ϵ�explorer.exe�������Ͷ���CabinetWClass/ExploreWClass����֪������ϵͳ�ǲ��ǣ�����֤
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
		//SWC_EXPLORER ָ����ȡexplorer�򿪵Ĵ��ڣ����û�д��κδ���������ʧ��  
		//VISTA���Ժ�İ汾������SWC_DESKTOP����ʱ��ȡ�������洰�ڣ���ʹû������shell ���ڣ��õ���Ҳ����ʧ��.  
		VARIANT v;
		V_VT(&v) = VT_I4;//�京���ǣ�����ָ��X��ȡ��ָ��Ķ����vt��Ա���� vpidl.vt = VT_I4;  
		v.ulVal = SWC_EXPLORER;//V_VT(&v)�����VT_I, ��VT_I4�����ͣ���ֵ��������ΪShell windows collection ��������
		BOOL fFound = FALSE; 
		long count = -1;
		IDispatch  *pdisp; 

		// ��ȡ�򿪵�shell���ڸ���.��ȡ��Ŀ���ܻ�����ж�if (pdisp == NULL)����
		if (SUCCEEDED(psw->get_Count(&count)))  
		{
			//#define V_I4(X)   V_UNION(X, lVal) 
			//#define V_UNION(X, Y)   ((X)->Y)
			for (V_I4(&v) = 0; !fFound && (V_I4(&v) < count); V_I4(&v)++)//V_I4(&v) == v->lVal
			{
				if (SUCCEEDED(psw->Item(v, &pdisp)))
				{
					//�����ȡ�������һ���ǻ�ȡ���ˣ�ֻ�д��ڵ�explorer�Ż���pdisp����������Ϊ��,�Ǿͻ�ȡ��һ������ȥ��ѯ�ӿڣ�
					if (pdisp == NULL)	
					{
						continue;
					}
					IWebBrowserApp *pwba;// ����ӿ���һ���ؼ��ӿ�
					if (SUCCEEDED(pdisp->QueryInterface(IID_IWebBrowserApp, (void**)&pwba))) 
					{
						HWND l_hwnd;
						if (SUCCEEDED(pwba->get_HWND((LONG_PTR*)&l_hwnd)) && l_hwnd == hExplorer) 
						{
							fFound = TRUE;
							IServiceProvider *psp;// �ӿ��๤�������ڴ�������һЩ�ӿڶ���
							if (SUCCEEDED(pwba->QueryInterface(IID_IServiceProvider, (void**)&psp))) 
							{
								// SID_STopLevelBrowserָ����ȡ�����Ǹ�shell����.  
								//���Ե��øýӿڵ�GetControlWindow��Ա������Ȼ������sp++�鿴һ�£���֪�����ĸ������ˡ�
								IShellBrowser *psb; //Shell�����
								if (SUCCEEDED(psp->QueryService(SID_STopLevelBrowser,IID_IShellBrowser, (void**)&psb))) 
								{
									IShellView *psv; //Shell��ͼ
									if (SUCCEEDED(psb->QueryActiveShellView(&psv))) 
									{
										IFolderView *pfv;//�ļ�����ͼ
										if (SUCCEEDED(psv->QueryInterface(IID_IFolderView,(void**)&pfv))) 
										{
											//////////////////////////////////////////////////////////////////////////										
											IPersistFolder2 *ppf2;
											if (SUCCEEDED(pfv->GetFolder(IID_IPersistFolder2,(void**)&ppf2))) {
												LPITEMIDLIST pidlFolder;
												if (SUCCEEDED(ppf2->GetCurFolder(&pidlFolder))) {
													//SHGetPathFromIDList �����ǰ���Ŀ��־���б�ת��Ϊ�ĵ�ϵͳ·��
													//pidlFolder--��� namespace �ĸ�ָ��һ���ĵ���Ŀ¼�ص��һ����Ŀ��ʶ����ĵ�ַ ( ���� ) ��
													if (!SHGetPathFromIDList(pidlFolder, g_szPath)) {   //��ȡ�ļ���·��
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
																			StrRetToBuf(&str, pidlItem, g_szItem, MAX_PATH);   //��ȡ�ļ��� 
																			bFile = ILIsFile(pidlItem);
																		}//SVSI_SELECT
																		psf->Release();
																	}
																	CoTaskMemFree(pidlItem);  //�˺��������ͷű�������ڴ�顣															
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

// �����ڡ������Ϣ�������
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
