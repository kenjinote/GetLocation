#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib, "locationapi.lib")

#include <windows.h>
#include <LocationApi.h>
#include <wchar.h>
#include <shlwapi.h>
#include <new>

#define DEFAULT_WAIT_FOR_LOCATION_REPORT 500
TCHAR szClassName[] = TEXT("Window");

DEFINE_PROPERTYKEY(SENSOR_DATA_TYPE_LATITUDE_DEGREES, 0x055c74d8, 0xca6f, 0x47d6, 0x95, 0xc6, 0x1e, 0xd3, 0x63, 0x7a, 0x0f, 0xf4, 2);
DEFINE_PROPERTYKEY(SENSOR_DATA_TYPE_LONGITUDE_DEGREES, 0x055c74d8, 0xca6f, 0x47d6, 0x95, 0xc6, 0x1e, 0xd3, 0x63, 0x7a, 0x0f, 0xf4, 3);

class CLocationCallback : public ILocationEvents
{
public:
	CLocationCallback() : _cRef(1), _hDataEvent(::CreateEvent(
		NULL,
		FALSE,
		FALSE,
		NULL))
	{
	}

	virtual ~CLocationCallback()
	{
		if (_hDataEvent)
		{
			::CloseHandle(_hDataEvent);
		}
	}

	IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv)
	{
		if ((riid == IID_IUnknown) ||
			(riid == IID_ILocationEvents))
		{
			*ppv = static_cast<ILocationEvents*>(this);
		}
		else
		{
			*ppv = NULL;
			return E_NOINTERFACE;
		}
		AddRef();
		return S_OK;
	}

	IFACEMETHODIMP_(ULONG) AddRef()
	{
		return InterlockedIncrement(&_cRef);
	}

	IFACEMETHODIMP_(ULONG) Release()
	{
		long cRef = InterlockedDecrement(&_cRef);
		if (!cRef)
		{
			delete this;
		}
		return cRef;
	}

	IFACEMETHODIMP OnLocationChanged(REFIID /*reportType*/, ILocationReport* /*pLocationReport*/)
	{
		::SetEvent(_hDataEvent);
		return S_OK;
	}

	IFACEMETHODIMP OnStatusChanged(REFIID /*reportType*/, LOCATION_REPORT_STATUS status)
	{
		if (REPORT_RUNNING == status)
		{
			::SetEvent(_hDataEvent);
		}
		return S_OK;
	}

	HANDLE GetEventHandle()
	{
		return _hDataEvent;
	}

private:
	long _cRef;
	HANDLE _hDataEvent;
};

HRESULT WaitForLocationReport(
	ILocation* pLocation, 
	REFIID reportType,
	DWORD dwTimeToWait,
	ILocationReport **ppLocationReport
)
{
	*ppLocationReport = NULL;

	CLocationCallback *pLocationCallback = new(std::nothrow) CLocationCallback();
	HRESULT hr = pLocationCallback ? S_OK : E_OUTOFMEMORY;
	if (SUCCEEDED(hr))
	{
		HANDLE hEvent = pLocationCallback->GetEventHandle();
		hr = hEvent ? S_OK : E_FAIL;
		if (SUCCEEDED(hr))
		{
			hr = pLocation->RegisterForReport(pLocationCallback, reportType, 0);
			if (SUCCEEDED(hr))
			{
				DWORD dwIndex;
				HRESULT hrWait = CoWaitForMultipleHandles(0, dwTimeToWait, 1, &hEvent, &dwIndex);
				if ((S_OK == hrWait) || (RPC_S_CALLPENDING == hrWait))
				{
					Sleep(5000);
					hr = pLocation->GetReport(reportType, ppLocationReport);
					if (FAILED(hr) && (RPC_S_CALLPENDING == hrWait))
					{
						hr = hrWait;
					}
				}
				pLocation->UnregisterForReport(reportType);
			}
		}
		pLocationCallback->Release();
	}
	return hr;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hButton;
	static HWND hEdit1;
	static HWND hEdit2;
	switch (msg)
	{
	case WM_CREATE:
		if (FAILED(CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE)))
		{
			return -1;
		}
		hButton = CreateWindow(TEXT("BUTTON"), TEXT("取得"), WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hWnd, (HMENU)IDOK, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hEdit1 = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | ES_READONLY, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hEdit2 = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | ES_READONLY, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		break;
	case WM_SIZE:
		MoveWindow(hButton, 10, 10, 256, 32, TRUE);
		MoveWindow(hEdit1, 10, 50, 256, 32, TRUE);
		MoveWindow(hEdit2, 10, 90, 256, 32, TRUE);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK)
		{
			SetWindowText(hEdit1, 0);
			SetWindowText(hEdit2, 0);
			if (TRUE)
			{
				DWORD const dwTimeToWait = DEFAULT_WAIT_FOR_LOCATION_REPORT;

				ILocation *pLocation;
				HRESULT hr = CoCreateInstance(CLSID_Location, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&pLocation));
				if (SUCCEEDED(hr))
				{
					IID REPORT_TYPES[] = { IID_ILatLongReport };
					if (FAILED(pLocation->RequestPermissions(NULL, REPORT_TYPES, ARRAYSIZE(REPORT_TYPES), TRUE)))
					{
						//wprintf_s(L"Warning: Unable to request permissions.\n");
					}
					ILocationReport *pLocationReport;
					hr = ::WaitForLocationReport(pLocation, IID_ILatLongReport, dwTimeToWait, &pLocationReport);
					if (SUCCEEDED(hr))
					{
						TCHAR szText[1024];
						PROPVARIANT variant;
						PropVariantInit(&variant);
						pLocationReport->GetValue(SENSOR_DATA_TYPE_LATITUDE_DEGREES, &variant);
						swprintf_s(szText, TEXT("%f"), variant.dblVal);
						SetWindowText(hEdit1, szText);
						PropVariantClear(&variant);
						PropVariantInit(&variant);
						pLocationReport->GetValue(SENSOR_DATA_TYPE_LONGITUDE_DEGREES, &variant);
						swprintf_s(szText, TEXT("%f"), variant.dblVal);
						SetWindowText(hEdit2, szText);
						PropVariantClear(&variant);
						pLocationReport->Release();
					}
					//else if (RPC_S_CALLPENDING == hr)
					//{
					//	wprintf_s(L"No LatLong data received.  Wait time of %lu elapsed.\n", dwTimeToWait);
					//}
					pLocation->Release();
				}
			}
		}
		break;
	case WM_DESTROY:
		::CoUninitialize();
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("位置情報を取得"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}
