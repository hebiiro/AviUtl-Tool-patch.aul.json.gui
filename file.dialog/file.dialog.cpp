#include "framework.h"
#include "file.dialog.h"

int APIENTRY _tWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpCmdLine, int nCmdShow)
{
	_tsetlocale(LC_ALL, _T(""));

	int c = __argc;
	if (c < 3) return E_INVALIDARG;
	LPCTSTR argFlag = __targv[1];
	LPCTSTR argInitDir = __targv[2];

	TCHAR fileName[MAX_PATH] = _T("");
	OPENFILENAME ofn = { sizeof(ofn) };
	ofn.lpstrInitialDir = argInitDir;
	ofn.lpstrFilter = TEXT("JSON Files\0*.json\0All Files\0*.*\0");
	ofn.lpstrFile = fileName;
	ofn.nMaxFile = MAX_PATH;
	if (::lstrcmpi(argFlag, TEXT("-Open")) == 0)
	{
		ofn.Flags = OFN_FILEMUSTEXIST;
		if (!::GetOpenFileName(&ofn)) return S_FALSE;
	}
	else if (::lstrcmpi(argFlag, TEXT("-Save")) == 0)
	{
		ofn.Flags = OFN_OVERWRITEPROMPT | OFN_HIDEREADONLY;
		if (!::GetSaveFileName(&ofn)) return S_FALSE;
	}
	else
	{
		return E_INVALIDARG;
	}

	_tprintf(fileName);
	return S_OK;
}
