#pragma once
#include <assert.h>
#include <string.h>
#include <Shlwapi.h>     
#pragma comment(lib, "Shlwapi.lib")    

//RegistryDll用来注册DLL，参数是DLL的全路径，返回值代表注册成功或者失败（TRUE为成功）。     
inline BOOL RegistryDll(CString& szDllPath)
{
	if (!(PathFileExists(szDllPath) && (!PathIsDirectory(szDllPath))))
	{
		throw wstring(L"注册") + szDllPath.GetBuffer() + L"文件的时候，发生错误：该文件不存在！";
	}
	LRESULT(CALLBACK* lpDllEntryPoint)();
	HINSTANCE hLib = LoadLibrary(szDllPath);
	if (hLib < (HINSTANCE)HINSTANCE_ERROR)
		return FALSE;
	(FARPROC&)lpDllEntryPoint = GetProcAddress(hLib, "DllRegisterServer");
	BOOL bRet = FALSE;
	if (lpDllEntryPoint != NULL)
	{
		HRESULT hr = (*lpDllEntryPoint)();
		bRet = SUCCEEDED(hr);
		if (FAILED(hr))
		{
			throw "注册dll时失败";
		}
	}
	FreeLibrary(hLib);
	return bRet;
}
inline unsigned int countGBK(const char * str)
{
	assert(str != NULL);
	unsigned int len = (unsigned int)strlen(str);
	unsigned int counter = 0;
	unsigned char head = 0x80;
	unsigned char firstChar, secondChar;

	for (unsigned int i = 0; i < len - 1; ++i)
	{
		firstChar = (unsigned char)str[i];
		if (!(firstChar & head))continue;
		secondChar = (unsigned char)str[i];
		if (firstChar >= 161 && firstChar <= 247 && secondChar >= 161 && secondChar <= 254)
		{
			counter += 2;
			++i;
		}
	}
	return counter;
}

inline unsigned int countUTF8(const char * str)
{
	assert(str != NULL);
	unsigned int len = (unsigned int)strlen(str);
	unsigned int counter = 0;
	unsigned char head = 0x80;
	unsigned char firstChar;
	for (unsigned int i = 0; i < len; ++i)
	{
		firstChar = (unsigned char)str[i];
		if (!(firstChar & head))continue;
		unsigned char tmpHead = head;
		unsigned int wordLen = 0, tPos = 0;
		while (firstChar & tmpHead)
		{
			++wordLen;
			tmpHead >>= 1;
		}
		if (wordLen <= 1)continue; //utf8最小长度为2  
		wordLen--;
		if (wordLen + i >= len)break;
		for (tPos = 1; tPos <= wordLen; ++tPos)
		{
			unsigned char secondChar = (unsigned char)str[i + tPos];
			if (!(secondChar & head))break;
		}
		if (tPos > wordLen)
		{
			counter += wordLen + 1;
			i += wordLen;
		}
	}
	return counter;
}

inline bool beUtf8(const char *str)
{
	//unsigned int iGBK = countGBK(str);
	//unsigned int iUTF8 = countUTF8(str);
	//if (iUTF8 > iGBK)return true;
	//return false;
	size_t length = strlen(str) + 1;
	size_t i;
	int nBytes;
	unsigned char chr;

	i = 0;
	nBytes = 0;
	while (i < length) {
		chr = *(str + i);

		if (nBytes == 0) { //计算字节数  
			if ((chr & 0x80) != 0) {
				while ((chr & 0x80) != 0) {
					chr <<= 1;
					nBytes++;
				}
				if ((nBytes < 2) || (nBytes > 6)) {
					return 0; //第一个字节最少为110x xxxx  
				}
				nBytes--; //减去自身占的一个字节  
			}
		}
		else { //多字节除了第一个字节外剩下的字节  
			if ((chr & 0xC0) != 0x80) {
				return 0; //剩下的字节都是10xx xxxx的形式  
			}
			nBytes--;
		}
		i++;
	}
	return (nBytes == 0);
}

inline char* U2G(const char* utf8)
{
	int len = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
	wchar_t* wstr = new wchar_t[len + 1];
	memset(wstr, 0, len + 1);
	MultiByteToWideChar(CP_UTF8, 0, utf8, -1, wstr, len);
	len = WideCharToMultiByte(CP_ACP, 0, wstr, -1, NULL, 0, NULL, NULL);
	char* str = new char[len + 1];
	memset(str, 0, len + 1);
	WideCharToMultiByte(CP_ACP, 0, wstr, -1, str, len, NULL, NULL);
	if (wstr) delete[] wstr;
	return str;
}

inline char* Unicode2UTF8(const wchar_t* unicode)
{
	int len = WideCharToMultiByte(CP_UTF8, 0, unicode, -1, NULL, 0, NULL, NULL);
	char* str = new char[len + 1];
	memset(str, 0, len + 1);
	WideCharToMultiByte(CP_UTF8, 0, unicode, -1, str, len, NULL, NULL);
	return str;

}