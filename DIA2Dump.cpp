// Dia2Dump.cpp : Defines the entry point for the console application.
//
// This is a part of the Debug Interface Access SDK
// Copyright (c) Microsoft Corporation.  All rights reserved.
//
// This source code is only intended as a supplement to the
// Debug Interface Access SDK and related electronic documentation
// provided with the library.
// See these sources for detailed information regarding the
// Debug Interface Access SDK API.
//

#include "stdafx.h"
#include "Dia2Dump.h"
#include "PrintSymbol.h"

#include "Callback.h"

#pragma warning (disable : 4100)

const wchar_t * g_szFilename;
IDiaDataSource * g_pDiaDataSource;
IDiaSession * g_pDiaSession;
IDiaSymbol * g_pGlobalSymbol;
DWORD g_dwMachineType = CV_CFL_80386;
ULONGLONG g_dwloadAddress = 0x400000;

#include <fstream>
#include <iostream>
#include <sys/stat.h>
#include <algorithm>
#include <set>
#include <ctime>
#include <map>
#include <string>
#include <unordered_map>
#include <vector>

#include "Shlwapi.h"
#pragma comment(lib, "Shlwapi.lib")
int __cdecl wmain2(int argc, wchar_t * argv[]);
bool dump_specifc_dwords = false;

////////////////////////////////////////////////////////////
//
int wmain(int argc, wchar_t * argv[])
{
	FILE * pFile;

	if (argc < 2) {
		PrintHelpOptions();
		return -1;
	}

	if (_wfopen_s(&pFile, argv[argc - 1], L"r") || !pFile) {
	  // invalid file name or file does not exist
		wprintf(L"Can't open file %s\n", argv[argc - 1]);
		PrintHelpOptions();
		return -1;
	}

	fclose(pFile);

	if (argv[1] != nullptr && argv[2] != nullptr) {
		if (wcsstr(argv[1], L".pdb") != nullptr && wcsstr(argv[2], L".pdb") != nullptr || wcsstr(argv[1], L"-ptype")) {
			wprintf(L"Comparing PDBs\n");
			return wmain2(argc, argv);
		}

		if (argv[1] != nullptr && wcsstr(argv[2], L".exe") != nullptr && wcsstr(argv[3], L".pdb") != nullptr) {
			dump_specifc_dwords = true;
		}
	}

	g_szFilename = argv[argc - 1];

	// CoCreate() and initialize COM objects

	if (!LoadDataFromPdb(g_szFilename, &g_pDiaDataSource, &g_pDiaSession, &g_pGlobalSymbol)) {
		return -1;
	}

	if (argc == 4 && dump_specifc_dwords) {
		DumpAllSpecificDwords(g_pDiaSession, argv[2], argv[1]);
	} else if (argc == 2) {
	  // no options passed; print all pdb info

		DumpAllPdbInfo(g_pDiaSession, g_pGlobalSymbol);
	}

	else if (!_wcsicmp(argv[1], L"-all")) {
		DumpAllPdbInfo(g_pDiaSession, g_pGlobalSymbol);
	}

	else if (!ParseArg(argc - 2, &argv[1])) {
		Cleanup();

		return -1;
	}

	// release COM objects and CoUninitialize()

	Cleanup();

	return 0;
}

////////////////////////////////////////////////////////////
// Create an IDiaData source and open a PDB file
//
bool LoadDataFromPdb(
	const wchar_t * szFilename,
	IDiaDataSource ** ppSource,
	IDiaSession ** ppSession,
	IDiaSymbol ** ppGlobal)
{
	wchar_t wszExt[MAX_PATH];
	const wchar_t * wszSearchPath = L"SRV**\\\\symbols\\symbols"; // Alternate path to search for debug data
	DWORD dwMachType = 0;

	HRESULT hr = CoInitialize(NULL);

	// Obtain access to the provider

	hr = CoCreateInstance(__uuidof(DiaSource),
						  NULL,
						  CLSCTX_INPROC_SERVER,
						  __uuidof(IDiaDataSource),
						  (void **)ppSource);

	if (FAILED(hr)) {
		wprintf(L"CoCreateInstance failed - HRESULT = %08X\n", hr);

		return false;
	}

	_wsplitpath_s(szFilename, NULL, 0, NULL, 0, NULL, 0, wszExt, MAX_PATH);

	if (!_wcsicmp(wszExt, L".pdb")) {
	  // Open and prepare a program database (.pdb) file as a debug data source

		hr = (*ppSource)->loadDataFromPdb(szFilename);

		if (FAILED(hr)) {
			wprintf(L"loadDataFromPdb failed - HRESULT = %08X\n", hr);

			return false;
		}
	}

	else {
		CCallback callback; // Receives callbacks from the DIA symbol locating procedure,
							// thus enabling a user interface to report on the progress of
							// the location attempt. The client application may optionally
							// provide a reference to its own implementation of this
							// virtual base class to the IDiaDataSource::loadDataForExe method.
		callback.AddRef();

		// Open and prepare the debug data associated with the executable

		hr = (*ppSource)->loadDataForExe(szFilename, wszSearchPath, &callback);

		if (FAILED(hr)) {
			wprintf(L"loadDataForExe failed - HRESULT = %08X\n", hr);

			return false;
		}
	}

	// Open a session for querying symbols

	hr = (*ppSource)->openSession(ppSession);

	if (FAILED(hr)) {
		wprintf(L"openSession failed - HRESULT = %08X\n", hr);

		return false;
	}

	// Retrieve a reference to the global scope

	hr = (*ppSession)->get_globalScope(ppGlobal);

	if (hr != S_OK) {
		wprintf(L"get_globalScope failed\n");

		return false;
	}

	// Set Machine type for getting correct register names

	if ((*ppGlobal)->get_machineType(&dwMachType) == S_OK) {
		switch (dwMachType) {
			case IMAGE_FILE_MACHINE_I386: g_dwMachineType = CV_CFL_80386; break;
			case IMAGE_FILE_MACHINE_IA64: g_dwMachineType = CV_CFL_IA64; break;
			case IMAGE_FILE_MACHINE_AMD64: g_dwMachineType = CV_CFL_AMD64; break;
		}
	}

	return true;
}

////////////////////////////////////////////////////////////
// Release DIA objects and CoUninitialize
//
void Cleanup()
{
	if (g_pGlobalSymbol) {
		g_pGlobalSymbol->Release();
		g_pGlobalSymbol = NULL;
	}

	if (g_pDiaSession) {
		g_pDiaSession->Release();
		g_pDiaSession = NULL;
	}

	CoUninitialize();
}

////////////////////////////////////////////////////////////
// Parse the arguments of the program
//
bool ParseArg(int argc, wchar_t * argv[])
{
	int iCount = 0;
	bool bReturn = true;

	if (!argc) {
		return true;
	}

	if (!_wcsicmp(argv[0], L"-?")) {
		PrintHelpOptions();

		return true;
	}

	else if (!_wcsicmp(argv[0], L"-help")) {
		PrintHelpOptions();

		return true;
	}

	else if (!_wcsicmp(argv[0], L"-m")) {
	  // -m                : print all the mods

		iCount = 1;
		bReturn = bReturn && DumpAllMods(g_pGlobalSymbol);
		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-p")) {
	  // -p                : print all the publics

		iCount = 1;
		bReturn = bReturn && DumpAllPublics(g_pGlobalSymbol);
		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-s")) {
	  // -s                : print symbols

		iCount = 1;
		bReturn = bReturn && DumpAllSymbols(g_pGlobalSymbol);
		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-g")) {
	  // -g                : print all the globals

		iCount = 1;
		bReturn = bReturn && DumpAllGlobals(g_pGlobalSymbol);
		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-t")) {
	  // -t                : print all the types

		iCount = 1;
		bReturn = bReturn && DumpAllTypes(g_pGlobalSymbol);
		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-f")) {
	  // -f                : print all the files

		iCount = 1;
		bReturn = bReturn && DumpAllFiles(g_pDiaSession, g_pGlobalSymbol);
		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-l")) {
		if (argc > 1 && *argv[1] != L'-') {
		  // -l RVA [bytes]    : print line number info at RVA address in the bytes range

			DWORD dwRVA = 0;
			DWORD dwRange = MAX_RVA_LINES_BYTES_RANGE;

			swscanf_s(argv[1], L"%x", &dwRVA);
			if (argc > 2 && *argv[2] != L'-') {
				swscanf_s(argv[2], L"%d", &dwRange);
				iCount = 3;
			}

			else {
				iCount = 2;
			}

			bReturn = bReturn && DumpAllLines(g_pDiaSession, dwRVA, dwRange);
		}

		else {
		  // -l            : print line number info

			bReturn = bReturn && DumpAllLines(g_pDiaSession, g_pGlobalSymbol);
			iCount = 1;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-c")) {
	  // -c                : print section contribution info

		iCount = 1;
		bReturn = bReturn && DumpAllSecContribs(g_pDiaSession);
		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-dbg")) {
	  // -dbg              : dump debug streams

		iCount = 1;
		bReturn = bReturn && DumpAllDebugStreams(g_pDiaSession);
		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-injsrc")) {
		if (argc > 1 && *argv[1] != L'-') {
		  // -injsrc filename          : dump injected source filename

			bReturn = bReturn && DumpInjectedSource(g_pDiaSession, argv[1]);
			iCount = 2;
		}

		else {
		  // -injsrc           : dump all injected source

			bReturn = bReturn && DumpAllInjectedSources(g_pDiaSession);
			iCount = 1;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-sf")) {
	  // -sf               : dump all source files

		iCount = 1;
		bReturn = bReturn && DumpAllSourceFiles(g_pDiaSession, g_pGlobalSymbol);
		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-oem")) {
	  // -oem              : dump all OEM specific types

		iCount = 1;
		bReturn = bReturn && DumpAllOEMs(g_pGlobalSymbol);
		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-fpo")) {
		if (argc > 1 && *argv[1] != L'-') {
			DWORD dwRVA = 0;

			if (iswdigit(*argv[1])) {
			  // -fpo [RVA]        : dump frame pointer omission information for a function address

				swscanf_s(argv[1], L"%x", &dwRVA);
				bReturn = bReturn && DumpFPO(g_pDiaSession, dwRVA);
			}

			else {
			  // -fpo [symbolname] : dump frame pointer omission information for a function symbol

				bReturn = bReturn && DumpFPO(g_pDiaSession, g_pGlobalSymbol, argv[1]);
			}

			iCount = 2;
		}

		else {
			bReturn = bReturn && DumpAllFPO(g_pDiaSession);
			iCount = 1;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-compiland")) {
		if ((argc > 1) && (*argv[1] != L'-')) {
		  // -compiland [name] : dump symbols for this compiland

			bReturn = bReturn && DumpCompiland(g_pGlobalSymbol, argv[1]);
			argc -= 2;
		}

		else {
			wprintf(L"ERROR - ParseArg(): missing argument for option '-line'");

			return false;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-compcontr")) {
		if ((argc > 1) && (*argv[1] != L'-')) {
		  // -compiland [name] : dump symbols for this compiland

			bReturn = bReturn && DumpCompilandContrib(g_pDiaSession,g_pGlobalSymbol, argv[1]);
			argc -= 2;
		}

		else {
			wprintf(L"ERROR - ParseArg(): missing argument for option '-line'");

			return false;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-lines")) {
		if ((argc > 1) && (*argv[1] != L'-')) {
			DWORD dwRVA = 0;

			if (iswdigit(*argv[1])) {
			  // -lines <RVA>                  : dump line numbers for this address\n"

				swscanf_s(argv[1], L"%x", &dwRVA);
				bReturn = bReturn && DumpLines(g_pDiaSession, dwRVA);
			}

			else {
			  // -lines <symbolname>           : dump line numbers for this function

				bReturn = bReturn && DumpLines(g_pDiaSession, g_pGlobalSymbol, argv[1]);
			}

			iCount = 2;
		}

		else {
			wprintf(L"ERROR - ParseArg(): missing argument for option '-compiland'");

			return false;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-type")) {
	  // -type <symbolname>: dump this type in detail

		if ((argc > 1) && (*argv[1] != L'-')) {
			bReturn = bReturn && DumpType(g_pGlobalSymbol, argv[1]);
			iCount = 2;
		}

		else {
			wprintf(L"ERROR - ParseArg(): missing argument for option '-type'");

			return false;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-label")) {
	  // -label <RVA>      : dump label at RVA
		if ((argc > 1) && (*argv[1] != L'-')) {
			DWORD dwRVA = 0;

			swscanf_s(argv[1], L"%x", &dwRVA);
			bReturn = bReturn && DumpLabel(g_pDiaSession, dwRVA);
			iCount = 2;
		}

		else {
			wprintf(L"ERROR - ParseArg(): missing argument for option '-label'");

			return false;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-sym")) {
		if ((argc > 1) && (*argv[1] != L'-')) {
			DWORD dwRVA = 0;
			const wchar_t * szChildname = NULL;

			iCount = 2;

			if (argc > 2 && *argv[2] != L'-') {
				szChildname = argv[2];
				iCount = 3;
			}

			if (iswdigit(*argv[1])) {
			  // -sym <RVA> [childname]        : dump child information of symbol at this address

				swscanf_s(argv[1], L"%x", &dwRVA);
				bReturn = bReturn && DumpSymbolWithRVA(g_pDiaSession, dwRVA, szChildname);
			}

			else {
			  // -sym <symbolname> [childname] : dump child information of this symbol

				bReturn = bReturn && DumpSymbolsWithRegEx(g_pGlobalSymbol, argv[1], szChildname);
			}
		}

		else {
			wprintf(L"ERROR - ParseArg(): missing argument for option '-sym'");

			return false;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-lsrc")) {
	  // -lsrc  <file> [line]          : dump line numbers for this source file

		if ((argc > 1) && (*argv[1] != L'-')) {
			DWORD dwLine = 0;

			iCount = 2;
			if (argc > 2 && *argv[2] != L'-') {
				swscanf_s(argv[1], L"%d", &dwLine);
				iCount = 3;
			}

			bReturn = bReturn && DumpLinesForSourceFile(g_pDiaSession, argv[1], dwLine);
		}

		else {
			wprintf(L"ERROR - ParseArg(): missing argument for option '-lsrc'");

			return false;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-ps")) {
	  // -ps <RVA> [-n <number>]       : dump symbols after this address, default 16

		if ((argc > 1) && (*argv[1] != L'-')) {
			DWORD dwRVA = 0;
			DWORD dwRange;

			swscanf_s(argv[1], L"%x", &dwRVA);
			if (argc > 3 && !_wcsicmp(argv[2], L"-n")) {
				swscanf_s(argv[3], L"%d", &dwRange);
				iCount = 4;
			}

			else {
				dwRange = 16;
				iCount = 2;
			}

			bReturn = bReturn && DumpPublicSymbolsSorted(g_pDiaSession, dwRVA, dwRange, true);
		}

		else {
			wprintf(L"ERROR - ParseArg(): missing argument for option '-ps'");

			return false;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-psr")) {
	  // -psr <RVA> [-n <number>]       : dump symbols before this address, default 16

		if ((argc > 1) && (*argv[1] != L'-')) {
			DWORD dwRVA = 0;
			DWORD dwRange;

			swscanf_s(argv[1], L"%x", &dwRVA);

			if (argc > 3 && !_wcsicmp(argv[2], L"-n")) {
				swscanf_s(argv[3], L"%d", &dwRange);
				iCount = 4;
			}

			else {
				dwRange = 16;
				iCount = 2;
			}

			bReturn = bReturn && DumpPublicSymbolsSorted(g_pDiaSession, dwRVA, dwRange, false);
		}

		else {
			wprintf(L"ERROR - ParseArg(): missing argument for option '-psr'");

			return false;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-annotations")) {
	  // -annotations <RVA>: dump annotation symbol for this RVA

		if ((argc > 1) && (*argv[1] != L'-')) {
			DWORD dwRVA = 0;

			swscanf_s(argv[1], L"%x", &dwRVA);
			bReturn = bReturn && DumpAnnotations(g_pDiaSession, dwRVA);
			iCount = 2;
		}

		else {
			wprintf(L"ERROR - ParseArg(): missing argument for option '-maptosrc'");

			return false;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-maptosrc")) {
	  // -maptosrc <RVA>   : dump src RVA for this image RVA

		if ((argc > 1) && (*argv[1] != L'-')) {
			DWORD dwRVA = 0;

			swscanf_s(argv[1], L"%x", &dwRVA);
			bReturn = bReturn && DumpMapToSrc(g_pDiaSession, dwRVA);
			iCount = 2;
		}

		else {
			wprintf(L"ERROR - ParseArg(): missing argument for option '-maptosrc'");

			return false;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else if (!_wcsicmp(argv[0], L"-mapfromsrc")) {
	  // -mapfromsrc <RVA> : dump image RVA for src RVA

		if ((argc > 1) && (*argv[1] != L'-')) {
			DWORD dwRVA = 0;

			swscanf_s(argv[1], L"%x", &dwRVA);
			bReturn = bReturn && DumpMapFromSrc(g_pDiaSession, dwRVA);
			iCount = 2;
		}

		else {
			wprintf(L"ERROR - ParseArg(): missing argument for option '-mapfromsrc'");

			return false;
		}

		argc -= iCount;
		bReturn = bReturn && ParseArg(argc, &argv[iCount]);
	}

	else {
		wprintf(L"ERROR - unknown option %s\n", argv[0]);

		PrintHelpOptions();

		return false;
	}

	return bReturn;
}

////////////////////////////////////////////////////////////
// Display the usage
//
void PrintHelpOptions()
{
	static const wchar_t * const helpString = L"usage: Dia2Dump.exe [ options ] <filename>\n"
		L"  -?                : print this help\n"
		L"  -all              : print all the debug info\n"
		L"  -m                : print all the mods\n"
		L"  -p                : print all the publics\n"
		L"  -g                : print all the globals\n"
		L"  -t                : print all the types\n"
		L"  -f                : print all the files\n"
		L"  -s                : print symbols\n"
		L"  -l [RVA [bytes]]  : print line number info at RVA address in the bytes range\n"
		L"  -c                : print section contribution info\n"
		L"  -dbg              : dump debug streams\n"
		L"  -injsrc [file]    : dump injected source\n"
		L"  -sf               : dump all source files\n"
		L"  -oem              : dump all OEM specific types\n"
		L"  -fpo [RVA]        : dump frame pointer omission information for a func addr\n"
		L"  -fpo [symbolname] : dump frame pointer omission information for a func symbol\n"
		L"  -compiland [name] : dump symbols for this compiland\n"
		L"  -compcontr [name] : dump symbols for this compiland contrib\n"
		L"  -lines <funcname> : dump line numbers for this function\n"
		L"  -lines <RVA>      : dump line numbers for this address\n"
		L"  -type <symbolname>: dump this type in detail\n"
		L"  -label <RVA>      : dump label at RVA\n"
		L"  -sym <symbolname> [childname] : dump child information of this symbol\n"
		L"  -sym <RVA> [childname]        : dump child information of symbol at this addr\n"
		L"  -lsrc  <file> [line]          : dump line numbers for this source file\n"
		L"  -ps <RVA> [-n <number>]       : dump symbols after this address, default 16\n"
		L"  -psr <RVA> [-n <number>]      : dump symbols before this address, default 16\n"
		L"  -annotations <RVA>: dump annotation symbol for this RVA\n"
		L"  -maptosrc <RVA>   : dump src RVA for this image RVA\n"
		L"  -mapfromsrc <RVA> : dump image RVA for src RVA\n"
		L"  Or Specify two pdbs to compare types in them\n"
		L"  Or Specify a typename, exe and pdb to print specific dwords\n"
		;

	wprintf(helpString);
}

////////////////////////////////////////////////////////////
// Dump all the data stored in a PDB
//
void DumpAllPdbInfo(IDiaSession * pSession, IDiaSymbol * pGlobal)
{
	DumpAllMods(pGlobal);
	DumpAllPublics(pGlobal);
	DumpAllSymbols(pGlobal);
	DumpAllGlobals(pGlobal);
	DumpAllTypes(pGlobal);
	DumpAllFiles(pSession, pGlobal);
	DumpAllLines(pSession, pGlobal);
	DumpAllSecContribs(pSession);
	DumpAllDebugStreams(pSession);
	DumpAllInjectedSources(pSession);
	DumpAllFPO(pSession);
	DumpAllOEMs(pGlobal);
}

////////////////////////////////////////////////////////////
// Dump all the modules information
//
bool DumpAllMods(IDiaSymbol * pGlobal)
{
	wprintf(L"\n\n*** MODULES\n\n");

	// Retrieve all the compiland symbols

	IDiaEnumSymbols * pEnumSymbols;

	if (FAILED(pGlobal->findChildren(SymTagCompiland, NULL, nsNone, &pEnumSymbols))) {
		return false;
	}

	IDiaSymbol * pCompiland;
	ULONG celt = 0;
	ULONG iMod = 1;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pCompiland, &celt)) && (celt == 1)) {
		BSTR bstrName;

		if (pCompiland->get_name(&bstrName) != S_OK) {
			wprintf(L"ERROR - Failed to get the compiland's name\n");

			pCompiland->Release();
			pEnumSymbols->Release();

			return false;
		}

		wprintf(L"%04X %s\n", iMod++, bstrName);

		// Deallocate the string allocated previously by the call to get_name

		SysFreeString(bstrName);

		pCompiland->Release();
	}

	pEnumSymbols->Release();

	putwchar(L'\n');

	return true;
}

////////////////////////////////////////////////////////////
// Dump all the public symbols - SymTagPublicSymbol
//
bool DumpAllPublics(IDiaSymbol * pGlobal)
{
	wprintf(L"\n\n*** PUBLICS\n\n");

	// Retrieve all the public symbols

	IDiaEnumSymbols * pEnumSymbols;

	if (FAILED(pGlobal->findChildren(SymTagPublicSymbol, NULL, nsNone, &pEnumSymbols))) {
		return false;
	}

	IDiaSymbol * pSymbol;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pSymbol, &celt)) && (celt == 1)) {
		PrintPublicSymbol(pSymbol);

		pSymbol->Release();
	}

	pEnumSymbols->Release();

	putwchar(L'\n');

	return true;
}

////////////////////////////////////////////////////////////
// Dump all the symbol information stored in the compilands
//
bool DumpAllSymbols(IDiaSymbol * pGlobal)
{
	wprintf(L"\n\n*** SYMBOLS\n\n\n");

	// Retrieve the compilands first

	IDiaEnumSymbols * pEnumSymbols;

	if (FAILED(pGlobal->findChildren(SymTagCompiland, NULL, nsNone, &pEnumSymbols))) {
		return false;
	}

	IDiaSymbol * pCompiland;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pCompiland, &celt)) && (celt == 1)) {
		wprintf(L"\n** Module: ");

		// Retrieve the name of the module

		BSTR bstrName;

		if (pCompiland->get_name(&bstrName) != S_OK) {
			wprintf(L"(???)\n\n");
		}

		else {
			wprintf(L"%s\n\n", bstrName);

			SysFreeString(bstrName);
		}

		// Find all the symbols defined in this compiland and print their info

		IDiaEnumSymbols * pEnumChildren;

		if (SUCCEEDED(pCompiland->findChildren(SymTagNull, NULL, nsNone, &pEnumChildren))) {
			IDiaSymbol * pSymbol;
			ULONG celtChildren = 0;

			while (SUCCEEDED(pEnumChildren->Next(1, &pSymbol, &celtChildren)) && (celtChildren == 1)) {
				PrintSymbol(pSymbol, 0);
				pSymbol->Release();
			}

			pEnumChildren->Release();
		}

		pCompiland->Release();
	}

	pEnumSymbols->Release();

	putwchar(L'\n');

	return true;
}

////////////////////////////////////////////////////////////
// Dump all the global symbols - SymTagFunction,
//  SymTagThunk and SymTagData
//
bool DumpAllGlobals(IDiaSymbol * pGlobal)
{
	IDiaEnumSymbols * pEnumSymbols;
	IDiaSymbol * pSymbol;
	enum SymTagEnum dwSymTags[] = {SymTagFunction, SymTagThunk, SymTagData};
	ULONG celt = 0;

	wprintf(L"\n\n*** GLOBALS\n\n");

	for (size_t i = 0; i < _countof(dwSymTags); i++, pEnumSymbols = NULL) {
		if (SUCCEEDED(pGlobal->findChildren(dwSymTags[i], NULL, nsNone, &pEnumSymbols))) {
			while (SUCCEEDED(pEnumSymbols->Next(1, &pSymbol, &celt)) && (celt == 1)) {
				PrintGlobalSymbol(pSymbol);

				pSymbol->Release();
			}

			pEnumSymbols->Release();
		}

		else {
			wprintf(L"ERROR - DumpAllGlobals() returned no symbols\n");

			return false;
		}
	}

	putwchar(L'\n');

	return true;
}

////////////////////////////////////////////////////////////
// Dump all the types information
//  (type symbols can be UDTs, enums or typedefs)
//
bool DumpAllTypes(IDiaSymbol * pGlobal)
{
	wprintf(L"\n\n*** TYPES\n");

	bool f1 = DumpAllUDTs(pGlobal);
	bool f2 = DumpAllEnums(pGlobal);
	bool f3 = DumpAllTypedefs(pGlobal);

	return f1 && f2 && f3;
}

////////////////////////////////////////////////////////////
// Dump all the user defined types (UDT)
//
bool DumpAllUDTs(IDiaSymbol * pGlobal)
{
	wprintf(L"\n\n** User Defined Types\n\n");

	IDiaEnumSymbols * pEnumSymbols;

	if (FAILED(pGlobal->findChildren(SymTagUDT, NULL, nsNone, &pEnumSymbols))) {
		wprintf(L"ERROR - DumpAllUDTs() returned no symbols\n");

		return false;
	}

	IDiaSymbol * pSymbol;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pSymbol, &celt)) && (celt == 1)) {
		PrintTypeInDetail(pSymbol, 0);

		pSymbol->Release();
	}

	pEnumSymbols->Release();

	putwchar(L'\n');

	return true;
}

////////////////////////////////////////////////////////////
// Dump all the enum types from the pdb
//
bool DumpAllEnums(IDiaSymbol * pGlobal)
{
	wprintf(L"\n\n** ENUMS\n\n");

	IDiaEnumSymbols * pEnumSymbols;

	if (FAILED(pGlobal->findChildren(SymTagEnum, NULL, nsNone, &pEnumSymbols))) {
		wprintf(L"ERROR - DumpAllEnums() returned no symbols\n");

		return false;
	}

	IDiaSymbol * pSymbol;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pSymbol, &celt)) && (celt == 1)) {
		PrintTypeInDetail(pSymbol, 0);

		pSymbol->Release();
	}

	pEnumSymbols->Release();

	putwchar(L'\n');

	return true;
}

////////////////////////////////////////////////////////////
// Dump all the typedef types from the pdb
//
bool DumpAllTypedefs(IDiaSymbol * pGlobal)
{
	wprintf(L"\n\n** TYPEDEFS\n\n");

	IDiaEnumSymbols * pEnumSymbols;

	if (FAILED(pGlobal->findChildren(SymTagTypedef, NULL, nsNone, &pEnumSymbols))) {
		wprintf(L"ERROR - DumpAllTypedefs() returned no symbols\n");

		return false;
	}

	IDiaSymbol * pSymbol;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pSymbol, &celt)) && (celt == 1)) {
		PrintTypeInDetail(pSymbol, 0);

		pSymbol->Release();
	}

	pEnumSymbols->Release();

	putwchar(L'\n');

	return true;
}

////////////////////////////////////////////////////////////
// Dump OEM specific types
//
bool DumpAllOEMs(IDiaSymbol * pGlobal)
{
	wprintf(L"\n\n*** OEM Specific types\n\n");

	IDiaEnumSymbols * pEnumSymbols;

	if (FAILED(pGlobal->findChildren(SymTagCustomType, NULL, nsNone, &pEnumSymbols))) {
		wprintf(L"ERROR - DumpAllOEMs() returned no symbols\n");

		return false;
	}

	IDiaSymbol * pSymbol;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pSymbol, &celt)) && (celt == 1)) {
		PrintTypeInDetail(pSymbol, 0);

		pSymbol->Release();
	}

	pEnumSymbols->Release();

	putwchar(L'\n');

	return true;
}

////////////////////////////////////////////////////////////
// For each compiland in the PDB dump all the source files
//
bool DumpAllFiles(IDiaSession * pSession, IDiaSymbol * pGlobal)
{
	wprintf(L"\n\n*** FILES\n\n");

	// In order to find the source files, we have to look at the image's compilands/modules

	IDiaEnumSymbols * pEnumSymbols;

	if (FAILED(pGlobal->findChildren(SymTagCompiland, NULL, nsNone, &pEnumSymbols))) {
		return false;
	}

	IDiaSymbol * pCompiland;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pCompiland, &celt)) && (celt == 1)) {
		BSTR bstrName;

		if (pCompiland->get_name(&bstrName) == S_OK) {
			wprintf(L"\nCompiland = %s\n\n", bstrName);

			SysFreeString(bstrName);
		}

		// Every compiland could contain multiple references to the source files which were used to build it
		// Retrieve all source files by compiland by passing NULL for the name of the source file

		IDiaEnumSourceFiles * pEnumSourceFiles;

		if (SUCCEEDED(pSession->findFile(pCompiland, NULL, nsNone, &pEnumSourceFiles))) {
			IDiaSourceFile * pSourceFile;

			while (SUCCEEDED(pEnumSourceFiles->Next(1, &pSourceFile, &celt)) && (celt == 1)) {
				PrintSourceFile(pSourceFile, nullptr);
				putwchar(L'\n');

				pSourceFile->Release();
			}

			pEnumSourceFiles->Release();
		}

		putwchar(L'\n');

		pCompiland->Release();
	}

	pEnumSymbols->Release();

	return true;
}

////////////////////////////////////////////////////////////
// Dump all the line numbering information contained in the PDB
//  Only function symbols have corresponding line numbering information
bool DumpAllLines(IDiaSession * pSession, IDiaSymbol * pGlobal)
{
	wprintf(L"\n\n*** LINES\n\n");

#if 1
	IDiaEnumSectionContribs * pEnumSecContribs;

	if (FAILED(GetTable(pSession, __uuidof(IDiaEnumSectionContribs), (void **)&pEnumSecContribs))) {
		return false;
	}

	IDiaSectionContrib * pSecContrib;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSecContribs->Next(1, &pSecContrib, &celt)) && (celt == 1)) {

		DWORD dwRVA;
		DWORD dwLength;

		if (SUCCEEDED(pSecContrib->get_relativeVirtualAddress(&dwRVA)) &&
			SUCCEEDED(pSecContrib->get_length(&dwLength))) {

			DumpAllLines(pSession, dwRVA, dwLength);
		}

		pSecContrib->Release();
	}

	pEnumSecContribs->Release();

	putwchar(L'\n');
#else

  // First retrieve the compilands/modules

	IDiaEnumSymbols * pEnumSymbols;

	if (FAILED(pGlobal->findChildren(SymTagCompiland, NULL, nsNone, &pEnumSymbols))) {
		return false;
	}

	IDiaSymbol * pCompiland;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pCompiland, &celt)) && (celt == 1)) {
		IDiaEnumSymbols * pEnumFunction;

		// For every function symbol defined in the compiland, retrieve and print the line numbering info

		if (SUCCEEDED(pCompiland->findChildren(SymTagFunction, NULL, nsNone, &pEnumFunction))) {
			IDiaSymbol * pFunction;
			ULONG celt = 0;

			while (SUCCEEDED(pEnumFunction->Next(1, &pFunction, &celt)) && (celt == 1)) {
				PrintLines(pSession, pFunction);

				pFunction->Release();
			}

			pEnumFunction->Release();
		}

		pCompiland->Release();
	}

	pEnumSymbols->Release();

	putwchar(L'\n');
#endif

	return true;
}

////////////////////////////////////////////////////////////
// Dump all the line numbering information for a given RVA
// and a given range
//
bool DumpAllLines(IDiaSession * pSession, DWORD dwRVA, DWORD dwRange)
{
  // Retrieve and print the lines that corresponds to a specified RVA

	IDiaEnumLineNumbers * pLines;

	if (FAILED(pSession->findLinesByRVA(dwRVA, dwRange, &pLines))) {
		return false;
	}

	PrintLines(pLines, nullptr);

	pLines->Release();

	putwchar(L'\n');

	return true;
}

struct Contribs
{
	uint32_t rva;
	uint32_t size;
	uint32_t crc;
};

std::vector<Contribs> contribs;

////////////////////////////////////////////////////////////
// Dump all the section contributions from the PDB
//
//  Section contributions are stored in a table which will
//  be retrieved via IDiaSession->getEnumTables through
//  QueryInterface()using the REFIID of the IDiaEnumSectionContribs
//
bool DumpAllSecContribs(IDiaSession * pSession)
{
	wprintf(L"\n\n*** SECTION CONTRIBUTION\n\n");

	IDiaEnumSectionContribs * pEnumSecContribs;

	if (FAILED(GetTable(pSession, __uuidof(IDiaEnumSectionContribs), (void **)&pEnumSecContribs))) {
		return false;
	}

	//wprintf(L"    RVA        Address       Size    Module\n");

	IDiaSectionContrib * pSecContrib;
	ULONG celt = 0;

	DWORD sect = 0;
	while (SUCCEEDED(pEnumSecContribs->Next(1, &pSecContrib, &celt)) && (celt == 1)) {
		pSecContrib->get_addressSection(&sect);
		//only care about text
		//if (sect != 1) {
		//  break;
		//}
		PrintSecContribs(pSession, pSecContrib);
		pSecContrib->Release();
		//putwchar(L'\n');
	}

	pEnumSecContribs->Release();

	putwchar(L'\n');

	return true;
}

////////////////////////////////////////////////////////////
// Dump all debug data streams contained in the PDB
//
bool DumpAllDebugStreams(IDiaSession * pSession)
{
	IDiaEnumDebugStreams * pEnumStreams;

	wprintf(L"\n\n*** DEBUG STREAMS\n\n");

	// Retrieve an enumerated sequence of debug data streams

	if (FAILED(pSession->getEnumDebugStreams(&pEnumStreams))) {
		return false;
	}

	IDiaEnumDebugStreamData * pStream;
	ULONG celt = 0;

	for (; SUCCEEDED(pEnumStreams->Next(1, &pStream, &celt)) && (celt == 1); pStream = NULL) {
		PrintStreamData(pStream);

		pStream->Release();
	}

	pEnumStreams->Release();

	putwchar(L'\n');

	return true;
}

////////////////////////////////////////////////////////////
// Dump all the injected source from the PDB
//
//  Injected sources data is stored in a table which will
//  be retrieved via IDiaSession->getEnumTables through
//  QueryInterface()using the REFIID of the IDiaEnumSectionContribs
//
bool DumpAllInjectedSources(IDiaSession * pSession)
{
	wprintf(L"\n\n*** INJECTED SOURCES TABLE\n\n");

	IDiaEnumInjectedSources * pEnumInjSources = NULL;

	if (FAILED(GetTable(pSession, __uuidof(IDiaEnumInjectedSources), (void **)&pEnumInjSources))) {
		return false;
	}

	IDiaInjectedSource * pInjSource;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumInjSources->Next(1, &pInjSource, &celt)) && (celt == 1)) {
		PrintGeneric(pInjSource);

		pInjSource->Release();
	}

	pEnumInjSources->Release();

	putwchar(L'\n');

	return true;
}

////////////////////////////////////////////////////////////
// Dump info corresponing to a given injected source filename
//
bool DumpInjectedSource(IDiaSession * pSession, const wchar_t * szName)
{
  // Retrieve a source that has been placed into the symbol store by attribute providers or
  //  other components of the compilation process

	IDiaEnumInjectedSources * pEnumInjSources;

	if (FAILED(pSession->findInjectedSource(szName, &pEnumInjSources))) {
		wprintf(L"ERROR - DumpInjectedSources() could not find %s\n", szName);

		return false;
	}

	IDiaInjectedSource * pInjSource;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumInjSources->Next(1, &pInjSource, &celt)) && (celt == 1)) {
		PrintGeneric(pInjSource);

		pInjSource->Release();
	}

	pEnumInjSources->Release();

	return true;
}


class CSourceFile {
	public:
	const std::wstring & GetSourceFileName();
	CV_SourceChksum_t GetCheckSumType() const;
	const char * GetCheckSum() const;

	CSourceFile(IDiaSourceFile *);
	virtual ~CSourceFile();

	private:
	IDiaSourceFile * m_pIDiaSourceFile;

	std::wstring m_fileName;

	// Checksum
	CV_SourceChksum_t m_checkSumType;
	char m_checkSum[32 + 1];
};

CSourceFile::CSourceFile(IDiaSourceFile * file) :
	m_pIDiaSourceFile(file),
	m_checkSumType(CHKSUM_TYPE_NONE),
	m_checkSum()
{
	if (m_pIDiaSourceFile == NULL) {
		return;
	}

	BSTR filename;
	if (m_pIDiaSourceFile->get_fileName(&filename) == S_OK) {
		m_fileName = filename;
		SysFreeString(filename);
	}

	memset(m_checkSum, 0, sizeof(m_checkSum));

	DWORD checksumType = CHKSUM_TYPE_NONE;
	if (m_pIDiaSourceFile->get_checksumType(&checksumType) == S_OK) {
		switch (checksumType) {
			case CHKSUM_TYPE_NONE:
				m_checkSumType = CHKSUM_TYPE_NONE;
				break;

			case CHKSUM_TYPE_MD5:
				m_checkSumType = CHKSUM_TYPE_MD5;
				break;

			case CHKSUM_TYPE_SHA1:
				m_checkSumType = CHKSUM_TYPE_SHA1;
				break;

			case CHKSUM_TYPE_SHA_256:
				m_checkSumType = CHKSUM_TYPE_SHA_256;
				break;

			default:
				m_checkSumType = CHKSUM_TYPE_NONE;
				break;
		}
	}

	char buffer[MAX_PATH] = {0};
	DWORD size = sizeof(buffer);
	if (m_checkSumType == CHKSUM_TYPE_MD5 && file->get_checksum(size, &size, (BYTE *)buffer) == S_OK) {
		char hexlut[16] = {'0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'a', 'b', 'c', 'd', 'e', 'f'};
		for (int i = 0; i < sizeof(hexlut); i++) {
			m_checkSum[i * 2] = hexlut[((buffer[i] >> 4) & 0xF)];
			m_checkSum[(i * 2) + 1] = hexlut[(buffer[i]) & 0x0F];
		}
	} else {
		strncpy_s(m_checkSum, "BAD HASH TYPE", sizeof(m_checkSum) - 1);
	}
}

CSourceFile::~CSourceFile()
{
/*
	if (m_pIDiaSourceFile) {
		m_pIDiaSourceFile->Release();
		m_pIDiaSourceFile = NULL;
	}
*/
}

const std::wstring & CSourceFile::GetSourceFileName()
{
	return m_fileName;
}

CV_SourceChksum_t CSourceFile::GetCheckSumType() const
{
	return m_checkSumType;
}

const char * CSourceFile::GetCheckSum() const
{
	return m_checkSum;
}

std::map<std::wstring, CSourceFile *> source_files_;

void CacheSourceFile(const std::wstring & file, CSourceFile * src)
{
	source_files_[file] = src;
}

bool IsSourceFileCached(const std::wstring & file)
{
	return source_files_.find(file) != source_files_.end();
}

std::vector< CSourceFile *>	m_vectorSources;

//dont need anything msvc or sdk
bool IsSkipableSourceFile(std::wstring & name)
{
	if (name.find(L"f:\\rtm\\") != std::string::npos ||
		name.find(L"f:\\sp\\public\\") != std::string::npos ||
		name.find(L"F:\\RTM\\vctools\\") != std::string::npos ||
		name.find(L"F:\\SP\\vctools\\") != std::string::npos ||
		name.find(L"f:\\sp\\vctools\\") != std::string::npos ||
		name.find(L"f:\\binaries.x86ret\\") != std::string::npos ||
		name.find(L"e:\\wm.obj.x86fre\\") != std::string::npos ||
		name.find(L"d:\\winmain\\") != std::string::npos ||
		name.find(L"d:\\winmain.public.x86fre\\") != std::string::npos ||
		name.find(L"microsoft visual studio") != std::string::npos ||
		name.find(L"\\DXSDK\\") != std::string::npos ||
		name.find(L"microsoft directx sdk") != std::string::npos) {
		return true;
	}
	return false;
}

////////////////////////////////////////////////////////////
// Dump all the source file information stored in the PDB
// We have to go through every compiland in order to retrieve
//   all the information otherwise checksums for instance
//   will not be available
// Compilands can have multiple source files with the same
//   name but different content which produces diffrent
//   checksums
//
bool DumpAllSourceFiles(IDiaSession * pSession, IDiaSymbol * pGlobal)
{
#if 0
	wprintf(L"\n\n*** SOURCE FILES\n\n");

	// To get the complete source file info we must go through the compiland first
	// by passing NULL instead all the source file names only will be retrieved

	IDiaEnumSymbols * pEnumSymbols;

	if (FAILED(pGlobal->findChildren(SymTagCompiland, NULL, nsNone, &pEnumSymbols))) {
		return false;
	}

	IDiaSymbol * pCompiland;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pCompiland, &celt)) && (celt == 1)) {
		BSTR bstrName;

		if (pCompiland->get_name(&bstrName) == S_OK) {
			wprintf(L"\nCompiland = %s\n\n", bstrName);

			SysFreeString(bstrName);
		}

		// Every compiland could contain multiple references to the source files which were used to build it
		// Retrieve all source files by compiland by passing NULL for the name of the source file

		IDiaEnumSourceFiles * pEnumSourceFiles;

		if (SUCCEEDED(pSession->findFile(pCompiland, NULL, nsNone, &pEnumSourceFiles))) {
			IDiaSourceFile * pSourceFile;

			while (SUCCEEDED(pEnumSourceFiles->Next(1, &pSourceFile, &celt)) && (celt == 1)) {
				PrintSourceFile(pSourceFile);
				putwchar(L'\n');

				pSourceFile->Release();
			}

			pEnumSourceFiles->Release();
		}

		putwchar(L'\n');

		pCompiland->Release();
	}

	pEnumSymbols->Release();
#endif

#if 0
	ULONG celt = 0;

	IDiaEnumSymbols * pEnumSymbols;
	if (FAILED(pGlobal->findChildren(SymTagCompiland, NULL, nsNone, &pEnumSymbols))) {
		return false;
	}
	IDiaEnumSourceFiles * pEnumSourceFiles;
	if (SUCCEEDED(pSession->findFile(NULL, NULL, nsNone, &pEnumSourceFiles))) {
		LONG lCount = 0;
		pEnumSourceFiles->get_Count(&lCount);
		IDiaSourceFile * pSourceFile = 0;
		for (int i = 0; i < lCount - 1; i++) {
			ULONG res;
			if (SUCCEEDED(pEnumSourceFiles->Next(1, &pSourceFile, &res)) && res == 1) {
				PrintSourceFile(pSourceFile);
				putwchar(L'\n');

				pSourceFile->Release();
			}

		}
		if (pEnumSourceFiles) {
			pEnumSourceFiles->Release();
		}


	}

#endif

#if 1
	IDiaSymbol * global;
	if (FAILED(pSession->get_globalScope(&global))) {
		fprintf(stderr, "get_globalScope failed\n");
		return false;
	}


	BSTR pdbname;
	if (SUCCEEDED(global->get_name(&pdbname))) {
		wprintf(L";PDB: %s\n", pdbname);
		SysReleaseString(pdbname);
	}

	DWORD time;
	if (SUCCEEDED(global->get_signature(&time))) {
		__time32_t now = time;
		std::tm ptm;
		_localtime32_s(&ptm, &now);
		wchar_t buffer[32];
		// Format: Mo, 15.06.2009 20:20:00
		//DebugBreak();
		wcsftime(buffer, 32, L"%d.%m.%Y %H:%M:%S", &ptm);

		wprintf(L";TimeStamp: %s (%x)\n", buffer, time);
	}

	GUID guid;
	if (SUCCEEDED(global->get_guid(&guid))) {
		const int maxGUIDStrLen = 64 + 1;
		std::vector<wchar_t> guidStr(maxGUIDStrLen);
		if (StringFromGUID2(guid, guidStr.data(), maxGUIDStrLen) > 0) {
			wprintf(L";GUID: %ls\n", guidStr.data());
		}
	}


	IDiaEnumSymbols * compilands;
	if (FAILED(global->findChildren(SymTagCompiland, NULL,
		nsNone, &compilands))) {
		fprintf(stderr, "findChildren failed\n");
		return false;
	}

	LONG c;
	;

	if (FAILED(compilands->get_Count(&c))) {
		fprintf(stderr, "compilands->get_Count failed\n");
		return false;
	}
	//ULONG count;

	IDiaEnumSourceFiles * source_files;
	IDiaSymbol * compiland = NULL;

	wprintf(L";Compilands: %d\n", c);
	for (int cidx = 0; cidx < c; cidx++) {
	//while (SUCCEEDED(compilands->Next(i, &compiland, &count))) {
	  //if (count > 1) {
	  //    fprintf(stderr, "Next failed %d\n", count);
	  //    return false;
	  //}

		if (FAILED(compilands->Item(cidx, &compiland))) {
			fprintf(stderr, "compilands->Item(%d) failed\n", cidx);
			return false;
		}

		if (FAILED(pSession->findFile(compiland, NULL, nsNone, &source_files))) {
			fprintf(stderr, "pSession->findFile(%d) failed\n", cidx);
			return false;
		}

		LONG sc;
		if (FAILED(source_files->get_Count(&sc))) {
			fprintf(stderr, "source_files->get_Count failed\n");
			return false;
		}
		//fprintf(stderr, "source_files %d\n", sc);

		BSTR cname;
		if (FAILED(compiland->get_name(&cname))) {
			fprintf(stderr, "compiland->get_name failed\n");
			return false;
		}
		//wprintf(L";Compiland \"%s\" source_files %d\n", cname, sc);
		if (sc == 0) {
			wprintf(L";WARNING No Symbols in Compiland \"%s\"\n", cname);
		}
		SysFreeString(cname);

		IDiaSourceFile * source_file;
		for (int sidx = 0; sidx < sc; sidx++) {
			if (FAILED(source_files->Item(sidx, &source_file))) {
				fprintf(stderr, "source_files->Item(%d) failed\n", sidx);
				return false;
			}

			BSTR file_name;
			if (FAILED(source_file->get_fileName(&file_name))) {
				fprintf(stderr, "source_file->get_fileName(%d) failed\n", sidx);
				return false;
			}

			std::wstring file_name_string(file_name);
			if (!IsSourceFileCached(file_name_string)) {
				// this is a new file name, cache it and output a FILE line.
				CSourceFile * pCSourceFile = new CSourceFile(source_file);
				CacheSourceFile(file_name_string, pCSourceFile);
			}
			SysFreeString(file_name);
			source_file->Release();
		}
		compiland->Release();
	}

	for (auto it = source_files_.begin(); it != source_files_.end(); it++) {
		CSourceFile * src = it->second;
		if (src != nullptr) {
			std::wstring name = src->GetSourceFileName();
			auto checksumType = src->GetCheckSumType();
			auto checksum = src->GetCheckSum();
			bool skip = false;
#if 1
	  // skip msvc
			if (IsSkipableSourceFile(name)) {
				skip = true;
			}
#endif
			if (!name.empty()) {
#if 0
				if (checksumType == CHKSUM_TYPE_MD5) {
					//wprintf(L"%hs *%ls\n", checksum, name.c_str());
					wprintf(L"%s*%ls %hs\n", skip ? L";" : L"", name.c_str(), checksum);
				} else {
					wprintf(L"%s*%ls %hs\n", skip ? L";" : L"", name.c_str(), "ERRRRRRRRRRRRRRRRRRRRR");
				}
#endif				
				
				if (checksumType != CHKSUM_TYPE_MD5) {
					checksum = "ERRRRRRRRRRRRRRRRRRRRR";
					fprintf(stderr, "Unhandled checksum(%d)\n", checksumType);
					skip = true;
				}
				//wprintf(L"%s*%ls %hs\n", skip ? L";" : L"", name.c_str(), checksum);
				wprintf(L"%s %hs *%ls\n", skip ? L";" : L"", checksum, name.c_str());
			}
		}
	}
#endif
#if 0
	wprintf(L"\n\n*** FILES\n\n");

	IDiaEnumSourceFiles * pEnumSourceFiles;

	if (FAILED(GetTable(g_pDiaSession, __uuidof(IDiaEnumSourceFiles), (void **)&pEnumSourceFiles))) {
		return false;
	}

	IDiaSourceFile * pSourceFile;
	ULONG celt = 0;

	IDiaEnumSymbols * compilands;
	if (FAILED(pGlobal->findChildren(SymTagCompiland, NULL,
		nsNone, &compilands))) {
		fprintf(stderr, "findChildren failed\n");
		return false;
	}

	while (SUCCEEDED(pEnumSourceFiles->Next(1, &pSourceFile, &celt)) && (celt == 1)) {

		BSTR file_name;
		if (FAILED(pSourceFile->get_fileName(&file_name))) {
			return false;
		}

		std::wstring file_name_string(file_name);
		if (!IsSourceFileCached(file_name_string)) {
		  // this is a new file name, cache it and output a FILE line.
			CSourceFile * pCSourceFile = new CSourceFile(pSourceFile);
			CacheSourceFile(file_name_string, pCSourceFile);
		}
		SysFreeString(file_name);

		pSourceFile->Release();
	}

	pEnumSourceFiles->Release();

	for (auto it = source_files_.begin(); it != source_files_.end(); it++) {
		CSourceFile * src = it->second;
		if (src != nullptr) {
			std::wstring name = src->GetSourceFileName();
			auto checksumType = src->GetCheckSumType();
			auto checksum = src->GetCheckSum();
			// skip msvc
			if (IsSkipableSourceFile(name)) {
				continue;
			}
			if (!name.empty() && checksumType == CheckSumType_MD5) {
				wprintf(L"%hs *%ls\n", checksum, name.c_str());
			}
		}
	}
#endif

#if 0
	IDiaSymbol * global;
	if (FAILED(pSession->get_globalScope(&global))) {
		fprintf(stderr, "get_globalScope failed\n");
		return false;
	}

	IDiaEnumSourceFiles * pEnumSourceFiles;

	if (FAILED(GetTable(g_pDiaSession, __uuidof(IDiaEnumSourceFiles), (void **)&pEnumSourceFiles))) {
		return false;
	}

	IDiaSession * pIDiaSession = pSession;
	if (pIDiaSession) {

		if (1) {
			LONG lCount = 0;
			pEnumSourceFiles->get_Count(&lCount);
			if (lCount == 0) {
				DebugBreak();
			}

			for (DWORD dw = 0; dw < (DWORD)lCount; dw++) {
				IDiaSourceFile * pIDiaSourceFile = NULL;
				if (pEnumSourceFiles->Item(dw, &pIDiaSourceFile) == S_OK) {
					BSTR file_name;
					if (FAILED(pIDiaSourceFile->get_fileName(&file_name))) {
						DebugBreak();
						return false;
					}

					std::wstring file_name_string(file_name);
					if (!IsSourceFileCached(file_name_string)) {
						//DebugBreak();
					  // this is a new file name, cache it and output a FILE line.
						CSourceFile * pCSourceFile = new CSourceFile(pIDiaSourceFile);
						CacheSourceFile(file_name_string, pCSourceFile);
					}
					SysFreeString(file_name);

				} else {
					DebugBreak();
				}
			}
		}
		if (pEnumSourceFiles) {
			pEnumSourceFiles->Release();
		}
	}


	for (auto it = source_files_.begin(); it != source_files_.end(); it++) {
		CSourceFile * src = it->second;
		if (src != nullptr) {
			std::wstring name = src->GetSourceFileName();
			auto checksumType = src->GetCheckSumType();
			auto checksum = src->GetCheckSum();
			// skip msvc
			if (IsSkipableSourceFile(name)) {
				continue;
			}
			if (!name.empty() && checksumType == CheckSumType_MD5) {
				wprintf(L"%hs *%ls\n", checksum, name.c_str());
			} else {
				wprintf(L"%hs *%ls\n", "0", name.c_str());

			}
		}
	}
#endif
	return true;
}

////////////////////////////////////////////////////////////
// Dump all the FPO info
//
//  FPO data stored in a table which will be retrieved via
//    IDiaSession->getEnumTables through QueryInterface()
//    using the REFIID of the IDiaEnumFrameData
//
bool DumpAllFPO(IDiaSession * pSession)
{
	IDiaEnumFrameData * pEnumFrameData;

	wprintf(L"\n\n*** FPO\n\n");

	if (FAILED(GetTable(pSession, __uuidof(IDiaEnumFrameData), (void **)&pEnumFrameData))) {
		return false;
	}

	IDiaFrameData * pFrameData;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumFrameData->Next(1, &pFrameData, &celt)) && (celt == 1)) {
		PrintFrameData(pFrameData);

		pFrameData->Release();
	}

	pEnumFrameData->Release();

	putwchar(L'\n');

	return true;
}

////////////////////////////////////////////////////////////
// Dump FPO info for a function at the specified RVA
//
bool DumpFPO(IDiaSession * pSession, DWORD dwRVA)
{
	IDiaEnumFrameData * pEnumFrameData;

	// Retrieve first the table holding all the FPO info

	if ((dwRVA != 0) && SUCCEEDED(GetTable(pSession, __uuidof(IDiaEnumFrameData), (void **)&pEnumFrameData))) {
		IDiaFrameData * pFrameData;

		// Retrieve the frame data corresponding to the given RVA

		if (SUCCEEDED(pEnumFrameData->frameByRVA(dwRVA, &pFrameData))) {
			PrintGeneric(pFrameData);

			pFrameData->Release();
		}

		else {
		  // Some function might not have FPO data available (see ASM funcs like strcpy)

			wprintf(L"ERROR - DumpFPO() frameByRVA invalid RVA: 0x%08X\n", dwRVA);

			pEnumFrameData->Release();

			return false;
		}

		pEnumFrameData->Release();
	}

	else {
		wprintf(L"ERROR - DumpFPO() GetTable\n");

		return false;
	}

	putwchar(L'\n');

	return true;
}

////////////////////////////////////////////////////////////
// Dump FPO info for a specified function symbol using its
//  name (a regular expression string is used for the search)
//
bool DumpFPO(IDiaSession * pSession, IDiaSymbol * pGlobal, const wchar_t * szSymbolName)
{
	IDiaEnumSymbols * pEnumSymbols;
	IDiaSymbol * pSymbol;
	ULONG celt = 0;
	DWORD dwRVA;

	// Find first all the function symbols that their names matches the search criteria

	if (FAILED(pGlobal->findChildren(SymTagFunction, szSymbolName, nsRegularExpression, &pEnumSymbols))) {
		wprintf(L"ERROR - DumpFPO() findChildren could not find symol %s\n", szSymbolName);

		return false;
	}

	while (SUCCEEDED(pEnumSymbols->Next(1, &pSymbol, &celt)) && (celt == 1)) {
		if (pSymbol->get_relativeVirtualAddress(&dwRVA) == S_OK) {
			PrintPublicSymbol(pSymbol);

			DumpFPO(pSession, dwRVA);
		}

		pSymbol->Release();
	}

	pEnumSymbols->Release();

	putwchar(L'\n');

	return true;
}

////////////////////////////////////////////////////////////
// Dump a specified compiland and all the symbols defined in it
//
bool DumpCompiland(IDiaSymbol * pGlobal, const wchar_t * szCompName)
{
	IDiaEnumSymbols * pEnumSymbols;

	// was nsCaseInsensitive
	if (FAILED(pGlobal->findChildren(SymTagCompiland, szCompName, nsCaseInRegularExpression, &pEnumSymbols))) {
		return false;
	}

	IDiaSymbol * pCompiland;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pCompiland, &celt)) && (celt == 1)) {
		wprintf(L"\n** Module: ");

		// Retrieve the name of the module

		BSTR bstrName;

		if (pCompiland->get_name(&bstrName) != S_OK) {
			wprintf(L"(???)\n\n");
		}

		else {
			wprintf(L"%s\n\n", bstrName);

			SysFreeString(bstrName);
		}

		IDiaEnumSymbols * pEnumChildren;

		if (SUCCEEDED(pCompiland->findChildren(SymTagNull, NULL, nsNone, &pEnumChildren))) {
			IDiaSymbol * pSymbol;
			ULONG celt_ = 0;

			while (SUCCEEDED(pEnumChildren->Next(1, &pSymbol, &celt_)) && (celt_ == 1)) {
				PrintSymbol(pSymbol, 0);

				pSymbol->Release();
			}

			pEnumChildren->Release();
		}

		pCompiland->Release();
	}
#if 0
	if (celt)
	{
		DebugBreak();
	}
#endif
	pEnumSymbols->Release();

	return true;
}

////////////////////////////////////////////////////////////
// Dump the line numbering information for a specified RVA
//
bool DumpLines(IDiaSession * pSession, DWORD dwRVA)
{
	IDiaEnumLineNumbers * pLines;

	if (FAILED(pSession->findLinesByRVA(dwRVA, MAX_RVA_LINES_BYTES_RANGE, &pLines))) {
		return false;
	}

	PrintLines(pLines, nullptr);

	pLines->Release();

	return true;
}

////////////////////////////////////////////////////////////
// Dump the all line numbering information for a specified
//  function symbol name (as a regular expression string)
//
bool DumpLines(IDiaSession * pSession, IDiaSymbol * pGlobal, const wchar_t * szFuncName)
{
	IDiaEnumSymbols * pEnumSymbols;

	if (FAILED(pGlobal->findChildren(SymTagFunction, szFuncName, nsRegularExpression, &pEnumSymbols))) {
		return false;
	}

	IDiaSymbol * pFunction;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pFunction, &celt)) && (celt == 1)) {
		PrintLines(pSession, pFunction);

		pFunction->Release();
	}

	pEnumSymbols->Release();

	return true;
}

////////////////////////////////////////////////////////////
// Dump the symbol information corresponding to a specified RVA
//
bool DumpSymbolWithRVA(IDiaSession * pSession, DWORD dwRVA, const wchar_t * szChildname)
{
	IDiaSymbol * pSymbol;
	LONG lDisplacement;

	if (FAILED(pSession->findSymbolByRVAEx(dwRVA, SymTagNull, &pSymbol, &lDisplacement))) {
		return false;
	}

	wprintf(L"Displacement = 0x%X\n", lDisplacement);

	PrintGeneric(pSymbol);

	DumpSymbolWithChildren(pSymbol, szChildname);

	while (pSymbol != NULL) {
		IDiaSymbol * pParent;

		if ((pSymbol->get_lexicalParent(&pParent) == S_OK) && pParent) {
			wprintf(L"\nParent\n");

			PrintSymbol(pParent, 0);

			pSymbol->Release();

			pSymbol = pParent;
		}

		else {
			pSymbol->Release();
			break;
		}
	}

	return true;
}

////////////////////////////////////////////////////////////
// Dump the symbols information where their names matches a
//  specified regular expression string
//
bool DumpSymbolsWithRegEx(IDiaSymbol * pGlobal, const wchar_t * szRegEx, const wchar_t * szChildname)
{
	IDiaEnumSymbols * pEnumSymbols;

	if (FAILED(pGlobal->findChildren(SymTagNull, szRegEx, nsRegularExpression, &pEnumSymbols))) {
		return false;
	}

	bool bReturn = true;

	IDiaSymbol * pSymbol;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pSymbol, &celt)) && (celt == 1)) {
		PrintGeneric(pSymbol);

		bReturn = DumpSymbolWithChildren(pSymbol, szChildname);

		pSymbol->Release();
	}

	pEnumSymbols->Release();

	return bReturn;
}

////////////////////////////////////////////////////////////
// Dump the information corresponding to a symbol name which
//  is a children of the specified parent symbol
//
bool DumpSymbolWithChildren(IDiaSymbol * pSymbol, const wchar_t * szChildname)
{
	if (szChildname != NULL) {
		IDiaEnumSymbols * pEnumSyms;

		if (FAILED(pSymbol->findChildren(SymTagNull, szChildname, nsRegularExpression, &pEnumSyms))) {
			return false;
		}

		IDiaSymbol * pChild;
		DWORD celt = 1;

		while (SUCCEEDED(pEnumSyms->Next(celt, &pChild, &celt)) && (celt == 1)) {
			PrintGeneric(pChild);
			PrintSymbol(pChild, 0);

			pChild->Release();
		}

		pEnumSyms->Release();
	}

	else {
	  // If the specified name is NULL then only the parent symbol data is displayed

		DWORD dwSymTag;

		if ((pSymbol->get_symTag(&dwSymTag) == S_OK) && (dwSymTag == SymTagPublicSymbol)) {
			PrintPublicSymbol(pSymbol);
		}

		else {
			PrintSymbol(pSymbol, 0);
		}
	}

	return true;
}

////////////////////////////////////////////////////////////
// Dump all the type symbols information that matches their
//  names to a specified regular expression string
//
bool DumpType(IDiaSymbol * pGlobal, const wchar_t * szRegEx)
{
	IDiaEnumSymbols * pEnumSymbols;

	if (FAILED(pGlobal->findChildren(SymTagUDT, szRegEx, nsRegularExpression, &pEnumSymbols))) {
		return false;
	}

	IDiaSymbol * pSymbol;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pSymbol, &celt)) && (celt == 1)) {
		PrintTypeInDetail(pSymbol, 0);

		pSymbol->Release();
	}

	pEnumSymbols->Release();

	return true;
}

////////////////////////////////////////////////////////////
// Dump line numbering information for a given file name and
//  an optional line number
//
bool DumpLinesForSourceFile(IDiaSession * pSession, const wchar_t * szFileName, DWORD dwLine)
{
	IDiaEnumSourceFiles * pEnumSrcFiles;

	if (FAILED(pSession->findFile(NULL, szFileName, nsFNameExt, &pEnumSrcFiles))) {
		return false;
	}

	IDiaSourceFile * pSrcFile;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSrcFiles->Next(1, &pSrcFile, &celt)) && (celt == 1)) {
		IDiaEnumSymbols * pEnumCompilands;

		if (pSrcFile->get_compilands(&pEnumCompilands) == S_OK) {
			IDiaSymbol * pCompiland;

			celt = 0;
			while (SUCCEEDED(pEnumCompilands->Next(1, &pCompiland, &celt)) && (celt == 1)) {
				BSTR bstrName;

				if (pCompiland->get_name(&bstrName) == S_OK) {
					wprintf(L"//Compiland = %s\n", bstrName);

					SysFreeString(bstrName);
				}

				else {
					wprintf(L"//Compiland = (???)\n");
				}

				IDiaEnumLineNumbers * pLines;

				if (dwLine != 0) {
					if (SUCCEEDED(pSession->findLinesByLinenum(pCompiland, pSrcFile, dwLine, 0, &pLines))) {
						PrintLines(pLines, szFileName);

						pLines->Release();
					}
				}

				else {
					if (SUCCEEDED(pSession->findLines(pCompiland, pSrcFile, &pLines))) {
						PrintLines(pLines, szFileName);

						pLines->Release();
					}
				}

				pCompiland->Release();
			}

			pEnumCompilands->Release();
		}

		pSrcFile->Release();
	}

	pEnumSrcFiles->Release();

	return true;
}

////////////////////////////////////////////////////////////
// Dump public symbol information for a given number of
//  symbols around a given RVA address
//
bool DumpPublicSymbolsSorted(IDiaSession * pSession, DWORD dwRVA, DWORD dwRange, bool bReverse)
{
	IDiaEnumSymbolsByAddr * pEnumSymsByAddr;

	if (FAILED(pSession->getSymbolsByAddr(&pEnumSymsByAddr))) {
		return false;
	}

	IDiaSymbol * pSymbol;

	if (SUCCEEDED(pEnumSymsByAddr->symbolByRVA(dwRVA, &pSymbol))) {
		if (dwRange == 0) {
			PrintPublicSymbol(pSymbol);
		}

		ULONG celt;
		ULONG i;

		if (bReverse) {
			pSymbol->Release();

			i = 0;

			for (pSymbol = NULL; (i < dwRange) && SUCCEEDED(pEnumSymsByAddr->Next(1, &pSymbol, &celt)) && (celt == 1); i++) {
				PrintPublicSymbol(pSymbol);

				pSymbol->Release();
			}
		}

		else {
			PrintPublicSymbol(pSymbol);

			pSymbol->Release();

			i = 1;

			for (pSymbol = NULL; (i < dwRange) && SUCCEEDED(pEnumSymsByAddr->Prev(1, &pSymbol, &celt)) && (celt == 1); i++) {
				PrintPublicSymbol(pSymbol);
			}
		}
	}

	pEnumSymsByAddr->Release();

	return true;
}

////////////////////////////////////////////////////////////
// Dump label symbol information at a given RVA
//
bool DumpLabel(IDiaSession * pSession, DWORD dwRVA)
{
	IDiaSymbol * pSymbol;
	LONG lDisplacement;

	if (FAILED(pSession->findSymbolByRVAEx(dwRVA, SymTagLabel, &pSymbol, &lDisplacement)) || (pSymbol == NULL)) {
		return false;
	}

	wprintf(L"Displacement = 0x%X\n", lDisplacement);

	PrintGeneric(pSymbol);

	pSymbol->Release();

	return true;
}

////////////////////////////////////////////////////////////
// Dump annotation symbol information at a given RVA
//
bool DumpAnnotations(IDiaSession * pSession, DWORD dwRVA)
{
	IDiaSymbol * pSymbol;
	LONG lDisplacement;

	if (FAILED(pSession->findSymbolByRVAEx(dwRVA, SymTagAnnotation, &pSymbol, &lDisplacement)) || (pSymbol == NULL)) {
		return false;
	}

	wprintf(L"Displacement = 0x%X\n", lDisplacement);

	PrintGeneric(pSymbol);

	pSymbol->Release();

	return true;
}

struct OMAP_DATA
{
	DWORD dwRVA;
	DWORD dwRVATo;
};

////////////////////////////////////////////////////////////
//
bool DumpMapToSrc(IDiaSession * pSession, DWORD dwRVA)
{
	IDiaEnumDebugStreams * pEnumStreams;
	IDiaEnumDebugStreamData * pStream;
	ULONG celt;

	if (FAILED(pSession->getEnumDebugStreams(&pEnumStreams))) {
		return false;
	}

	celt = 0;

	for (; SUCCEEDED(pEnumStreams->Next(1, &pStream, &celt)) && (celt == 1); pStream = NULL) {
		BSTR bstrName;

		if (pStream->get_name(&bstrName) != S_OK) {
			bstrName = NULL;
		}

		if (bstrName && wcscmp(bstrName, L"OMAPTO") == 0) {
			OMAP_DATA data, datasav;
			DWORD cbData;
			DWORD dwRVATo = 0;
			unsigned int i;

			datasav.dwRVATo = 0;
			datasav.dwRVA = 0;

			while (SUCCEEDED(pStream->Next(1, sizeof(data), &cbData, (BYTE *)&data, &celt)) && (celt == 1)) {
				if (dwRVA > data.dwRVA) {
					datasav = data;
					continue;
				}

				else if (dwRVA == data.dwRVA) {
					dwRVATo = data.dwRVATo;
				}

				else if (datasav.dwRVATo) {
					dwRVATo = datasav.dwRVATo + (dwRVA - datasav.dwRVA);
				}
				break;
			}

			wprintf(L"image rva = %08X ==> source rva = %08X\n\nRelated OMAP entries:\n", dwRVA, dwRVATo);
			wprintf(L"image rva ==> source rva\n");
			wprintf(L"%08X  ==> %08X\n", datasav.dwRVA, datasav.dwRVATo);

			i = 0;

			do {
				wprintf(L"%08X  ==> %08X\n", data.dwRVA, data.dwRVATo);
			} while ((++i) < 5 && SUCCEEDED(pStream->Next(1, sizeof(data), &cbData, (BYTE *)&data, &celt)) && (celt == 1));
		}

		if (bstrName != NULL) {
			SysFreeString(bstrName);
		}

		pStream->Release();
	}

	pEnumStreams->Release();

	return true;
}

////////////////////////////////////////////////////////////
//
bool DumpMapFromSrc(IDiaSession * pSession, DWORD dwRVA)
{
	IDiaEnumDebugStreams * pEnumStreams;

	if (FAILED(pSession->getEnumDebugStreams(&pEnumStreams))) {
		return false;
	}

	IDiaEnumDebugStreamData * pStream;
	ULONG celt = 0;

	for (; SUCCEEDED(pEnumStreams->Next(1, &pStream, &celt)) && (celt == 1); pStream = NULL) {
		BSTR bstrName;

		if (pStream->get_name(&bstrName) != S_OK) {
			bstrName = NULL;
		}

		if (bstrName && wcscmp(bstrName, L"OMAPFROM") == 0) {
			OMAP_DATA data;
			OMAP_DATA datasav;
			DWORD cbData;
			DWORD dwRVATo = 0;
			unsigned int i;

			datasav.dwRVATo = 0;
			datasav.dwRVA = 0;

			while (SUCCEEDED(pStream->Next(1, sizeof(data), &cbData, (BYTE *)&data, &celt)) && (celt == 1)) {
				if (dwRVA > data.dwRVA) {
					datasav = data;
					continue;
				}

				else if (dwRVA == data.dwRVA) {
					dwRVATo = data.dwRVATo;
				}

				else if (datasav.dwRVATo) {
					dwRVATo = datasav.dwRVATo + (dwRVA - datasav.dwRVA);
				}
				break;
			}

			wprintf(L"source rva = %08X ==> image rva = %08X\n\nRelated OMAP entries:\n", dwRVA, dwRVATo);
			wprintf(L"source rva ==> image rva\n");
			wprintf(L"%08X  ==> %08X\n", datasav.dwRVA, datasav.dwRVATo);

			i = 0;

			do {
				wprintf(L"%08X  ==> %08X\n", data.dwRVA, data.dwRVATo);
			} while ((++i) < 5 && SUCCEEDED(pStream->Next(1, sizeof(data), &cbData, (BYTE *)&data, &celt)) && (celt == 1));
		}

		if (bstrName != NULL) {
			SysFreeString(bstrName);
		}

		pStream->Release();
	}

	pEnumStreams->Release();

	return true;
}

////////////////////////////////////////////////////////////
// Retreive the table that matches the given iid
//
//  A PDB table could store the section contributions, the frame data,
//  the injected sources
//
HRESULT GetTable(IDiaSession * pSession, REFIID iid, void ** ppUnk)
{
	IDiaEnumTables * pEnumTables;

	if (FAILED(pSession->getEnumTables(&pEnumTables))) {
		wprintf(L"ERROR - GetTable() getEnumTables\n");

		return E_FAIL;
	}

	IDiaTable * pTable;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumTables->Next(1, &pTable, &celt)) && (celt == 1)) {
	  // There's only one table that matches the given IID

		if (SUCCEEDED(pTable->QueryInterface(iid, (void **)ppUnk))) {
			pTable->Release();
			pEnumTables->Release();

			return S_OK;
		}

		pTable->Release();
	}

	pEnumTables->Release();

	return E_FAIL;
}

//PDBTypeMatch code
//from https://github.com/dotnet/coreclr/tree/release/1.1.0/src/ToolBox/PdbTypeMatch




#define DEBUG_VERBOSE 0

#pragma warning (disable : 4100)

const wchar_t * g_szFilename1, * g_szFilename2;
IDiaDataSource * g_pDiaDataSource1, * g_pDiaDataSource2;
IDiaSession * g_pDiaSession1, * g_pDiaSession2;
IDiaSymbol * g_pGlobalSymbol1, * g_pGlobalSymbol2;

typedef std::set<std::wstring> IDiaSymbolSet;
IDiaSymbolSet g_excludedTypes;
IDiaSymbolSet g_excludedTypePatterns;

bool InitDiaSource(IDiaDataSource ** ppSource);
void Cleanup2();
LPSTR UnicodeToAnsi(LPCWSTR s);
bool EnumTypesInPdb(IDiaSymbolSet * types, IDiaSession * pSession, IDiaSymbol * pGlobal);
bool LayoutMatches(IDiaSymbol * pSymbol1, IDiaSymbol * pSymbol2);
void PrintHelpOptions2();

////////////////////////////////////////////////////////////
//
int __cdecl wmain2(int argc, wchar_t * argv[])
{
	FILE * pFile;

	if (argc < 3) {
		PrintHelpOptions2();
		return -1;
	}

	if (!_wcsicmp(argv[1], L"-ptype")) {

	  // -type <symbolname> <pdbfilename>: dump this type in detail

		if (argc < 4) {
			PrintHelpOptions2();
			return -1;
		}

		if ((argc > 1) && (*argv[2] != L'-')) {

			if (_wfopen_s(&pFile, argv[3], L"r") || !pFile) {
			// invalid file name or file does not exist

				PrintHelpOptions2();
				return -1;
			}
			fclose(pFile);
			// CoCreate() and initialize COM objects
			if (!InitDiaSource(&g_pDiaDataSource1)) {
				return -1;
			}
			if (!LoadDataFromPdb(argv[3], &g_pDiaDataSource1, &g_pDiaSession1, &g_pGlobalSymbol1)) {
				return -1;
			}

			DumpType(g_pGlobalSymbol1, argv[2]);

			// release COM objects and CoUninitialize()
			Cleanup2();

			return 0;
		}
	}

	if (argc < 3) {
		PrintHelpOptions2();
		return -1;
	}

	if (_wfopen_s(&pFile, argv[1], L"r") || !pFile) {
	  // invalid file name or file does not exist

		PrintHelpOptions2();
		return -1;
	}

	fclose(pFile);

	if (_wfopen_s(&pFile, argv[2], L"r") || !pFile) {
	  // invalid file name or file does not exist

		PrintHelpOptions2();
		return -1;
	}

	fclose(pFile);

	g_szFilename1 = argv[1];

	// CoCreate() and initialize COM objects
	if (!InitDiaSource(&g_pDiaDataSource1)) {
		return -1;
	}
	if (!LoadDataFromPdb(g_szFilename1, &g_pDiaDataSource1, &g_pDiaSession1, &g_pGlobalSymbol1)) {
		return -1;
	}

	g_szFilename2 = argv[2];

	InitDiaSource(&g_pDiaDataSource2);
	if (!LoadDataFromPdb(g_szFilename2, &g_pDiaDataSource2, &g_pDiaSession2, &g_pGlobalSymbol2)) {
		return -1;
	}

	if (argv[3]) {
	  // Read exclusion list.
		struct stat fileStatus;
		if (stat(UnicodeToAnsi(argv[3]), &fileStatus) != 0) {
			wprintf(L"Could not open type_exclusion_list file!\n");
			return -1;
		}

		char linec[2048];
		FILE * file;
		fopen_s(&file, UnicodeToAnsi(argv[3]), "r");
		while (fgets(linec, sizeof(linec), file) != NULL) {
			std::string line(linec);
			line.erase(std::remove_if(line.begin(), line.end(), isspace), line.end());
			if (line.empty() || line.length() <= 1) {
				continue;
			}
			if (line.front() == '#') continue;
			int len;
			int slength = (int)line.length() + 1;
			len = MultiByteToWideChar(CP_ACP, 0, line.c_str(), slength, 0, 0);
			wchar_t * buf = new wchar_t[len];
			MultiByteToWideChar(CP_ACP, 0, line.c_str(), slength, buf, len);
			std::wstring wLine(buf);
			delete[] buf;

			/// Add *str in the patterns list.
			if (line.front() == '*') {
				g_excludedTypePatterns.insert((std::wstring)(wLine.substr(1, wLine.size() - 1)));
			} else {
				g_excludedTypes.insert((std::wstring)(wLine));
			}
		}
		fclose(file);
	}

	IDiaSymbolSet types1;
	IDiaSymbolSet types2;
	if (!EnumTypesInPdb(&types1, g_pDiaSession1, g_pGlobalSymbol1)) {
		return -1;
	}

	if (!EnumTypesInPdb(&types2, g_pDiaSession2, g_pGlobalSymbol2)) {
		return -1;
	}

	IDiaSymbolSet commonTypes;

	// Intersect types
	for (IDiaSymbolSet::iterator i = types1.begin(); i != types1.end(); i++) {
		std::wstring typeName = *i;

		/// Skip excluded types
		if (g_excludedTypes.find(typeName) != g_excludedTypes.end()) {
			continue;
		}
		bool skipType = false;
		/// Skip if includes one pattern string.
		for (IDiaSymbolSet::iterator j = g_excludedTypePatterns.begin(); j != g_excludedTypePatterns.end(); j++) {
			std::wstring patternStr = *j;
			if (wcsstr(typeName.c_str(), patternStr.c_str()) != NULL) {
				skipType = true;
				break;
			}
		}

		if (skipType) continue;

		if (types2.find(typeName) != types2.end()) {
			commonTypes.insert(typeName);
		}
	}

	bool matchedSymbols = true;
	ULONG failuresNb = 0;


	// Compare layout for common types
	for (IDiaSymbolSet::iterator i = commonTypes.begin(); i != commonTypes.end(); i++) {
		std::wstring typeName = *i;

		IDiaEnumSymbols * pEnumSymbols1;
		BSTR pwstrTypeName = SysAllocString(typeName.c_str());
		if (FAILED(g_pGlobalSymbol1->findChildren(SymTagUDT, pwstrTypeName, nsNone, &pEnumSymbols1))) {
			SysFreeString(pwstrTypeName);
			return false;
		}

		IDiaEnumSymbols * pEnumSymbols2;
		if (FAILED(g_pGlobalSymbol2->findChildren(SymTagUDT, pwstrTypeName, nsNone, &pEnumSymbols2))) {
			SysFreeString(pwstrTypeName);
			return false;
		}

		IDiaSymbol * pSymbol1;
		IDiaSymbol * pSymbol2;
		ULONG celt = 0;


		while (SUCCEEDED(pEnumSymbols1->Next(1, &pSymbol1, &celt)) && (celt == 1)) {
			if (SUCCEEDED(pEnumSymbols2->Next(1, &pSymbol2, &celt)) && (celt == 1)) {

				BSTR bstrSymbol1Name;
				if (pSymbol1->get_name(&bstrSymbol1Name) != S_OK) {
					bstrSymbol1Name = NULL;
					pSymbol2->Release();
					continue;
				}
				BSTR bstrSymbol2Name;
				if (pSymbol2->get_name(&bstrSymbol2Name) != S_OK) {
					bstrSymbol2Name = NULL;
					pSymbol2->Release();
					continue;
				}
				if (_wcsicmp(bstrSymbol1Name, bstrSymbol2Name) != 0) {
					pSymbol2->Release();
					continue;
				}
				ULONGLONG sym1Size;
				ULONGLONG sym2Size;
				if (pSymbol1->get_length(&sym1Size) != S_OK) {
					wprintf(L"ERROR - can't retrieve the symbol's length\n");
					pSymbol2->Release();
					continue;
				}
				//wprintf(L"sym1Size = %x\n", sym1Size);
				if (pSymbol2->get_length(&sym2Size) != S_OK) {
					wprintf(L"ERROR - can't retrieve the symbol's length\n");
					pSymbol2->Release();
					continue;
				}
				//wprintf(L"sym2Size = %x\n", sym2Size);
				if (sym1Size == 0 || sym2Size == 2) {
					pSymbol2->Release();
					continue;
				}

				if (!LayoutMatches(pSymbol1, pSymbol2)) {
					wprintf(L"Type \"%s\" is not matching in %s and %s\n", pwstrTypeName, g_szFilename1, g_szFilename2);
					pSymbol2->Release();

					matchedSymbols = false;
					failuresNb++;
					// Continue to compare and report all inconsistencies.
					continue;
				} else {
#if	DEBUG_VERBOSE
					wprintf(L"Matched type: %s\n", pwstrTypeName);
#endif
				}

				pSymbol2->Release();
			}

			pSymbol1->Release();
		}

		SysFreeString(pwstrTypeName);
		pEnumSymbols1->Release();
		pEnumSymbols2->Release();
	}

	// release COM objects and CoUninitialize()
	Cleanup2();

	if (matchedSymbols) {
		wprintf(L"OK: All %d common types of %s and %s match!\n", commonTypes.size(), g_szFilename1, g_szFilename2);
		return 0;
	} else {
		wprintf(L"FAIL: Failed to match %d common types of %s and %s!\n", failuresNb, g_szFilename1, g_szFilename2);
		wprintf(L"Matched %d common types!\n", commonTypes.size() - failuresNb);
		return -1;
	}
}

LPSTR UnicodeToAnsi(LPCWSTR s)
{
	if (s == NULL) return NULL;
	int cw = lstrlenW(s);
	if (cw == 0) {
		CHAR * psz = new CHAR[1]; *psz = '\0'; return psz;
	}
	int cc = WideCharToMultiByte(CP_ACP, 0, s, cw, NULL, 0, NULL, NULL);
	if (cc == 0) return NULL;
	CHAR * psz = new CHAR[cc + 1];
	cc = WideCharToMultiByte(CP_ACP, 0, s, cw, psz, cc, NULL, NULL);
	if (cc == 0) { delete[] psz; return NULL; }
	psz[cc] = '\0';
	return psz;
}


bool InitDiaSource(IDiaDataSource ** ppSource)
{
	HRESULT hr = CoInitialize(NULL);

	// Obtain access to the provider

	hr = CoCreateInstance(__uuidof(DiaSource),
						NULL,
						CLSCTX_INPROC_SERVER,
						__uuidof(IDiaDataSource),
						(void **)ppSource);

	if (FAILED(hr)) {
		ACTCTX actCtx;
		memset((void *)&actCtx, 0, sizeof(ACTCTX));
		actCtx.cbSize = sizeof(ACTCTX);
		WCHAR   dllPath[MAX_PATH * 2];
		GetModuleFileName(NULL, dllPath, _countof(dllPath));
		PathRemoveFileSpec(dllPath);
		wcscat_s(dllPath, L"\\msdia100.sxs.manifest");
		actCtx.lpSource = dllPath;

		HANDLE hCtx = ::CreateActCtx(&actCtx);
		if (hCtx == INVALID_HANDLE_VALUE)
			wprintf(L"CreateActCtx returned: INVALID_HANDLE_VALUE\n");
		else {
			ULONG_PTR cookie;
			if (::ActivateActCtx(hCtx, &cookie)) {
				hr = CoCreateInstance(__uuidof(DiaSource),
						   NULL,
						   CLSCTX_INPROC_SERVER,
						   __uuidof(IDiaDataSource),
						   (void **)ppSource);
				::DeactivateActCtx(0, cookie);
				if (FAILED(hr)) {
					wprintf(L"CoCreateInstance failed - HRESULT = %08X\n", hr);
					return false;
				}
			}
		}
	}
	if (FAILED(hr)) {
		wprintf(L"CoCreateInstance failed - HRESULT = %08X\n", hr);

		return false;
	}

	return true;
}
bool LayoutMatches(IDiaSymbol * pSymbol1, IDiaSymbol * pSymbol2)
{
	DWORD dwTag1, dwTag2;
	DWORD dwLocType1, dwLocType2;
	LONG lOffset1, lOffset2;
	DWORD lType1, lType2;

	if (pSymbol1->get_symTag(&dwTag1) != S_OK) {
		wprintf(L"ERROR - can't retrieve the symbol's SymTag\n");
		return false;
	}
	if (pSymbol2->get_symTag(&dwTag2) != S_OK) {
		wprintf(L"ERROR - can't retrieve the symbol's SymTag\n");
		return false;
	}

	if (dwTag1 == SymTagUDT) {
		if (dwTag2 != SymTagUDT) {

			wprintf(L"ERROR - symbols don't match\n");
			wprintf(L"Symbol 1:\n");
			PrintTypeInDetail(pSymbol1, 0);
			wprintf(L"Symbol 2:\n");
			PrintTypeInDetail(pSymbol2, 0);
			return false;
		}

		// First check that types size match
		ULONGLONG sym1Size;
		ULONGLONG sym2Size;
		if (pSymbol1->get_length(&sym1Size) != S_OK) {
			wprintf(L"ERROR - can't retrieve the symbol's length\n");
			return false;
		}
		if (pSymbol2->get_length(&sym2Size) != S_OK) {
			wprintf(L"ERROR - can't retrieve the symbol's length\n");
			return false;
		}
		if (sym1Size == 0 || sym2Size == 0) {
			return true;
		}
		if (sym1Size != sym2Size) {
			wprintf(L"Failed to match type size: (sizeof(sym1)=%llu) != (sizeof(sym2)=%llu)\n", sym1Size, sym2Size);
			return false;
		}
		IDiaEnumSymbols * pEnumChildren1, * pEnumChildren2;
		IDiaSymbol * pChild1, * pChild2;
		ULONG celt = 0;
		BSTR bstrName1, bstrName2;
		IDiaSymbol * pBaseType1, * pBaseType2;

		if (SUCCEEDED(pSymbol1->findChildren(SymTagNull, NULL, nsNone, &pEnumChildren1))) {
			while (SUCCEEDED(pEnumChildren1->Next(1, &pChild1, &celt)) && (celt == 1)) {
				if (pChild1->get_symTag(&dwTag1) != S_OK) {
					wprintf(L"ERROR - can't retrieve the symbol's SymTag\n");
					pChild1->Release();
					return false;
				}
				if (dwTag1 != SymTagData) { pChild1->Release(); continue; }
				if (pChild1->get_locationType(&dwLocType1) != S_OK) {
					wprintf(L"symbol in optmized code");
					pChild1->Release();
					return false;
				}
				if (dwLocType1 != LocIsThisRel) { pChild1->Release(); continue; }
				if (pChild1->get_offset(&lOffset1) != S_OK) {
					wprintf(L"ERROR - geting field offset\n");
					pChild1->Release();
					return false;
				}
				if (pChild1->get_type(&pBaseType1) != S_OK || pBaseType1->get_baseType(&lType1) != S_OK) {
					//wprintf(L"ERROR - geting field type\n");
					//pChild1->Release();
					//return false;
				}

				if (pChild1->get_name(&bstrName1) != S_OK) {
					bstrName1 = NULL;
				}
				/// Search in the second symbol the field at lOffset1
				bool fieldMatched = false;
				if (SUCCEEDED(pSymbol2->findChildren(SymTagNull, NULL, nsNone, &pEnumChildren2))) {
					while (SUCCEEDED(pEnumChildren2->Next(1, &pChild2, &celt)) && (celt == 1)) {
						if (pChild2->get_symTag(&dwTag2) != S_OK) {
							wprintf(L"ERROR - can't retrieve the symbol's SymTag\n");
							pChild2->Release();
							return false;
						}
						if (dwTag2 != SymTagData) { pChild2->Release(); continue; }
						if (pChild2->get_locationType(&dwLocType2) != S_OK) {
							wprintf(L"symbol in optmized code");
							pChild2->Release();
							return false;
						}
						if (dwLocType2 != LocIsThisRel) { pChild2->Release(); continue; }
						if (pChild2->get_offset(&lOffset2) != S_OK) {
							wprintf(L"ERROR - geting field offset\n");
							pChild2->Release();
							return false;
						}

						if (pChild2->get_type(&pBaseType2) != S_OK || pBaseType2->get_baseType(&lType2) != S_OK) {
							//wprintf(L"ERROR - geting field type\n");
							//pChild2->Release();
							//return false;
						}

						// my addition
						if (lType1 != lType2) {
							wprintf(L"Failed to match type of field %s at offset %x\n", bstrName1, lOffset1);
							pChild1->Release();
							pChild2->Release();
							return false;
						}


						if (pChild2->get_name(&bstrName2) != S_OK) {
							bstrName2 = NULL;
						}
						if (lOffset2 == lOffset1) {
							if (_wcsicmp(bstrName1, bstrName2) == 0
							|| wcsstr(bstrName1, bstrName2) == bstrName1
							|| wcsstr(bstrName2, bstrName1) == bstrName2) {
								//wprintf(L"Matched field %s at offset %x\n", bstrName1, lOffset2);
								fieldMatched = true;
								pChild2->Release();
								break;
							}
						}
						pChild2->Release();
					}
					pEnumChildren2->Release();
				}
				if (!fieldMatched) {
					BSTR bstrSymbol1Name;
					if (pSymbol1->get_name(&bstrSymbol1Name) != S_OK) {
						bstrSymbol1Name = NULL;
					}
					wprintf(L"Failed to match %s field %s at offset %x\n", bstrSymbol1Name, bstrName1, lOffset1);
					pChild1->Release();
					return false;
				}
				pChild1->Release();
			}

			pEnumChildren1->Release();
		}
	}

	return true;
}

////////////////////////////////////////////////////////////
// Release DIA objects and CoUninitialize
//
void Cleanup2()
{
	if (g_pGlobalSymbol1) {
		g_pGlobalSymbol1->Release();
		g_pGlobalSymbol1 = NULL;
	}

	if (g_pGlobalSymbol2) {
		g_pGlobalSymbol2->Release();
		g_pGlobalSymbol2 = NULL;
	}

	if (g_pDiaSession1) {
		g_pDiaSession1->Release();
		g_pDiaSession1 = NULL;
	}

	if (g_pDiaSession2) {
		g_pDiaSession2->Release();
		g_pDiaSession2 = NULL;
	}

	CoUninitialize();
}


////////////////////////////////////////////////////////////
// Display the usage
//
void PrintHelpOptions2()
{
	static const wchar_t * const helpString = L"usage: PdbTypeMatch.exe <pdb_filename_1> <pdb_filename_2> <type_exclusion_list_file> : compare all common types by size and fields\n"
		L"       PdbTypeMatch.exe -type <symbolname>  <pdb_filename_1>: dump this type in detail\n";

	wprintf(helpString);
}

bool EnumTypesInPdb(IDiaSymbolSet * types, IDiaSession * pSession, IDiaSymbol * pGlobal)
{
	IDiaEnumSymbols * pEnumSymbols;

	if (FAILED(pGlobal->findChildren(SymTagUDT, NULL, nsNone, &pEnumSymbols))) {
		wprintf(L"ERROR - EnumTypesInPdb() returned no symbols\n");

		return false;
	}

	IDiaSymbol * pSymbol;
	ULONG celt = 0;

	while (SUCCEEDED(pEnumSymbols->Next(1, &pSymbol, &celt)) && (celt == 1)) {
		std::wstring typeName;
		GetSymbolName(typeName, pSymbol);
		types->insert(std::wstring(typeName));
		pSymbol->Release();
	}

	pEnumSymbols->Release();

	return true;
}

//my addition
bool DumpAllSpecificDwords(IDiaSession * pSession, wchar_t * filename, wchar_t *name)
{
	wprintf(L"\n\n*** SPECIFIC DWORDS\n");

	IDiaEnumSectionContribs * pEnumSecContribs;

	if (FAILED(GetTable(pSession, __uuidof(IDiaEnumSectionContribs), (void **)&pEnumSecContribs))) {
		return false;
	}

	IDiaSectionContrib * pSecContrib;
	ULONG celt = 0;

	if (SUCCEEDED(pEnumSecContribs->Next(1, &pSecContrib, &celt)) && (celt == 1)) {
		PrintSpecificDword(pSession, pSecContrib, filename, name);
		pSecContrib->Release();
		putwchar(L'\n');
	}

	pEnumSecContribs->Release();

	putwchar(L'\n');

	return true;
}

bool DumpCompilandContrib(IDiaSession * pSession, IDiaSymbol * pGlobal, const wchar_t * szCompName)
{
	wprintf(L"\n\n*** COMPILAND SECTION CONTRIBUTION\n\n");

	std::wstring s = szCompName;
	bool wasfound = false;

	IDiaEnumSectionContribs * pEnumSecContribs;

	if (FAILED(GetTable(pSession, __uuidof(IDiaEnumSectionContribs), (void **)&pEnumSecContribs))) {
		return false;
	}

	//wprintf(L"    RVA        Address       Size      Module\n");

	IDiaSectionContrib * pSecContrib;
	ULONG celt = 0;

	DWORD sect = 0;
	while (SUCCEEDED(pEnumSecContribs->Next(1, &pSecContrib, &celt)) && (celt == 1)) {
		IDiaSymbol *pCompiland;
		if (FAILED(pSecContrib->get_compiland(&pCompiland))) {
			continue;
		}
		pSecContrib->get_addressSection(&sect);
		//only care about text
		//if (sect != 1) {
		//  break;
		//}

		BSTR bstrName;
		std::wstring n;
		n.clear();
		if (pCompiland->get_name(&bstrName) == S_OK) {
			n = bstrName;
			SysFreeString(bstrName);
		}

		if (!n.empty() && n.find(s) != std::string::npos) {
			if (!wasfound) {
				wprintf(L"%s\n", n.c_str());
				wasfound = true;
			}
			PrintSecContribs(pSession, pSecContrib);
			putwchar(L'\n');

		}

		pSecContrib->Release();

	}

	pEnumSecContribs->Release();

	if (!wasfound)
	{
		wprintf(L"Could not find %s\n\n", s.c_str());
	}

	putwchar(L'\n');

	return true;
}