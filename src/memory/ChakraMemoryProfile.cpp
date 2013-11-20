#include "stdafx.h"
#include <activprof.h>
#include <msopc.h>
#include <string>
#include <stack>
#include <queue>

using namespace std;

typedef void *MemoryProfileHandle;

static IOpcFactory *factory = nullptr;

struct MemoryProfile
{
    IOpcPackage *package;
    IOpcPartSet *partSet;
    int snapshotCount;

    MemoryProfile() :
        package(nullptr),
        partSet(nullptr),
        snapshotCount(0)
    {
    }
};

#define IfComFailError(v) \
	{ \
		hr = (v) ; \
		if (FAILED(hr)) \
		{ \
			goto error; \
		} \
	}

#define IfComFailRet(v) \
	{ \
		HRESULT hr = (v) ; \
		if (FAILED(hr)) \
		{ \
			return hr; \
		} \
	}

class JsonSerializer
{
public:
	JsonSerializer(IStream *stream) :
		_stream(stream)
	{
	}

	HRESULT EndArray()
	{
		IfComFailRet(Write(L"]"));
		_scopeStack.pop();
		return S_OK;
	}

	HRESULT EndProfile()
	{
		return S_OK;
	}

	HRESULT EndSummary()
	{
		return S_OK;
	}

	HRESULT EndProperty()
	{
		_scopeStack.pop();
		return S_OK;
	}

	HRESULT EndHeapObject()
	{
		IfComFailRet(EndJsonObject());
		IfComFailRet(EndArray());
		IfComFailRet(EndJsonObject());
		return S_OK;
	}

	HRESULT EndJsonObject()
	{
		IfComFailRet(Write(L"}"));
		_scopeStack.pop();
		return S_OK;
	}

	HRESULT StartArray()
	{
		IfComFailRet(AppendDelimiterIfNecessary());
		IfComFailRet(Write(L"["));
		_scopeStack.push(false);
		return S_OK;
	}

	HRESULT StartProfile()
	{
		IfComFailRet(WriteBOM());
		IfComFailRet(StartJsonObject());
		IfComFailRet(WriteVersion());
		IfComFailRet(WriteTimestamp());
		IfComFailRet(EndJsonObject());
		return S_OK;
	}

	HRESULT StartSummary()
	{
		IfComFailRet(WriteBOM());
		return S_OK;
	}

	HRESULT StartHeapObject()
	{
		IfComFailRet(WriteNewLine());
		IfComFailRet(StartJsonObject());
		IfComFailRet(WriteVersion());
		IfComFailRet(StartProperty(L"data"));
		IfComFailRet(StartArray());
		IfComFailRet(StartJsonObject());
		return S_OK;
	}

	HRESULT StartJsonObject()
	{
		IfComFailRet(Write(L"{"));
		return S_OK;
	}

	HRESULT StartJsonObjectNested()
	{
		IfComFailRet(AppendDelimiterIfNecessary());
		IfComFailRet(Write(L"{"));
		_scopeStack.push(false);
		return S_OK;
	}

	HRESULT StartProperty(const wchar_t * name)
	{
		IfComFailRet(AppendDelimiterIfNecessary());
		IfComFailRet(AppendPropertyToData(name));
		_scopeStack.push(false);
		return S_OK;
	}

	HRESULT WriteProperty(const wchar_t * name, const wchar_t * value)
	{
		IfComFailRet(AppendDelimiterIfNecessary());
		IfComFailRet(AppendPropertyAndValueToData(name, EncodeAsJsonString(value).c_str()));
		return S_OK;
	}

	HRESULT WriteProperty(const wchar_t * name, const int value)
	{
		IfComFailRet(AppendDelimiterIfNecessary());
		wchar_t buffer[11];
		int iResult = _itow_s(value, buffer, ARRAYSIZE(buffer), 10);
		if (iResult == 0)
		{
			IfComFailRet(AppendPropertyAndValueToDataNoQuotes(name, (const wchar_t *) buffer));
			return S_OK;
		}

		return E_FAIL;
	}

	HRESULT WriteProperty(const wchar_t * name, const unsigned value)
	{
		IfComFailRet(AppendDelimiterIfNecessary());
		wchar_t buffer[11];
		int iResult = _itow_s(value, buffer, ARRAYSIZE(buffer), 10);
		if (iResult == 0)
		{
			IfComFailRet(AppendPropertyAndValueToDataNoQuotes(name, (const wchar_t *) buffer));
			return S_OK;
		}

		return E_FAIL;
	}

	HRESULT WriteProperty(const wchar_t * name, const double value)
	{
		IfComFailRet(AppendDelimiterIfNecessary());
		wchar_t buffer[256];
		if (swprintf_s(buffer, ARRAYSIZE(buffer), L"%g", value) != -1)
		{
			IfComFailRet(AppendPropertyAndValueToDataNoQuotes(name, (const wchar_t *) buffer));
			return S_OK;
		}

		return E_FAIL;
	}

	HRESULT WriteProperty(const wchar_t * name, const bool value)
	{
		IfComFailRet(AppendDelimiterIfNecessary());
		IfComFailRet(AppendPropertyAndValueToDataNoQuotes(name, value ? L"true" : L"false"));
		return S_OK;
	}

	HRESULT WriteValue(const wchar_t * value)
	{
		IfComFailRet(AppendDelimiterIfNecessary());
		IfComFailRet(AppendValueToData(EncodeAsJsonString(value).c_str()));
		return S_OK;
	}

	HRESULT WriteValue(const int value)
	{
		IfComFailRet(AppendDelimiterIfNecessary());
		wchar_t buffer[11];
		int iResult = _itow_s(value, buffer, ARRAYSIZE(buffer), 10);
		if (iResult == 0)
		{
			IfComFailRet(AppendValueToDataNoQuotes((const wchar_t *) buffer));
			return S_OK;
		}

		return E_FAIL;
	}

private:
	HRESULT WriteBOM()
	{
		ULONG written;
		const char *BOM = "\xEF\xBB\xBF";
		IfComFailRet(_stream->Write(BOM, strlen(BOM), &written));
		return S_OK;
	}

	HRESULT Write(const wchar_t *s)
	{
		ULONG written;
		unsigned characterCount = wcslen(s);
		char smallBuffer[256];
		char *buffer;
		bool freeBuffer = false;
		HRESULT hr = S_OK;

		if (characterCount * 2 < 256)
		{
			buffer = smallBuffer;
		}
		else
		{
			buffer = new char[characterCount * 2];
			freeBuffer = true;
		}

		int byteCount = WideCharToMultiByte(CP_UTF8, 0, s, characterCount, buffer, characterCount * 2, nullptr, nullptr);

		if (byteCount)
		{
			hr = _stream->Write(buffer, byteCount, &written);
		}

		if (freeBuffer)
		{
			delete buffer;
		}

		return S_OK;
	}

	HRESULT AppendDelimiterIfNecessary()
	{
		if (_scopeStack.size() > 0)
		{
			if (_scopeStack.top())
			{
				// Append the delimiter
				IfComFailRet(Write(L","));
			}
			else
			{
				// Change this scope so the next time
				// we will append a delimiter.
				_scopeStack.pop();
				_scopeStack.push(true);
			}
		}

		return S_OK;
	}

	HRESULT AppendPropertyToData(const wchar_t * name)
	{
		IfComFailRet(Write(L"\""));
		IfComFailRet(Write(name));
		IfComFailRet(Write(L"\":"));
		return S_OK;
	}

	HRESULT AppendPropertyAndValueToData(const wchar_t * name, const wchar_t * value)
	{
		IfComFailRet(Write(L"\""));
		IfComFailRet(Write(name));
		IfComFailRet(Write(L"\":\""));
		IfComFailRet(Write(value));
		IfComFailRet(Write(L"\""));
		return S_OK;
	}

	HRESULT AppendPropertyAndValueToDataNoQuotes(const wchar_t * name, const wchar_t * value)
	{
		IfComFailRet(Write(L"\""));
		IfComFailRet(Write(name));
		IfComFailRet(Write(L"\":"));
		IfComFailRet(Write(value));
		return S_OK;
	}

	HRESULT AppendValueToData(const wchar_t * value)
	{
		IfComFailRet(Write(L"\""));
		IfComFailRet(Write(value));
		IfComFailRet(Write(L"\""));
		return S_OK;
	}

	HRESULT AppendValueToDataNoQuotes(const wchar_t * value)
	{
		IfComFailRet(Write(value));
		return S_OK;
	}

	wstring EncodeAsJsonString(const wchar_t * value)
	{
		wstring escapedValue = L"";

		while (*value != L'\0')
		{
			switch (*value)
			{
			case L'"':
				escapedValue.append(L"\\\"");
				break;
			case L'/':
				escapedValue.append(L"\\/");
				break;
			case L'\\':
				escapedValue.append(L"\\\\");
				break;
			case L'\b':
				escapedValue.append(L"\\b");
				break;
			case L'\f':
				escapedValue.append(L"\\f");
				break;
			case L'\n':
				escapedValue.append(L"\\n");
				break;
			case L'\r':
				escapedValue.append(L"\\r");
				break;
			case L'\t':
				escapedValue.append(L"\\t");
				break;
			default:
				// Based on the JSON specification (RFC 4627) any character can
				// be escaped, but only the above characters and control
				// characters (U+0000 through U+001F) must be escaped.  However,
				// we will also escape extended characters (U+007F-) to be safe.
				// Example:  the unescaped EOF character 0xFFFF cannot be parsed.
				if (((*value) >= 0x0000 && (*value) <= 0x001F) || (*value) > 0x007F)
				{
					escapedValue.append(L"\\u");

					// Go from the int value to the hex value
					int intValue = (int) (*value);
					wchar_t buffer[11];
					int iResult = _itow_s(intValue, buffer, ARRAYSIZE(buffer), 16);
					if (iResult == 0)
					{
						long zerosToAppend = 4 - wcslen((const wchar_t *) buffer);
						if (zerosToAppend >= 0)
						{
							for (int index = 0; index < zerosToAppend; index++)
							{
								escapedValue.append(L"0", wcslen(L"0"));
							}
							escapedValue.append((const wchar_t *) buffer);
						}
					}
				}
				else
				{
					escapedValue.append(value, 1);
				}
				break;
			}
			++value;
		}

		return escapedValue;
	}

	HRESULT WriteNewLine()
	{
		IfComFailRet(Write(L"\r\n"));
		return S_OK;
	}

	HRESULT WriteVersion()
	{
		IfComFailRet(Write(L"\"version\":\"1.0\""));
		_scopeStack.push(true);
		return S_OK;
	}

	HRESULT WriteTimestamp()
	{
		IfComFailRet(AppendDelimiterIfNecessary());

		LARGE_INTEGER time;
		QueryPerformanceCounter(&time);

		LONGLONG longTime = time.QuadPart;
		wchar_t buffer[20];
		int iResult = _i64tow_s(longTime, buffer, ARRAYSIZE(buffer), 10);
		if (iResult == 0)
		{
			IfComFailRet(AppendPropertyAndValueToData(L"timestamp", (const wchar_t *) buffer));
			return S_OK;
		}

		return E_FAIL;
	}

	IStream *_stream;

	// If the current scope == true, then we append a delimiter.
	stack<bool> _scopeStack;
};

enum ExternalObjectKind {
	ExternalObjectKind_Default = 1,
	ExternalObjectKind_Unknown = 2,
	ExternalObjectKind_Dispatch = 3,
};

enum WinRTObjectKind {
	WinRTObjectKind_Instance = 1,
	WinRTObjectKind_RunTimeClass = 2,
	WinRTObjectKind_Delegate = 3,
	WinRTObjectKind_NameSpace = 4,
};

const wchar_t * GetTypeName(const wchar_t **nameIdMap, UINT nameCount, ULONG ulId)
{
	const wchar_t * name = L"<Type Name Not Found>";

	if (ulId < nameCount)
	{
		name = nameIdMap[ulId];
		if (name == NULL)
		{
			name = L"<Type Name Not Found>";
		}
	}
	return name;
}

unsigned AddSizes(unsigned uSizeToAdd, unsigned uSize)
{
	if ((uSize + uSizeToAdd) < uSizeToAdd)
	{
		return UINT_MAX;
	}
	return uSize + uSizeToAdd;
}

HRESULT SerializeIdProperty(JsonSerializer *snapshotSerializer, const wchar_t * name, ULONG_PTR ulId)
{
#ifdef _WIN64
	if (ulId > MAXINT32)
	{
		wchar_t szId[21] = L""; // Maximum decimal in UINT64 is 20
		_ui64tow_s(ulId, szId, ARRAYSIZE(szId), 10);
		IfComFailRet(snapshotSerializer->WriteProperty(name, szId));
	}
	else
#endif
	{
		IfComFailRet(snapshotSerializer->WriteProperty(name, (int) ulId));
	}

	return S_OK;
}

HRESULT SerializeProperty(JsonSerializer *snapshotSerializer, const wchar_t **nameIdMap, UINT nameCount, PROFILER_HEAP_OBJECT_RELATIONSHIP *profilerHeapObjectProperty, bool indexList)
{
	IfComFailRet(snapshotSerializer->StartJsonObjectNested());
	if (profilerHeapObjectProperty->relationshipId != PROFILER_HEAP_OBJECT_NAME_ID_UNAVAILABLE)
	{
		wchar_t indexName[265] = L"";
		const wchar_t * name = nullptr;

		if (indexList)
		{
			wchar_t indexString[265];
			_itow_s(profilerHeapObjectProperty->relationshipId, indexString, 10);
			wcsncat_s(indexName, L"[", _TRUNCATE);
			wcsncat_s(indexName, indexString, _TRUNCATE);
			wcsncat_s(indexName, L"]", _TRUNCATE);
			name = indexName;
		}
		else
		{
			name = GetTypeName(nameIdMap, nameCount, profilerHeapObjectProperty->relationshipId);
		}

		if (name != nullptr && wcscmp(name, L"") != 0)
		{
			IfComFailRet(snapshotSerializer->WriteProperty(L"name", name));
		}
	}

	switch (profilerHeapObjectProperty->relationshipInfo)
	{
	case PROFILER_PROPERTY_TYPE_NUMBER:
		if (_isnan(profilerHeapObjectProperty->numberValue))
		{
			IfComFailRet(snapshotSerializer->WriteProperty(L"numberValue", L"NaN"));
		}
		else if (!_finite(profilerHeapObjectProperty->numberValue))
		{
			if (_fpclass(profilerHeapObjectProperty->numberValue) == _FPCLASS_PINF)
			{
				IfComFailRet(snapshotSerializer->WriteProperty(L"numberValue", L"Infinity"));
			}
			else
			{
				IfComFailRet(snapshotSerializer->WriteProperty(L"numberValue", L"-Infinity"));
			}
		}
		else
		{
			IfComFailRet(snapshotSerializer->WriteProperty(L"numberValue", profilerHeapObjectProperty->numberValue));
		}
		break;

	case PROFILER_PROPERTY_TYPE_STRING:
		IfComFailRet(snapshotSerializer->WriteProperty(L"stringValue", profilerHeapObjectProperty->stringValue));
		break;

	case PROFILER_PROPERTY_TYPE_HEAP_OBJECT:
		IfComFailRet(SerializeIdProperty(snapshotSerializer, L"objectId", profilerHeapObjectProperty->objectId));
		break;

	case PROFILER_PROPERTY_TYPE_EXTERNAL_OBJECT:
		IfComFailRet(SerializeIdProperty(snapshotSerializer, L"objectId", (ULONG_PTR) profilerHeapObjectProperty->externalObjectAddress));
		break;

	case PROFILER_PROPERTY_TYPE_BSTR:
		IfComFailRet(snapshotSerializer->WriteProperty(L"stringValue", profilerHeapObjectProperty->bstrValue));
		break;

	default:
		IfComFailRet(snapshotSerializer->WriteProperty(L"UNKNOWN relationshipinfo", L"UNKNOWN"));
		break;
	}

	IfComFailRet(snapshotSerializer->EndJsonObject());

	return S_OK;
}

HRESULT SerializeHeapObjectFlags(JsonSerializer *snapshotSerializer, PROFILER_HEAP_OBJECT *profilerHeapObject)
{
	// flags
	if (!(profilerHeapObject->flags & PROFILER_HEAP_OBJECT_FLAGS_NEW_STATE_UNAVAILABLE))
	{
		IfComFailRet(snapshotSerializer->WriteProperty(L"isNew", (bool) (profilerHeapObject->flags & PROFILER_HEAP_OBJECT_FLAGS_NEW_OBJECT)));
	}

	if (profilerHeapObject->flags & PROFILER_HEAP_OBJECT_FLAGS_IS_ROOT)
	{
		IfComFailRet(snapshotSerializer->WriteProperty(L"isRoot", true));
	}

	if (profilerHeapObject->flags & PROFILER_HEAP_OBJECT_FLAGS_SITE_CLOSED)
	{
		IfComFailRet(snapshotSerializer->WriteProperty(L"isSiteClosed", true));
	}

	// external flag
	if (profilerHeapObject->flags & PROFILER_HEAP_OBJECT_FLAGS_EXTERNAL)
	{
		IfComFailRet(snapshotSerializer->WriteProperty(L"external", ExternalObjectKind_Default));
	}
	else if (profilerHeapObject->flags & PROFILER_HEAP_OBJECT_FLAGS_EXTERNAL_UNKNOWN)
	{
		IfComFailRet(snapshotSerializer->WriteProperty(L"external", ExternalObjectKind_Unknown));
	}
	else if (profilerHeapObject->flags & PROFILER_HEAP_OBJECT_FLAGS_EXTERNAL_DISPATCH)
	{
		IfComFailRet(snapshotSerializer->WriteProperty(L"external", ExternalObjectKind_Dispatch));
	}

	// winrt flags
	if (profilerHeapObject->flags & PROFILER_HEAP_OBJECT_FLAGS_WINRT_INSTANCE)
	{
		IfComFailRet(snapshotSerializer->WriteProperty(L"winrt", WinRTObjectKind_Instance));
	}
	else if (profilerHeapObject->flags & PROFILER_HEAP_OBJECT_FLAGS_WINRT_RUNTIMECLASS)
	{
		IfComFailRet(snapshotSerializer->WriteProperty(L"winrt", WinRTObjectKind_RunTimeClass));
	}
	else if (profilerHeapObject->flags & PROFILER_HEAP_OBJECT_FLAGS_WINRT_DELEGATE)
	{
		IfComFailRet(snapshotSerializer->WriteProperty(L"winrt", WinRTObjectKind_Delegate));
	}
	else if (profilerHeapObject->flags & PROFILER_HEAP_OBJECT_FLAGS_WINRT_NAMESPACE)
	{
		IfComFailRet(snapshotSerializer->WriteProperty(L"winrt", WinRTObjectKind_NameSpace));
	}

	return S_OK;
}

HRESULT SerializePropertyList(JsonSerializer *snapshotSerializer, const wchar_t **nameIdMap, UINT nameCount, const wchar_t * propertyListName, PROFILER_HEAP_OBJECT_RELATIONSHIP_LIST *propertyList, bool indexList)
{
	IfComFailRet(snapshotSerializer->StartProperty(propertyListName));
	IfComFailRet(snapshotSerializer->StartArray());

	for (unsigned index = 0; index < propertyList->count; index++)
	{
		IfComFailRet(SerializeProperty(snapshotSerializer, nameIdMap, nameCount, &propertyList->elements[index], indexList));
	}

	IfComFailRet(snapshotSerializer->EndArray());
	IfComFailRet(snapshotSerializer->EndProperty());

	return S_OK;
}

HRESULT SerializeIdValue(JsonSerializer *snapshotSerializer, ULONG_PTR ulId)
{
	HRESULT hr = S_OK;
#ifdef _WIN64
	if (ulId > MAXINT32)
	{
		wchar_t id[21] = L""; // Maximum decimal in UINT64 is 20
		_ui64tow_s(ulId, id, ARRAYSIZE(id), 10);
		IfComFailRet(snapshotSerializer->WriteValue(id);)
	}
	else
#endif
	{
		IfComFailRet(snapshotSerializer->WriteValue((int) ulId));
	}

	return S_OK;
}

HRESULT SerializeKeyValuePropertyList(JsonSerializer *snapshotSerializer, const wchar_t **nameIdMap, UINT nameCount, const wchar_t * propertyListName, PROFILER_HEAP_OBJECT_RELATIONSHIP_LIST *propertyList)
{
	IfComFailRet(snapshotSerializer->StartProperty(propertyListName));
	IfComFailRet(snapshotSerializer->StartArray());
	for (unsigned index = 0; (index + 1) < propertyList->count; index = index + 2)
	{
		IfComFailRet(snapshotSerializer->StartJsonObjectNested());
		IfComFailRet(snapshotSerializer->StartProperty(L"key"));
		IfComFailRet(SerializeProperty(snapshotSerializer, nameIdMap, nameCount, &propertyList->elements[index], false));
		IfComFailRet(snapshotSerializer->EndProperty());
		IfComFailRet(snapshotSerializer->StartProperty(L"value"));
		IfComFailRet(SerializeProperty(snapshotSerializer, nameIdMap, nameCount, &propertyList->elements[index + 1], false));
		IfComFailRet(snapshotSerializer->EndProperty());
		IfComFailRet(snapshotSerializer->EndJsonObject());
	}

	IfComFailRet(snapshotSerializer->EndArray());
	IfComFailRet(snapshotSerializer->EndProperty());

	return S_OK;
}

HRESULT SerializeHeapObjectOptionalInfo(IActiveScriptProfilerHeapEnum *enumerator, JsonSerializer *snapshotSerializer, const wchar_t **nameIdMap, UINT nameCount, PROFILER_HEAP_OBJECT *profilerHeapObject, unsigned *size)
{
	*size = profilerHeapObject->size;
	if (profilerHeapObject->optionalInfoCount > 0)
	{
		HRESULT hr = S_OK;
		unsigned optionalInfoCount = profilerHeapObject->optionalInfoCount;
		PROFILER_HEAP_OBJECT_OPTIONAL_INFO *optionalInfo = new PROFILER_HEAP_OBJECT_OPTIONAL_INFO[optionalInfoCount];
		queue<unsigned> internalProperties;

		if (optionalInfo == nullptr)
		{
			return E_FAIL;
		}

		IfComFailError(enumerator->GetOptionalInfo(profilerHeapObject, optionalInfoCount, optionalInfo));

		for (unsigned index = 0; index < optionalInfoCount; index++)
		{
			switch (optionalInfo[index].infoType)
			{
			case PROFILER_HEAP_OBJECT_OPTIONAL_INFO_NAME_PROPERTIES:
				IfComFailError(SerializePropertyList(snapshotSerializer, nameIdMap, nameCount, L"properties", optionalInfo[index].namePropertyList, false));
				*size = AddSizes((unsigned) (optionalInfo[index].namePropertyList)->count * sizeof(void*) , *size);
				break;
			case PROFILER_HEAP_OBJECT_OPTIONAL_INFO_INDEX_PROPERTIES:
				IfComFailError(SerializePropertyList(snapshotSerializer, nameIdMap, nameCount, L"indices", optionalInfo[index].indexPropertyList, true));
				*size = AddSizes((unsigned) (optionalInfo[index].indexPropertyList)->count * sizeof(void*) , *size);
				break;
			case PROFILER_HEAP_OBJECT_OPTIONAL_INFO_RELATIONSHIPS:
				IfComFailError(SerializePropertyList(snapshotSerializer, nameIdMap, nameCount, L"relationships", optionalInfo[index].relationshipList, false));
				break;
			case PROFILER_HEAP_OBJECT_OPTIONAL_INFO_WINRTEVENTS:
				IfComFailError(SerializePropertyList(snapshotSerializer, nameIdMap, nameCount, L"events", optionalInfo[index].eventList, false));
				break;
			case PROFILER_HEAP_OBJECT_OPTIONAL_INFO_INTERNAL_PROPERTY:
				internalProperties.push(index);
				break;
			case PROFILER_HEAP_OBJECT_OPTIONAL_INFO_PROTOTYPE:
				IfComFailError(SerializeIdProperty(snapshotSerializer, L"prototype", optionalInfo[index].prototype));
				break;
			case PROFILER_HEAP_OBJECT_OPTIONAL_INFO_FUNCTION_NAME:
				IfComFailError(snapshotSerializer->WriteProperty(L"functionName", optionalInfo[index].functionName));
				break;
			case PROFILER_HEAP_OBJECT_OPTIONAL_INFO_SCOPE_LIST:
				IfComFailError(snapshotSerializer->StartProperty(L"scopes"));
				IfComFailError(snapshotSerializer->StartArray());
				for (unsigned scopeIndex = 0; scopeIndex < optionalInfo[index].scopeList->count; scopeIndex++)
				{
					IfComFailError(SerializeIdValue(snapshotSerializer, optionalInfo[index].scopeList->scopes[scopeIndex]));
				}

				IfComFailError(snapshotSerializer->EndArray());
				IfComFailError(snapshotSerializer->EndProperty());
				break;
			case PROFILER_HEAP_OBJECT_OPTIONAL_INFO_ELEMENT_ATTRIBUTES_SIZE:
				IfComFailError(snapshotSerializer->WriteProperty(L"elementAttributesSize", optionalInfo[index].elementAttributesSize));
				*size = AddSizes((unsigned) optionalInfo[index].elementAttributesSize, *size);
				break;
			case PROFILER_HEAP_OBJECT_OPTIONAL_INFO_ELEMENT_TEXT_CHILDREN_SIZE:
				IfComFailError(snapshotSerializer->WriteProperty(L"elementTextChildrenSize", optionalInfo[index].elementTextChildrenSize));
				*size = AddSizes((unsigned) optionalInfo[index].elementTextChildrenSize, *size);
				break;
			case PROFILER_HEAP_OBJECT_OPTIONAL_INFO_WEAKMAP_COLLECTION_LIST:
				IfComFailError(SerializeKeyValuePropertyList(snapshotSerializer, nameIdMap, nameCount, L"map", optionalInfo[index].weakMapCollectionList));
				*size = AddSizes((unsigned) (optionalInfo[index].weakMapCollectionList)->count * sizeof(void*) , *size);
				break;
			case PROFILER_HEAP_OBJECT_OPTIONAL_INFO_MAP_COLLECTION_LIST:
				IfComFailError(SerializeKeyValuePropertyList(snapshotSerializer, nameIdMap, nameCount, L"map", optionalInfo[index].mapCollectionList));
				*size = AddSizes((unsigned) (optionalInfo[index].mapCollectionList)->count * sizeof(void*) , *size);
				break;
			case PROFILER_HEAP_OBJECT_OPTIONAL_INFO_SET_COLLECTION_LIST:
				IfComFailError(SerializePropertyList(snapshotSerializer, nameIdMap, nameCount, L"set", optionalInfo[index].setCollectionList, false));
				*size = AddSizes((unsigned) (optionalInfo[index].setCollectionList)->count * sizeof(void*) , *size);
				break;
			}
		}

		if (internalProperties.size() > 0)
		{
			IfComFailError(snapshotSerializer->StartProperty(L"internalProperties"));
			IfComFailError(snapshotSerializer->StartArray());
			unsigned optionalInfoIndex = 0;
			while (internalProperties.size())
			{
				IfComFailError(SerializeProperty(snapshotSerializer, nameIdMap, nameCount, optionalInfo[internalProperties.front()].internalProperty, false));
				internalProperties.pop();
			}
			IfComFailError(snapshotSerializer->EndArray());
			IfComFailError(snapshotSerializer->EndProperty());
		}
	error:
		delete [] optionalInfo;
	}

	IfComFailRet(snapshotSerializer->WriteProperty(L"size", *size));
	return S_OK;
}

HRESULT SerializeObject(IActiveScriptProfilerHeapEnum *enumerator, JsonSerializer *snapshotSerializer, const wchar_t **nameIdMap, UINT nameCount, PROFILER_HEAP_OBJECT *profilerHeapObject, unsigned *size)
{
	IfComFailRet(snapshotSerializer->StartHeapObject());
	IfComFailRet(SerializeIdProperty(snapshotSerializer, L"objectId", profilerHeapObject->objectId));

	// sizeIsApproximate
	if (!(profilerHeapObject->flags & PROFILER_HEAP_OBJECT_FLAGS_SIZE_UNAVAILABLE))
	{
		// Note: Size is serialized with the optional info
		if ((profilerHeapObject->flags & PROFILER_HEAP_OBJECT_FLAGS_SIZE_APPROXIMATE))
		{
			IfComFailRet(snapshotSerializer->WriteProperty(L"sizeIsApproximate", true));
		}
	}

	// typeNameId
	if (profilerHeapObject->typeNameId != PROFILER_HEAP_OBJECT_NAME_ID_UNAVAILABLE)
	{
		const wchar_t * pszTypeName = GetTypeName(nameIdMap, nameCount, profilerHeapObject->typeNameId);

		if (pszTypeName != nullptr)
		{
			IfComFailRet(snapshotSerializer->WriteProperty(L"kind", pszTypeName));
		}
	}

	IfComFailRet(SerializeHeapObjectFlags(snapshotSerializer, profilerHeapObject));
	IfComFailRet(SerializeHeapObjectOptionalInfo(enumerator, snapshotSerializer, nameIdMap, nameCount, profilerHeapObject, size));
	IfComFailRet(snapshotSerializer->EndHeapObject());

	return S_OK;
}

HRESULT GetNextHeapObject(IActiveScriptProfilerHeapEnum *enumerator, JsonSerializer *snapshotSerializer, const wchar_t **nameIdMap, UINT nameCount, bool *moreObjects, unsigned *size)
{
	PROFILER_HEAP_OBJECT *profilerHeapObject[1];
	ULONG fetchedObjectCount = 0;

	IfComFailRet(enumerator->Next(1, profilerHeapObject, &fetchedObjectCount));

	if (fetchedObjectCount == 0)
	{
		*moreObjects = false;
		return S_OK;
	}
	else
	{
		*moreObjects = true;
		IfComFailRet(SerializeObject(enumerator, snapshotSerializer, nameIdMap, nameCount, profilerHeapObject[0], size));
		return enumerator->FreeObjectAndOptionalInfo(fetchedObjectCount, profilerHeapObject);
	}
}

HRESULT WriteSnapshot(IActiveScriptProfilerHeapEnum *enumerator, IStream *snapshotPartStream, unsigned *objectsCount, unsigned *objectsSize)
{
	bool moreObjects = true;
	const wchar_t **nameIdMap = nullptr;
	UINT nameCount;
	JsonSerializer snapshotSerializer(snapshotPartStream);
	HRESULT hr = S_OK;

	IfComFailError(enumerator->GetNameIdMap(&nameIdMap, &nameCount));
	IfComFailError(snapshotSerializer.StartProfile());

	while (moreObjects)
	{
		unsigned size;
		(*objectsCount)++;
		IfComFailError(GetNextHeapObject(enumerator, &snapshotSerializer, nameIdMap, nameCount, &moreObjects, &size));

		if (moreObjects)
		{
			*objectsSize += size;
		}
	}

	IfComFailError(snapshotSerializer.EndProfile());

error:
	if (nameIdMap != nullptr)
	{
		CoTaskMemFree(nameIdMap);
		nameIdMap = nullptr;
	}

	return hr;
}

HRESULT WriteSummary(IStream *summaryPartStream, std::wstring snapshotName, unsigned id, unsigned objectsCount, unsigned objectsSize)
{
	JsonSerializer summarySerializer(summaryPartStream);

	IfComFailRet(summarySerializer.StartSummary());
	IfComFailRet(summarySerializer.StartJsonObjectNested());

	IfComFailRet(summarySerializer.StartProperty(L"snapshotFile"));
	IfComFailRet(summarySerializer.StartJsonObjectNested());
	IfComFailRet(summarySerializer.WriteProperty(L"relativePath", snapshotName.c_str()));
	IfComFailRet(summarySerializer.EndJsonObject());
	IfComFailRet(summarySerializer.EndProperty());

	IfComFailRet(summarySerializer.WriteProperty(L"totalObjectSize", objectsSize));
	IfComFailRet(summarySerializer.WriteProperty(L"objectsCount", objectsCount));

    IfComFailRet(summarySerializer.WriteProperty(L"id", id));

	IfComFailRet(summarySerializer.EndJsonObject());
	IfComFailRet(summarySerializer.EndSummary());

	return S_OK;
}

extern "C" __declspec(dllexport) bool InitializeMemoryProfileWriter()
{
    HRESULT hr = S_OK;
    IfComFailError(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED));
    IfComFailError(CoCreateInstance(__uuidof(OpcFactory), NULL, CLSCTX_INPROC_SERVER, __uuidof(IOpcFactory), (LPVOID*) &factory));

error:
    return SUCCEEDED(hr);
}

extern "C" __declspec(dllexport) MemoryProfileHandle StartMemoryProfile()
{
    HRESULT hr = S_OK;
    MemoryProfile *memoryProfile = nullptr;

    try
    {
        memoryProfile = new MemoryProfile();
        IfComFailError(factory->CreatePackage(&memoryProfile->package));
        IfComFailError(memoryProfile->package->GetPartSet(&memoryProfile->partSet));
        return (MemoryProfileHandle) memoryProfile;
    }
    catch (...)
    {
    }

error:
    if (memoryProfile)
    {
        if (memoryProfile->package)
        {
            memoryProfile->package->Release();
        }

        delete memoryProfile;
    }

    return nullptr;
}

extern "C" __declspec(dllexport) bool WriteSnapshot(MemoryProfileHandle memoryProfileHandle, IActiveScriptProfilerHeapEnum *enumerator)
{
    HRESULT hr = S_OK;
    MemoryProfile *memoryProfile = (MemoryProfile *) memoryProfileHandle;

    IOpcPartUri *summaryPartUri = nullptr;
    IOpcPart *summaryPart = nullptr;
    IStream *summaryPartStream = nullptr;

    IOpcPartUri *snapshotPartUri = nullptr;
    IOpcPart *snapshotPart = nullptr;
    IStream *snapshotPartStream = nullptr;

    std::wstring snapshotName = L"snapshot" + to_wstring(++(memoryProfile->snapshotCount)) + L".snapjs";

    IfComFailError(factory->CreatePartUri((snapshotName + L".snapshotsummary").c_str(), &summaryPartUri));
    IfComFailError(memoryProfile->partSet->CreatePart(summaryPartUri, L"application/json", OPC_COMPRESSION_NORMAL, &summaryPart));
    IfComFailError(summaryPart->GetContentStream(&summaryPartStream));

    IfComFailError(factory->CreatePartUri(snapshotName.c_str(), &snapshotPartUri));
    IfComFailError(memoryProfile->partSet->CreatePart(snapshotPartUri, L"application/json", OPC_COMPRESSION_NORMAL, &snapshotPart));
    IfComFailError(snapshotPart->GetContentStream(&snapshotPartStream));

    unsigned objectsCount = 0;
    unsigned objectsSize = 0;
    IfComFailError(WriteSnapshot(enumerator, snapshotPartStream, &objectsCount, &objectsSize))

    IfComFailError(WriteSummary(summaryPartStream, snapshotName.c_str(), memoryProfile->snapshotCount, objectsCount, objectsSize));

    summaryPartStream->Release();
    summaryPartStream = nullptr;

    snapshotPartStream->Release();
    snapshotPartStream = nullptr;

error:
    if (snapshotPartStream)
    {
        snapshotPartStream->Release();
    }

    if (snapshotPart)
    {
        snapshotPart->Release();
    }

    if (snapshotPartUri)
    {
        snapshotPartUri->Release();
    }

    if (summaryPartStream)
    {
        summaryPartStream->Release();
    }

    if (summaryPart)
    {
        summaryPart->Release();
    }

    if (summaryPartUri)
    {
        summaryPartUri->Release();
    }

    return SUCCEEDED(hr);
}

extern "C" __declspec(dllexport) bool EndMemoryProfile(MemoryProfileHandle memoryProfileHandle, const wchar_t *filename)
{
    HRESULT hr = S_OK;
    IStream *packageStream = nullptr;

    MemoryProfile *memoryProfile = (MemoryProfile *) memoryProfileHandle;

    IfComFailError(factory->CreateStreamOnFile(filename, OPC_STREAM_IO_WRITE, nullptr, 0, &packageStream));
    IfComFailError(factory->WritePackageToStream(memoryProfile->package, OPC_WRITE_DEFAULT, packageStream));

error:
    if (packageStream)
    {
        packageStream->Release();
    }

    if (memoryProfile->partSet)
    {
        memoryProfile->partSet->Release();
    }

    if (memoryProfile->package)
    {
        memoryProfile->package->Release();
    }

    delete memoryProfile;

    return SUCCEEDED(hr);
}

extern "C" __declspec(dllexport) void ReleaseMemoryProfileWriter()
{
    if (factory)
    {
        factory->Release();
        factory = nullptr;
    }
}
