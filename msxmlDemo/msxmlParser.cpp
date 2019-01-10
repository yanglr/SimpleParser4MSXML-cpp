#include <msxml6.h>   // Including MSXML latest version
#include <atlbase.h>
#include "atlstr.h"  // Including CString, CStringW and CW2A
#include <iostream>  // Including wcout method
#include <string>    // Including c_str()method, wcout
#include <direct.h>
#include "comutil.h" // Including _bstr_t
using namespace std;

const wchar_t *src = L""
L"<?xml version=\"1.0\" encoding=\"utf-16\"?>\r\n"
L"<root desc=\"Great\">\r\n"
L"  <text>Hey</text>\r\n"
L"    <layouts>\r\n"
L"    <lay index=\"15\" bold=\"true\"/>\r\n"
L"    <layoff index=\"12\"/>\r\n"
L"    <layin index=\"17\"/>\r\n"
L"  </layouts>\r\n"
L"</root>\r\n";

int main()
{
	CoInitialize(NULL); // Initialize COM

	CComPtr<IXMLDOMDocument> iXMLDoc;  // Or use CComPtr<IXMLDOMDocument2>, CComPtr<IXMLDOMDocument3>

	try
	{
		HRESULT hr = iXMLDoc.CoCreateInstance(__uuidof(DOMDocument));
		// 	iXMLDoc.CoCreateInstance(__uuidof(DOMDocument60));

		// Load the file. 
		VARIANT_BOOL bSuccess = false;

		// Load it from a url/filename...
		hr = iXMLDoc->load(CComVariant(L"./test.xml"), &bSuccess);
		// filePath = "./test.xml";
		// hr = iXMLDoc->load(CComVariant(filePath.c_str()), &bSuccess);

		// or from a BSTR...
		// iXMLDoc->loadXML(CComBSTR(src), &bSuccess);

		// Get a smart pointer (sp) to the root
		CComPtr<IXMLDOMElement> pRootElement;
		hr = iXMLDoc->get_documentElement(&pRootElement); // Root elements

		// Get Attribute of the note "root"
		CComBSTR ssDesc("desc");
		CComVariant deVal(VT_EMPTY);
		hr = pRootElement->getAttribute(ssDesc, &deVal);

		CComBSTR sstrRoot(L"root"); // sstrRoot("root");
		CComPtr<IXMLDOMNode> rootNode;
		hr = iXMLDoc->selectSingleNode(sstrRoot, &rootNode);  // Search "root"

		CComBSTR rootText;
		hr = rootNode->get_text(&rootText);
		if (SUCCEEDED(hr))
		{
			wstring bstrText(rootText);
			wcout << "Text of root: " << bstrText << endl;
		}

		CComPtr<IXMLDOMNode> descAttribute;
		hr = rootNode->selectSingleNode(CComBSTR("@desc"), &descAttribute);
		CComBSTR descVal;
		hr = descAttribute->get_text(&descVal);
		if (SUCCEEDED(hr))
		{
			wstring bstrText(descVal);
			wcout << "Desc Attribute: " << bstrText << endl;
		}

		if (!FAILED(hr))
		{
			wstring strVal;
			if (deVal.vt == VT_BSTR)
				strVal = deVal.bstrVal;

			wcout << "desc: " << strVal << endl;
		}

		CComPtr<IXMLDOMNodeList> pNodeList;
		hr = pRootElement->get_childNodes(&pNodeList); // Child node list
		long nLen;
		hr = pNodeList->get_length(&nLen);    // Child node list
		for (long i = 0; i != nLen; ++i) // Traverse
		{
			CComPtr<IXMLDOMNode> pNode;
			hr = pNodeList->get_item(i, &pNode);

			CComBSTR ssName;
			CComVariant val(VT_EMPTY);
			hr = pNode->get_nodeName(&ssName);
			if (SUCCEEDED(hr))
			{
				wstring bstrText(ssName);
				wcout << "Name of node " << (i + 1) << ": " << bstrText << endl;

				CString cstring(ssName);
				// To display a CStringW correctly, use wcout and cast cstring to (LPCTSTR), an easier way to 
				// display wide character strings.
				wcout << (LPCTSTR)cstring << endl;

				// CW2A converts the string in ccombstr to a multi-byte string in printstr, used for display output.
				CW2A printstr(ssName);
				cout << printstr << endl;
			}
		}

		/// Add(Append) node
		CComPtr<IXMLDOMDocument>& xmlDocData(iXMLDoc);
		CComPtr<IXMLDOMElement> imageElement;
		CComPtr<IXMLDOMNode> newImageNode;
		string imageType = "jpeg";
		char buffer[MAX_PATH];
		_getcwd(buffer, MAX_PATH);    //  Get Current Directory
		string path(buffer);          // Copy content of char*, generate a string
		string imagePath = path + "\\com.jpg";

		xmlDocData->createElement(CComBSTR(L"Image"), &imageElement);
		imageElement->setAttribute(CComBSTR(L"Type"), CComVariant(CComBSTR(imageType.c_str())));
		imageElement->setAttribute(CComBSTR(L"FileName"), CComVariant(CComBSTR(imagePath.c_str())));
		rootNode->appendChild(imageElement, &newImageNode);

		/// Remove "text" node under "root" node
		CComPtr<IXMLDOMNode> xmlOldNode;
		CComPtr<IXMLDOMNode> textNode;
		hr = rootNode->selectSingleNode(CComBSTR(L"text"), &textNode); // Search "text" node		
		hr = rootNode->removeChild(textNode, &xmlOldNode);

		/// Update XML
		hr = iXMLDoc->save(CComVariant("updated.xml"));
	}
	catch (char* pStrErr) {
		// Some error...
		std::cout << pStrErr << std::endl << std::endl;
	} // catch
	catch (...) {
		// Unknown error...
		std::cout << "Unknown error..." << std::endl << std::endl;
	}

	// Release() - that gets done automatically, also can manually do for each opened node or elements.
	iXMLDoc.Release();

	// Stop COM
	CoUninitialize();

	system("pause");
	return 0;
}
