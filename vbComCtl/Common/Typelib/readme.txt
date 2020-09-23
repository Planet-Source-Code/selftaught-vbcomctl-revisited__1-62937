Sure would be nice to combine these type libraries into one.

We just need to meet either one of these goals:

1.	Make mktyplib.exe to align structures such as PRINTDLG on word alignment instead of dword.  Also make it recognize Safearray(void) or similar construct to allow "byref arg() as any".

2.	Make midl.exe compile the interface definitions with the correct uuids for iunknown, idispatch, etc.