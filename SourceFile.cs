using System;
using System.Collections.Generic;
using System.Data;
using UiPath.Core;
using UiPath.Core.Activities.Storage;
using UiPath.Orchestrator.Client.Models;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Office.Interop.Excel;
using System.Runtime.Versioning;
using System.Security;
namespace UiPath_Excel_Activities_OpenLibrary_BK
{
    public static class SourceFile
    {
        public static Microsoft.Office.Interop.Excel.Workbook getWorkbookFromFilePath(string excelFP)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = null;
	        if(System.Diagnostics.Process.GetProcessesByName("EXCEL").Length>0)
            {
                xlApp = (Microsoft.Office.Interop.Excel.Application)SlimShady.GetActiveObject("Excel.Application");
            }
            else 
            {
                throw new Exception("Get Or Init XL App Exception=Excel.Application could not be initialized! its null");
            }
            
            Microsoft.Office.Interop.Excel.Workbook rezWB = null;
            foreach(Microsoft.Office.Interop.Excel.Workbook wb in xlApp.Workbooks)
            {
                if(wb.FullName == excelFP)
                {
                    rezWB = wb;
                    break;
                }
            }
            if(rezWB == null)
            {
                throw new Exception("This workbook does not exist in the open Excel application!-->"+excelFP);
            }
            return rezWB;
        }
    }
    public static class SlimShady
    {
        internal const String OLEAUT32 = "oleaut32.dll";
        internal const String OLE32 = "ole32.dll";
    
        [System.Security.SecurityCritical]  // auto-generated_required
        public static Object GetActiveObject(String progID)
        {
            Object obj = null;
            Guid clsid;
    
            // Call CLSIDFromProgIDEx first then fall back on CLSIDFromProgID if
            // CLSIDFromProgIDEx doesn't exist.
            try
            {
                CLSIDFromProgIDEx(progID, out clsid);
            }
            //            catch
            catch (Exception)
            {
                CLSIDFromProgID(progID, out clsid);
            }
    
            GetActiveObject(ref clsid, IntPtr.Zero, out obj);
            return obj;
        }
    
        //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
        [DllImport(OLE32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);
    
        //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
        [DllImport(OLE32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);
    
        //[DllImport(Microsoft.Win32.Win32Native.OLEAUT32, PreserveSig = false)]
        [DllImport(OLEAUT32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out Object ppunk);
    
    }
}