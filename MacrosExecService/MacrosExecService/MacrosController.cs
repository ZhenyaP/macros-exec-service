using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Web.Http;
using MacrosExecService.Entities;
using MacrosExecService.Helpers;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;

namespace MacrosExecService
{
    public class MacrosController : ApiController
    {
        // GET api/Macros 
        public IEnumerable<string> Get()
        {
            return new[] { "Hello", "World" };
        }



        // POST api/Macros/GetExcelFileAfterMacroExec/{macroDataListJsonStr} 
        //[Route("api/Macros/{macroDataListJsonStr}")]
        [HttpPost]
        public async Task<IHttpActionResult> GetExcelFileAfterMacrosExec(string macroDataListJsonStr)
        {
            Excel.Application excel = null;
            Excel._Workbook workbook = null;
            VBA.VBComponent module = null;
            bool saveChanges = false;
            string resultedTempFileName = null;
            byte[] resultedBuffer = null;
            try
            {
                if (!Request.Content.IsMimeMultipartContent())
                    throw new Exception();
                var provider = new MultipartMemoryStreamProvider();
                await Request.Content.ReadAsMultipartAsync(provider);
                var file = provider.Contents.First();
                var filename = file.Headers.ContentDisposition.FileName.Trim('\"');
                var buffer = await file.ReadAsByteArrayAsync();
                var tempFileName = Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, filename);
                resultedTempFileName = Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, $"{Path.GetFileNameWithoutExtension(filename)}_transformed");
                //var tempFileName = Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "TempFiles", filename);
                //resultedTempFileName = Path.Combine(AppDomain.CurrentDomain.SetupInformation.ApplicationBase, "TempFiles", filename + "_transformed");
                //File.WriteAllBytes(tempFileName, buffer);
                await FileHelper.WriteAllBytesAsync(tempFileName, buffer);
                excel = new Excel.Application { Visible = false };
                workbook = excel.Workbooks.Open(tempFileName, false, true, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, false, false, Type.Missing, false, true, Type.Missing);
                var macroDataList = JsonConvert.DeserializeObject<List<MacroData>>(macroDataListJsonStr);
                foreach (var macroData in macroDataList)
                {
                    module = workbook.VBProject.VBComponents.Add(VBA.vbext_ComponentType.vbext_ct_StdModule);
                    module.CodeModule.AddFromString(macroData.MacroCode);
                    // Run the named VBA Sub that we just added.  In our sample, we named the Sub FormatSheet

                    workbook.Application.Run(macroData.VbaSubName, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value);
                }
                // Let loose control of the Excel instance

                excel.Visible = false;
                excel.UserControl = false;

                // Set a flag saying that all is well and it is ok to save our changes to a file.
                saveChanges = true;
                //  Save the file to disk
                //workbook.SaveAs(fileNameToSave, Excel.XlFileFormat.xlWorkbookNormal,
                //        null, null, false, false, Excel.XlSaveAsAccessMode.xlShared,
                //        false, false, null, null, null);

                workbook.SaveAs(resultedTempFileName, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled, Missing.Value,
    Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
    Excel.XlSaveConflictResolution.xlLocalSessionChanges, true,
    Missing.Value, Missing.Value, Missing.Value);

            }
            catch (Exception e)
            {
                Console.WriteLine($"Error in WriteMacro: {e.Message}, Stack Trace: {e.StackTrace}");
            }
            finally
            {
                try
                {
                    // Repeat xl.Visible and xl.UserControl releases just to be sure
                    // we didn't error out ahead of time.

                    if (excel != null)
                    {
                        excel.Visible = false;
                        excel.UserControl = false;
                    }
                    // Close the document and avoid user prompts to save if our
                    // method failed.
                    workbook?.Close(saveChanges, null, null);
                    excel?.Workbooks.Close();
                }
                catch
                {
                    // ignored
                }

                // Gracefully exit out and destroy all COM objects to avoid hanging instances
                // of Excel.exe whether our method failed or not.
                excel?.Quit();
                if (module != null) Marshal.ReleaseComObject(module);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excel != null) Marshal.ReleaseComObject(excel);   //This is used to kill the EXCEL.exe process

                GC.Collect();
                if (!string.IsNullOrEmpty(resultedTempFileName))
                {
                    resultedTempFileName += ".xlsm";
                    resultedBuffer = File.ReadAllBytes(resultedTempFileName);
                }
            }
            return Ok(resultedBuffer);
        }
    }
}
