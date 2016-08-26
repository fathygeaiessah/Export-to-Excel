using EnvDTE;
using EnvDTE80;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using vbide = Microsoft.Vbe.Interop;


namespace Export_to_Excel
{
    public partial class frmMacros : Form
    {
        public frmMacros(Excel.Application FN)
        {
            InitializeComponent();

            xl = FN;
        }

        Form1 fr = new Form1();

        Excel.Application xl;
        Excel.Workbook wb;


        private void btnCreateMacro_Click(object sender, EventArgs e)
        {


            foreach (vbide.VBComponent item in proj.VBComponents)
            {
                vbide.VBComponent Md = item as vbide.VBComponent;
                vbide.CodeModule MdCode = Md.CodeModule;
                if (MdCode != null)
                {
                    if (MdCode.get_ProcOfLine(1, out projType) == txtMacroName.Text)
                    {

                        MdCode.DeleteLines(1, Md.CodeModule.get_ProcCountLines(txtMacroName.Text, projType));

                    }

                }

           }
               
                    vbide.VBComponent Mdn;
                    Mdn = proj.VBComponents.Add(vbide.vbext_ComponentType.vbext_ct_StdModule);

                    string sCode = "";
                    sCode = "public Sub " + txtMacroName.Text + "() \n";
                    int LC = txtVBACode.Lines.Count();
                    for (int i = 0; i < LC; i++)
                    {
                        sCode += "" + txtVBACode.Lines[i].ToString() + "\n";
                    }

                    Mdn.CodeModule.AddFromString(sCode);

                    cmbMacrosNames.Items.Clear();
            foreach (var item2 in proj.VBComponents)
            {
                vbide.VBComponent vbComponent = item2 as vbide.VBComponent;
                if (vbComponent != null)
                {
                    string componentName = vbComponent.Name;
                    vbide.CodeModule comCode = vbComponent.CodeModule;
                    int comCodeLines = comCode.CountOfLines;
                    int line = 1;
                    while (line <= comCodeLines)
                    {
                        string proceName = comCode.get_ProcOfLine(line, out projType);
                        if (line == comCode.get_ProcStartLine(proceName, projType))
                        {
                            if (proceName != null)
                            {
                                cmbMacrosNames.Items.Add(proceName);
                            }
                        }
                        line = line + 1;
                    }
                }
            }

                    MessageBox.Show(txtMacroName.Text + " Macro has been saved successfully");
                
            
           
        }

        //private void RunMacro(object oApp, object[] oRunArgs)
        //{
        //    oApp.GetType().InvokeMember("Run",
        //        System.Reflection.BindingFlags.Default |
        //        System.Reflection.BindingFlags.InvokeMethod,
        //        null, oApp, oRunArgs);
        //}

        static void NAR(object o) // MSDN article suggestion. 
        {
            try
            {
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
                {
                }
            }
            catch
            {
                // swallow 
            }
            finally
            {
                o = null;
            }
        }
        private void btnRunMacro_Click(object sender, EventArgs e)
        {
            try
            {
                
                xl.Run(txtMacroName.Text);
                wb.CheckCompatibility = false;
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        vbide.VBProject proj;

        Microsoft.Vbe.Interop.vbext_ProcKind projType;
        private void frmMacros_Load(object sender, EventArgs e)
        {

            wb = xl.ActiveWorkbook;
            proj = wb.VBProject;
            var projName = proj.Name;
            projType = Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc;
            cmbMacrosNames.Items.Clear();
            foreach (var item in proj.VBComponents)
            {
                vbide.VBComponent vbComponent = item as vbide.VBComponent;
                if (vbComponent != null)
                {
                    string componentName = vbComponent.Name;
                    vbide.CodeModule comCode = vbComponent.CodeModule;
                    int comCodeLines = comCode.CountOfLines;
                    int line = 1;
                    while (line <= comCodeLines)
                    {
                        string proceName = comCode.get_ProcOfLine(line, out projType);
                        if (line == comCode.get_ProcStartLine(proceName, projType))
                        {
                            if (proceName != null)
                            {
                                cmbMacrosNames.Items.Add(proceName);
                            }
                        }
                        line = line + 1;
                    }
                }
            }


        }

        private void cmbMacrosNames_SelectedIndexChanged(object sender, EventArgs e)
        {

            // proj = wb.VBProject;
            // projType = Microsoft.Vbe.Interop.vbext_ProcKind.vbext_pk_Proc;
            txtMacroName.Text = cmbMacrosNames.Text;
            foreach (var item in proj.VBComponents)
            {
                vbide.VBComponent vbComponent = item as vbide.VBComponent;
                vbide.CodeModule comCode = vbComponent.CodeModule;
                for (int i = 1; i <= comCode.CountOfLines; i++)
                {
                    if (comCode.ProcOfLine[i, out projType]==txtMacroName.Text)
                    {
                        int proceCount = comCode.ProcCountLines[txtMacroName.Text, projType] <= 1 ? 0 : comCode.ProcCountLines[txtMacroName.Text, projType];
                        if (proceCount > 2)
                        {

                            string proceCode = comCode.Lines[comCode.ProcStartLine[txtMacroName.Text, projType], proceCount];
                            if (proceCode != null)
                            {
                                txtVBACode.Text = proceCode;
                            }
                        }
                    }
                }
                
            }
        }




        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception occured while releasing object " + ex);
            }
            finally
            {
                GC.Collect();
            }
        }


        private void frmMacros_Deactivate(object sender, EventArgs e)
        {



        }
    }
}

   

    //public class MessageFilter : IoleMessageFilter
    //{
    //    public static void Register()
    //    {
    //        IoleMessageFilter newFilter = new MessageFilter();
    //        IoleMessageFilter oldFilter = null;
    //        CoRegisterMessageFilter(newFilter, out oldFilter);
    //    }

    //    public static void Revoke()
    //    {
    //        IoleMessageFilter oldFilter = null;
    //        CoRegisterMessageFilter(null, out oldFilter);
    //    }

    //    int IoleMessageFilter.handleIncomingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
    //    {
    //        return 0;
    //    }
    //    int IoleMessageFilter.RetryRejectCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    //    {
    //        if (dwRejectType == 2)
    //        {
    //            return 99;
    //        }
    //        return -1;
    //    }
    //    int IoleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwCount, int dwPendingType)
    //    {
    //        return 2;
    //    }

    //    [DllImport("Ole32.dll")]
    //    private static extern int CoRegisterMessageFilter(IoleMessageFilter newFilter, out IoleMessageFilter oldFilter);
    //}
    //    [ComImport(), Guid("00000016-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    // interface IoleMessageFilter
    //{
    //    [PreserveSig]
    //    int handleIncomingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    //    [PreserveSig]
    //    int RetryRejectCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    //    [PreserveSig]
    //    int MessagePending(IntPtr hTaskCallee, int dwCount, int dwPendingType);
    //}
    


