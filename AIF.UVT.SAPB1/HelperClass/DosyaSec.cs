using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIF.UVT.SAPB1.HelperClass
{
    class DosyaSec
    {

        System.Windows.Forms.OpenFileDialog openFileDialog1 = new System.Windows.Forms.OpenFileDialog();


        public static string Result;

        public DosyaSec()
        {

            System.Diagnostics.Trace.WriteLine("WorkerRole1 entry point called", "Information");

            System.Threading.Thread thread = new System.Threading.Thread(new System.Threading.ThreadStart(SetFileName));
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        public void SetFileName()
        {

            System.Windows.Forms.DialogResult result = openFileDialog1.ShowDialog();
            // Show the dialog.

            if (result == System.Windows.Forms.DialogResult.OK) // Test result.
            {
                Result = openFileDialog1.FileName;
            }


        }
        public string GetFileName()
        {
            return Result;
        }
    }
}
