using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

using System.Windows.Forms;

namespace JurisUtilityBase
{
    public class ExecutionClass : IDisposable
    {

        public void DragonsBreath(string company, List<Bill> bList, string textBox, string path, bool processExpense)
        {
            NewWrapper rr = null;
            try
            {
                rr = new NewWrapper();
            }
            catch (Exception exception)
            {
                var errorMessage = exception.Message;
                if (rr != null && rr.WrapperException != null)
                {
                    errorMessage += " " + rr.WrapperException.Message;
                }
                MessageBox.Show("Unable to view bill image...null wrapper: " + errorMessage,
            Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            rr.logOnAndDoWork(company, bList, textBox, path, processExpense);
            //rr.IDisposable_Dispose();
        }

        private bool _disposedValue;

        // Instantiate a SafeHandle instance.
        private SafeHandle _safeHandle = new SafeFileHandle(IntPtr.Zero, true);

        // Public implementation of Dispose pattern callable by consumers.
        public void Dispose() => Dispose(true);

        // Protected implementation of Dispose pattern.
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    _safeHandle.Dispose();
                }

                _disposedValue = true;
            }
        }



    }
}
