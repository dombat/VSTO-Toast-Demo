using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///25th January 2024
/// Microsoft provides sample code for various technologies and services. However, it is important to note that this sample code is not supported under any Microsoft standard support program or service. 
///The sample code is provided “AS IS” without warranty of any kind. 
///Microsoft further disclaims all implied warranties, including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. 
///The entire risk arising out of the use or performance of the sample code and documentation remains with the user. 
///In no event shall Microsoft, its authors, owners of this repository, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever 
///(including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability 
///to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


namespace Forms_Region_Sample
{
    public partial class ThisAddIn
    {
        BackgroundWorker worker;

        const string LabelMessage = "Warning: Your Outlook Needs to be Restarted";
        const int durationMs = 5000;
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            worker = new BackgroundWorker()
            {
                WorkerSupportsCancellation = true
            };

           
            worker.DoWork += (obj, ex) => // This is where your time-consuming operation goes.
            {
                ShowToast(LabelMessage, durationMs);
            };

            worker.RunWorkerAsync(); // This starts the operation.
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        public void ShowToast(string message, int duration)
        {
            Form toastForm = new Form
            {
                Size = new Size(400, 100),
                FormBorderStyle = FormBorderStyle.None,
                StartPosition = FormStartPosition.CenterScreen,
                ShowInTaskbar = false
            };

            Label messageLabel = new Label
            {
                Text = message,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Arial", 10F, FontStyle.Bold)
            };

            toastForm.Controls.Add(messageLabel);

            Timer timer = new Timer
            {
                Interval = duration
            };

            timer.Tick += (s, e) =>
            {
                toastForm.Close();
                toastForm.Dispose();

                timer.Tick -= (x, y) => { }; // Unregister the event handler.
                                             // The event source (the Timer in this case) holds a reference to the object that owns the event handler (the ThisAddIn class in this case).
                                             // This can prevent the garbage collector from collecting the object, even if there are no other references to it, leading to a memory leak.
                                             // To prevent this, you should unregister the event handler when you're done with it. 
                timer.Dispose();

                if (worker.IsBusy)
                {
                    worker.CancelAsync();
                    // Wait for the worker to finish...
                }

                worker.Dispose();
                worker = null;
            };
            
            timer.Start();
            toastForm.ShowDialog();
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
