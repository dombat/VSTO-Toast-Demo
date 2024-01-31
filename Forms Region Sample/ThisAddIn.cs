using System;
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


///Some things that you might want to look at:
///If the delay is too long, Outlook will generally disable the plugin. Might need a policy to allow all the time.
///Add error handling, so that if the plugin is disabled, it doesn't crash Outlook.
///It appears in the middle of the screen, but you might want to move it to the bottom right corner.

namespace Forms_Region_Sample
{
    public partial class ThisAddIn
    {
        BackgroundWorker worker;
        ToastForm toastForm;

        const string LabelMessage = "Warning: Your Outlook Needs to be Restarted";
        const int durationMs = 10000;

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
            InitialiseForm();

            AddLabelToForm(message);

            InitialiseOpacityTimer();

            toastForm.StartDate = DateTime.UtcNow;
            System.Media.SystemSounds.Exclamation.Play();
            toastForm.ShowDialog();
            
        }

        private void InitialiseOpacityTimer()
        {
            Timer FormOpacityTimer = new Timer
            {
                Interval = 100,
            };

            FormOpacityTimer.Tick += OpacityTimer_Tick;
            FormOpacityTimer.Start();
        }

        private void InitialiseForm()
        {
            toastForm = new ToastForm
            {
                Size = new Size(400, 100),
                FormBorderStyle = FormBorderStyle.None,
                StartPosition = FormStartPosition.Manual,
                ShowInTaskbar = false,
                Left = Screen.PrimaryScreen.WorkingArea.Width - 400,
                Top = Screen.PrimaryScreen.WorkingArea.Height - 100,
                Opacity = 0
            };
        }

        private void AddLabelToForm(string message)
        {
            Label messageLabel = new Label
            {
                Text = message,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Arial", 10F, FontStyle.Bold)
            };

            toastForm.Controls.Add(messageLabel);
        }

        private void OpacityTimer_Tick(object sender, EventArgs e)
        {
            //in the first second, we want to increase the opacity
            if (DateTime.UtcNow.Subtract(toastForm.StartDate).TotalMilliseconds <= 1000)
            {
                toastForm.Opacity += 0.1;
            }

            //if the toast form is 4 seconds old, we want to reduce opacity by 10% to start hiding it
            if (DateTime.UtcNow.Subtract(toastForm.StartDate).TotalMilliseconds >= 4000)
            {
                toastForm.Opacity -= 0.1;
            }

            //if the toast form is 5 seconds old, we want to close it and dispose of resources
            if (DateTime.UtcNow.Subtract(toastForm.StartDate).TotalMilliseconds >= 5000)
            {
                ((Timer)sender).Stop();
                toastForm.Close();
                toastForm.Dispose();
                ((Timer)sender).Dispose();

                DisposeBackgrounWorker();
            }
        }

        private void DisposeBackgrounWorker()
        {
            if (worker != null)
            {
                if (worker.IsBusy)
                {
                    worker.CancelAsync(); // Wait for the worker to finish...                   
                }

                worker.Dispose();
                worker = null;
            }
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
