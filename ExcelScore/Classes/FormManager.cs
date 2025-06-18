using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelScore.Classes
{

    public static class FormManager
    {
        // Stores persistent form instances
        private static Dictionary<string, Form> _formInstances = new Dictionary<string, Form>();

        /// <summary>
        /// Shows a form inside a container panel (e.g., in Form A),
        /// and reuses the instance if it's already opened.
        /// </summary>
        public static void ShowForm<T>(Panel container) where T : Form, new()
        {
            string key = typeof(T).FullName;

            // If form doesn't exist or was disposed, create a new one
            if (!_formInstances.ContainsKey(key) || _formInstances[key].IsDisposed)
            {
                T formInstance = new T();
                formInstance.TopLevel = false;
                formInstance.FormBorderStyle = FormBorderStyle.None;
                formInstance.Dock = DockStyle.Fill;

                _formInstances[key] = formInstance;
            }

            // Clear panel and add form
            container.Controls.Clear();
            container.Controls.Add(_formInstances[key]);
            _formInstances[key].Show();
        }
        public static void ShowStandaloneForm<T>(bool asDialog = false) where T : Form, new()
        {
            string key = typeof(T).FullName;

            if (!_formInstances.ContainsKey(key) || _formInstances[key].IsDisposed)
            {
                _formInstances[key] = new T();
            }

            Form instance = _formInstances[key];

            if (asDialog)
            {
                if (!instance.Visible)
                    instance.ShowDialog();
            }
            else
            {
                if (!instance.Visible)
                    instance.Show();
                else
                    instance.BringToFront();
            }
        }

        /// <summary>
        /// Hides (but doesn't dispose) a stored form of type T.
        /// </summary>
        public static void HideForm<T>() where T : Form
        {
            string key = typeof(T).FullName;

            if (_formInstances.ContainsKey(key) && !_formInstances[key].IsDisposed)
            {
                _formInstances[key].Hide();
            }
        }

        /// <summary>
        /// Close and remove the stored instance (optional cleanup).
        /// </summary>
        public static void CloseForm<T>() where T : Form
        {
            string key = typeof(T).FullName;

            if (_formInstances.ContainsKey(key))
            {
                _formInstances[key].Close();
                _formInstances.Remove(key);
            }
        }
    }

}
