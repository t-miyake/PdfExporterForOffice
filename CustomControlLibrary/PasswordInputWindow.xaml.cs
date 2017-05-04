using System.Windows;
namespace CustomControlLibrary
{
    /// <summary>
    /// PasswordInputWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class PasswordInputWindow : Window
    {
        public string Password = string.Empty;

        public PasswordInputWindow()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(PasswordBox1.Password)|| string.IsNullOrEmpty(PasswordBox2.Password))
            {
                MessageBox.Show("Enter the password.");
            }

            if (PasswordBox1.Password == PasswordBox2.Password)
            {
                Password = PasswordBox1.Password;
                DialogResult = true;
            }
            else
            {
                MessageBox.Show("Don't much password.");
            }
        }

        private void CanselButton_OnClickButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
