using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var text = Text();
            var result = Regex.Match(text, "Date:");
        }

        private static string Text()
        {
            var text = "From: \"postmaster@hexonline.local\" <postmaster@hexonline.local>\n" +
                       "To: Vermaak Beeslaar <vermaak@vbprokureurs.co.za>\n" +
                       "Subject: Delivered: FW: Ransomware Trojan Warning\n" +
                       "Thread-Topic: FW: Ransomware Trojan Warning\n" +
                       "Thread-Index: AdBB7Bknel+6b2gXST+sdPmakpq5NgAMep8AAAALgSs=\n" +
                       "Date: Fri, 6 Feb 2015 17:04:24 +0200\n" +
                       "Message-ID: <5808efe1-df9d-4600-af64-e8c9d3cb2318@hexonline.local>\n" +
                       "References: <602fbeb5131247cb974cb957de5bf6d9@portal.baselinecloud.co.za>\n" +
                       " <FFF0CC89B43D6040BC9A4B7353D5398A42F5B5B2@EX27A.hexonline.local>\n" +
                       "In-Reply-To: <FFF0CC89B43D6040BC9A4B7353D5398A42F5B5B2@EX27A.hexonline.local>\n" +
                       "Content-Language: en-ZA\n" +
                       "X-MS-Exchange-Organization-AuthAs: Internal\n" +
                       "X-MS-Exchange-Organization-AuthMechanism: 05\n" +
                       "X-MS-Exchange-Organization-AuthSource: ex02.hexonline.local\n" +
                       "X-MS-Has-Attach:\n" +
                       "X-Auto-Response-Suppress: All\n" +
                       "X-MS-Exchange-Organization-SCL: -1\n" +
                       "X-MS-TNEF-Correlator:\n" +
                       "Content-Type: multipart/report;\n" +
                       "	boundary=\"_000_5808efe1df9d4600af64e8c9d3cb2318hexonlinelocal_\";\n" +
                       "	report-type=delivery-status\n" +
                       "MIME-Version: 1.0\n" +
                       "\n" +
                       "--_000_5808efe1df9d4600af64e8c9d3cb2318hexonlinelocal_\n" +
                       "Content-Type: multipart/alternative;\n" +
                       "	boundary=\"_002_5808efe1df9d4600af64e8c9d3cb2318hexonlinelocal_\"\n" +
                       "\n" +
                       "--_002_5808efe1df9d4600af64e8c9d3cb2318hexonlinelocal_\n" +
                       "Content-Type: text/plain; charset=\"us-ascii\"\n" +
                       "\n" +
                       "Your message has been delivered to the following recipients:\n" +
                       "\n" +
                       "Tanel Lombaard (info@vbprokureurs.co.za)<mailto:info@vbprokureurs.co.za>\n" +
                       "\n" +
                       "Daniel Viller (daniel@vbprokureurs.co.za)<mailto:daniel@vbprokureurs.co.za>\n" +
                       "\n" +
                       "Bianca Sanders (bianca@vbprokureurs.co.za)<mailto:bianca@vbprokureurs.co.za>\n" +
                       "\n" +
                       "Subject: FW: Ransomware Trojan Warning\n" +
                       "\n" +
                       "--_002_5808efe1df9d4600af64e8c9d3cb2318hexonlinelocal_\n" +
                       "Content-Type: text/html; charset=\"us-ascii\"\n" +
                       "Content-ID: <4561731E9E527B45B9C41A6743ED50C1@hexonline.local>\n" +
                       "\n" +
                       "<html>\n" +
                       "<head>\n" +
                       "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=us-ascii\">\n" +
                       "</head>\n" +
                       "<body>\n" +
                       "<p><b><font color=\"#000066\" size=\"3\" face=\"Arial\">Your message has been delivered to the following recipients:</font></b></p>\n" +
                       "<font color=\"#000000\" size=\"2\" face=\"Tahoma\">\n" +
                       "<p><a href=\"mailto:info@vbprokureurs.co.za\">Tanel Lombaard (info@vbprokureurs.co.za)</a></p>\n" +
                       "<p><a href=\"mailto:daniel@vbprokureurs.co.za\">Daniel Viller (daniel@vbprokureurs.co.za)</a></p>\n" +
                       "<p><a href=\"mailto:bianca@vbprokureurs.co.za\">Bianca Sanders (bianca@vbprokureurs.co.za)</a></p>\n" +
                       "<p>Subject: FW: Ransomware Trojan Warning</p>\n" +
                       "</font>\n" +
                       "</body>\n" +
                       "</html>\n" +
                       "\n" +
                       "--_002_5808efe1df9d4600af64e8c9d3cb2318hexonlinelocal_--\n" +
                       "\n" +
                       "--_000_5808efe1df9d4600af64e8c9d3cb2318hexonlinelocal_\n" +
                       "Content-Type: message/delivery-status\n" +
                       "\n" +
                       "Reporting-MTA: dns; hexonline.local\n" +
                       "\n" +
                       "Final-recipient: RFC822; info@vbprokureurs.co.za\n" +
                       "Action: delivered\n" +
                       "Status: 5.4.0\n" +
                       "X-Supplementary-Info: < #2.0.0>\n" +
                       "X-Display-Name: Tanel Lombaard\n" +
                       "\n" +
                       "Final-recipient: RFC822; daniel@vbprokureurs.co.za\n" +
                       "Action: delivered\n" +
                       "Status: 5.4.0\n" +
                       "X-Supplementary-Info: < #2.0.0>\n" +
                       "X-Display-Name: Daniel Viller\n" +
                       "\n" +
                       "Final-recipient: RFC822; bianca@vbprokureurs.co.za\n" +
                       "Action: delivered\n" +
                       "Status: 5.4.0\n" +
                       "X-Supplementary-Info: < #2.0.0>\n" +
                       "X-Display-Name: Bianca Sanders\n" +
                       "\n" +
                       "\n" +
                       "--_000_5808efe1df9d4600af64e8c9d3cb2318hexonlinelocal_--\n";
            return text;
        }
    }
}
