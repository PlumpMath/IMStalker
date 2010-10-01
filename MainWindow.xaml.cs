using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections.ObjectModel;
using System.Threading;

using MessengerAPI;
using System.Runtime.InteropServices;
using System.IO;

namespace IMStalker
{
	/// <summary>
	/// Interaction logic for Window1.xaml
	/// </summary>
	public partial class Window1 : Window
	{
		Thread stalkerThread;
		string logfile = null;

		public Window1()
		{
			InitializeComponent();
		}

		private void Log(string l, params object[] args)
		{
			string text = "[" + DateTime.Now + "] " + string.Format(l, args);
			Dispatcher.Invoke(new Func<object, int>(StalkLog.Items.Add), text);
			if (logfile != null)
			{
				File.AppendAllText(logfile, text + Environment.NewLine);
			}
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			StartStalker();
		}

		private void Identifier_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.Enter)
			{
				StartStalker();
			}
		}

		private void StartStalker()
		{
			stalkerThread = new Thread(Run);
			stalkerThread.IsBackground = true;
			stalkerThread.Start(Identifier.Text);

			Identifier.IsEnabled = false;
			StalkButton.IsEnabled = false;
		}

		private void Run(object parameter)
		{
			string id = parameter.ToString();

			try
			{
				IMessenger4 msn = new MessengerClass();
				IMessengerContact contact = (IMessengerContact)msn.GetContact(id, "");

				logfile = id + ".log";
				Log("stalking <{0}>", id);

				MISTATUS last = MISTATUS.MISTATUS_UNKNOWN;

				while (true)
				{
					MISTATUS now = contact.Status;

					if (now != last)
					{
						Log("<{0}>: {1}", contact.FriendlyName, GetStatusText(now));
						last = now;
					}

					Thread.Sleep(1000);
				}
			}
			catch (COMException ce)
			{
				switch ((uint)ce.ErrorCode)
				{
					case 0x8100030a:
						Log("User {0} not available.", id);
						break;
					default:
						Log("Error 0x{0:x} happened", ce.ErrorCode);
						break;
				}
			}
			catch (Exception e)
			{
				Log("Unknown error: {0}", e.Message);
			}

			logfile = null;
			Dispatcher.Invoke(new Action<DependencyProperty, object>(Identifier.SetValue), Control.IsEnabledProperty, true);
			Dispatcher.Invoke(new Action<DependencyProperty, object>(StalkButton.SetValue), Control.IsEnabledProperty, true);
		}

		private static string GetStatusText(MISTATUS s)
		{
			return s.ToString().Substring(9).Replace('_', ' ').ToLowerInvariant();
		}
	}
}
