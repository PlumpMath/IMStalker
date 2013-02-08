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

using SKYPE4COMLib;
using System.Runtime.InteropServices;
using System.IO;
using System.Text.RegularExpressions;

namespace IMStalker
{
	/// <summary>
	/// Interaction logic for Window1.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		Thread stalkerThread;
		string logfile = null;

		public MainWindow()
		{
			InitializeComponent();

			SetPreviousLogs();
		}

		private void SetPreviousLogs()
		{
			var logs = from f in Directory.GetFiles(".", "*.log")
					   orderby File.GetLastWriteTime(f) descending
					   select System.IO.Path.GetFileNameWithoutExtension(f);

			foreach (var log in logs)
			{
				Identifier.Items.Add(log);
			}
		}

		private void SetNewLog(string id)
		{
			Identifier.Items.Remove(id);
			Identifier.Items.Insert(0, id);
			Identifier.Text = id;

		}

		private void Log(string l, params object[] args)
		{
			string text = "[" + DateTime.Now.ToString("s") + "] " + string.Format(l, args);
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

		private void StopButton_Click(object sender, RoutedEventArgs e)
		{
			StopStalker();
		}

		private void Window_Closed(object sender, EventArgs e)
		{
			StopStalker();
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
			StalkButton.Visibility = Visibility.Collapsed;
			StopButton.Visibility = Visibility.Visible;
		}

		private void StopStalker()
		{
			if (stalkerThread != null && stalkerThread.IsAlive)
			{
				stalkerThread.Abort();
				stalkerThread = null;
			}
		}

		private void Run(object parameter)
		{
			string name = string.Format("<{0}>", parameter.ToString());

			try
			{
				ISkype skype = new Skype();
				IUser user = (IUser)skype.get_User(parameter.ToString());
				name = string.Format("{0} <{1}>", user.FullName, user.Handle);

				logfile = user.Handle + ".log";

				Dispatcher.Invoke(new Action<string>(SetNewLog), user.Handle);

				Log("stalking {0}", name);

				TOnlineStatus last = TOnlineStatus.olsUnknown;

				while (true)
				{
					TOnlineStatus now = user.OnlineStatus;

					if (now != last)
					{
						Log("{0}: {1}", name, GetStatusText(now));
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
						Log("User {0} not available.", name);
						break;
					default:
						Log("Error 0x{0:x} happened", ce.ErrorCode);
						break;
				}
			}
			catch (ThreadAbortException)
			{
				Log("stopping {0}", name);
			}
			catch (Exception e)
			{
				Log("Unknown error: {0}", e.Message);
			}
			finally
			{
				logfile = null;
				Dispatcher.Invoke(new Action<DependencyProperty, object>(Identifier.SetValue), Control.IsEnabledProperty, true);
				Dispatcher.Invoke(new Action<DependencyProperty, object>(StalkButton.SetValue), Control.VisibilityProperty, Visibility.Visible);
				Dispatcher.Invoke(new Action<DependencyProperty, object>(StopButton.SetValue), Control.VisibilityProperty, Visibility.Collapsed);
			}
		}

		private static string GetStatusText(TOnlineStatus s)
		{
			return Regex.Replace(s.ToString().Substring(3), "([a-z])([A-Z])", "$1 $2").ToLowerInvariant();
		}
	}
}
