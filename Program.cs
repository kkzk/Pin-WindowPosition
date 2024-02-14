using System;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;

namespace Pin_WindowPosition {

    internal static class Program {
        private static AutomationPropertyChangedEventHandler propChangeHandler;
        private static AutomationEventHandler openEventHandler;
        private static AutomationEventHandler closeEventHandler;

        private static string windowName;
        private static string moduleName;
        private static string REGISTRY_KEY = "HKEY_CURRENT_USER\\Software\\kkzk\\Pin-WindowPosition";
        private static int[] windowRuntimeId = { };

        static Form1 mainWindow;

        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main(string[] args) {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (args.Length == 0) {
                windowName = "簡易再生";
                moduleName = null;
            } else if (args.Length == 1) {
                windowName = args[0];
                moduleName = null;
            } else if (args.Length == 2) {
                windowName = args[0];
                moduleName = args[1];
            }

            // フォーカス変更のイベントを監視する
            Automation.AddAutomationFocusChangedEventHandler(
                new AutomationFocusChangedEventHandler(ForcusChangedHandler)
                );

            // タスクトレイのみ表示
            mainWindow = new Form1();
            mainWindow.notifyIcon1.Text = $"{windowName}";
            Application.Run();
        }

        // フォーカス変更時のハンドラ
        private static void ForcusChangedHandler(object sender, AutomationEventArgs e) {
            AutomationElement windowElement;
            string nameProperty = "";
            AutomationElement sourceElement = sender as AutomationElement;

            windowElement = sourceElement;
            if (sourceElement.Current.ControlType != ControlType.Window) {
                // 自分がWindowでない場合は親ウインドウを探す
                Condition propCondition = new PropertyCondition(
                    AutomationElement.ControlTypeProperty,
                    ControlType.Window);
                TreeWalker treeWalker = new TreeWalker(propCondition);

                var parent = treeWalker.GetParent(sourceElement);
                if (parent != null) {
                    windowElement = parent;
                }
            }

            // ウインドウタイトルの取得
            try {
                nameProperty = windowElement.GetCurrentPropertyValue(AutomationElement.NameProperty) as string;
            }
            catch(ElementNotAvailableException) {
                Trace.WriteLine("すでに無いので何もしない");
                return;
            }

            // 実行モジュール名を取得
            int processId = (int)windowElement.GetCurrentPropertyValue(AutomationElement.ProcessIdProperty);
            Process process = Process.GetProcessById(processId);

            // ウインドウタイトルが一致しない場合は何もしない
            if (nameProperty != windowName) {
                Trace.WriteLine($"Open event not handled: \"{nameProperty}\"({process.MainModule.ModuleName})");
                return;
            }

            // （指定されている場合）モジュール名が一致しない場合は何もしない
            if (moduleName != null) {
                if (moduleName != process.MainModule.ModuleName) {
                    return;
                }
            }
            Trace.WriteLine($"Open event handled: \"{windowName}\"({process.MainModule.ModuleName})");

            int[] runtimeId = windowElement.GetRuntimeId();
            if (runtimeId.SequenceEqual(windowRuntimeId)) {
                // キャプチャ済みなので何もしない
                Trace.WriteLine("キャプチャ済みなので何もしない");
                return;
            }
            windowRuntimeId = runtimeId;
            mainWindow.notifyIcon1.Text = $"{windowName}(captured)";

            // レジストリにあるウインドウの位置を設定
            string retrievedRect = (string)Microsoft.Win32.Registry.GetValue(REGISTRY_KEY, windowName, null);
            if (retrievedRect != null) {
                string[] parts = retrievedRect.Split(',');
                double x = Convert.ToDouble(parts[0]);
                double y = Convert.ToDouble(parts[1]);
                double width = Convert.ToDouble(parts[2]);
                double height = Convert.ToDouble(parts[3]);
                Rect restoredRect = new Rect(x, y, width, height);
                Trace.WriteLine($"Restored Rect: {restoredRect}");
                
                TransformPattern transformPattern = windowElement.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                if (transformPattern != null) {
                    if (transformPattern.Current.CanMove) {
                        try {
                            transformPattern.Move(x, y);
                            Trace.WriteLine("移動しました");
                        }
                        catch (InvalidOperationException) {
                            Trace.WriteLine("移動できません");
                            return;
                        }
                    }
                    else {
                        Trace.WriteLine("移動できないウインドウです");
                        return;
                    }
                }
            }

            // 移動したときのイベントハンドラを設定
            try {
                Automation.AddAutomationPropertyChangedEventHandler(
                    windowElement,
                    TreeScope.Element,
                    propChangeHandler = new AutomationPropertyChangedEventHandler(OnPropertyChange),
                    AutomationElement.BoundingRectangleProperty);
            }
            catch(ElementNotAvailableException) {
                Trace.WriteLine("移動したときのハンドラを設定できませんでした");
                return;
            }

            // 閉じたときのイベントハンドラを設定
            try {
                Automation.AddAutomationEventHandler(
                    WindowPattern.WindowClosedEvent,
                    windowElement,
                    TreeScope.Element,
                    new AutomationEventHandler(OnWindowClosed));
            }
            catch(ElementNotAvailableException) {
                Trace.WriteLine("閉じたときのハンドラを設定できませんでした");
                return;
            }

        }


        private static void OnPropertyChange(object src, AutomationPropertyChangedEventArgs e) {
            AutomationElement sourceElement = src as AutomationElement;
            if (e.Property == AutomationElement.BoundingRectangleProperty) {
                Trace.WriteLine($"BoundingRectangleProperty:{e.NewValue}" );

                try {
                    object boundingRectNoDefault =
                        sourceElement.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty, true);
                    if (boundingRectNoDefault == AutomationElement.NotSupported) {
                        // TODO Handle the case where you do not wish to proceed using the default value.
                    } else {
                        // レジストリにウインドウの位置を保存
                        Rect rect = (Rect)boundingRectNoDefault;
                        string serializedRect = $"{rect.X},{rect.Y},{rect.Width},{rect.Height}";
                        Microsoft.Win32.Registry.SetValue(REGISTRY_KEY, windowName, serializedRect);
                    }
                }
                catch {
                    return;
                }

            }
        }

        private static void OnWindowClosed(object src, AutomationEventArgs e) {
            WindowClosedEventArgs windowClosedEventArgs = e as WindowClosedEventArgs;
            int[] runtimeId = windowClosedEventArgs.GetRuntimeId();
            if (runtimeId.SequenceEqual(windowRuntimeId)) {
                windowRuntimeId = new int[] { };
                mainWindow.notifyIcon1.Text = $"{windowName}";
                Trace.WriteLine("ウインドウが閉じました");
            }
            else {
                Trace.WriteLine("間違ったウインドウを閉じています");
            }
        }
    }
}
