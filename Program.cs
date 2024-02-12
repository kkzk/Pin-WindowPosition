using System;
using System.Diagnostics;
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

        static AutomationElement sourceElement;

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

            // フォーカス変更のイベントを監視し、フック対象を見つける
            Automation.AddAutomationFocusChangedEventHandler(
                new AutomationFocusChangedEventHandler(ForcusChangedHandler)
                );

            // タスクトレイのみ表示
            mainWindow = new Form1();
            mainWindow.notifyIcon1.Text = $"{windowName}";
            Application.Run();
        }

        // ウインドウが開いた時のハンドラ
        private static void ForcusChangedHandler(object sender, AutomationEventArgs e) {
            string nameProperty = "";
            sourceElement = sender as AutomationElement;

            if (sourceElement.Current.ControlType != ControlType.Window) {
                // 親ウインドウを探す
                Condition propCondition = new PropertyCondition(
                AutomationElement.ControlTypeProperty,
                    ControlType.Window);
                TreeWalker treeWalker = new TreeWalker(propCondition);

                var new_sourceElement = treeWalker.GetParent(sourceElement);
                if (new_sourceElement != null) {
                    sourceElement = new_sourceElement;
                }
            }

            try {
                nameProperty = sourceElement.GetCurrentPropertyValue(AutomationElement.NameProperty) as string;
            }
            catch(ElementNotAvailableException) {
                Console.WriteLine("すでに無いので何もしない");
                return;
            }


            // ウインドウタイトルが一致しない場合は何もしない
            if (nameProperty != windowName) {
                Console.WriteLine("Open event not handled:{0}", nameProperty);
                return;
            }

            // （指定されている場合）モジュール名が一致しない場合は何もしない
            int processId = (int)sourceElement.GetCurrentPropertyValue(AutomationElement.ProcessIdProperty);
            Process process = Process.GetProcessById(processId);
            if (moduleName != null) {
                if (moduleName != process.MainModule.ModuleName) {
                    return;
                }
            }
            Console.WriteLine($"Open event handled:{nameProperty} / {process.MainModule.ModuleName}");

            mainWindow.notifyIcon1.Text = $"{nameProperty}(captured)";

            // レジストリにあるウインドウの位置を設定
            string retrievedRect = (string)Microsoft.Win32.Registry.GetValue(REGISTRY_KEY, windowName, null);
            if (retrievedRect != null) {
                string[] parts = retrievedRect.Split(',');
                double x = Convert.ToDouble(parts[0]);
                double y = Convert.ToDouble(parts[1]);
                double width = Convert.ToDouble(parts[2]);
                double height = Convert.ToDouble(parts[3]);
                Rect restoredRect = new Rect(x, y, width, height);
                Console.WriteLine($"Restored Rect: {restoredRect}");
                
                
                TransformPattern transformPattern = sourceElement.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                if (transformPattern != null) {
                    if (transformPattern.Current.CanMove) {
                        try {
                            transformPattern.Move(x, y);
                           Console.WriteLine("移動しました");
                        }
                        catch (InvalidOperationException) {
                            Console.WriteLine("移動できません");
                            return;
                        }
                    }
                    else {
                        Console.WriteLine("移動できないウインドウです");
                        return;
                    }
                }
            }

            try {
                // 移動したときのイベントハンドラを設定
                Automation.AddAutomationPropertyChangedEventHandler(
                    sourceElement,
                    TreeScope.Element,
                    propChangeHandler = new AutomationPropertyChangedEventHandler(OnPropertyChange),
                    AutomationElement.BoundingRectangleProperty);
            }
            catch(ElementNotAvailableException) {
                Console.WriteLine("移動したときのハンドラを設定できませんでした");
                return;
            }

            try {
                // 閉じたときのハンドラを設定
                Automation.AddAutomationEventHandler(WindowPattern.WindowClosedEvent,
                    sourceElement,
                    TreeScope.Element,
                    closeEventHandler = new AutomationEventHandler(HandleCloseEvent));
            }
            catch (ElementNotAvailableException) {
                Console.WriteLine("閉じたときのハンドラを設定できませんでした");
            }
        }

        private static void HandleCloseEvent(object sender, AutomationEventArgs e) {
            WindowClosedEventArgs eventArgs = e as WindowClosedEventArgs;
            // AutomationElement sourceElement = sender as AutomationElement;
            // string window_title = sourceElement.GetCurrentPropertyValue(AutomationElement.NameProperty) as string;
            Console.WriteLine("Close event:" + eventArgs.GetRuntimeId().ToString());
            //Automation.RemoveAllEventHandlers();
            //SetRootHandler();
        }

        private static void OnPropertyChange(object src, AutomationPropertyChangedEventArgs e) {
            AutomationElement sourceElement = src as AutomationElement;
            if (e.Property == AutomationElement.BoundingRectangleProperty) {
                Console.WriteLine("BoundingRectangleProperty:{0}", e.NewValue);

                try {
                    object boundingRectNoDefault =
                        sourceElement.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty, true);
                    if (boundingRectNoDefault == AutomationElement.NotSupported) {
                        // TODO Handle the case where you do not wish to proceed using the default value.
                    } else {
                        // レジストリにウインドウの位置を保存
                        Rect rect = (System.Windows.Rect)boundingRectNoDefault;
                        string serializedRect = $"{rect.X},{rect.Y},{rect.Width},{rect.Height}";
                        Microsoft.Win32.Registry.SetValue(REGISTRY_KEY, windowName, serializedRect);
                    }
                }
                catch {
                    return;
                }

            } else {
                // TODO: Handle other property-changed events.
            }
        }
    }
}
