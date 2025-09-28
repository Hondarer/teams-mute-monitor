using System.Diagnostics;
using System.Windows.Automation;

namespace MuteMonitor
{
    class Program
    {
        private static bool? lastMicState;
        private static bool debugMode = false;

        private static readonly Condition ButtonCondition =
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
        private static readonly Condition MicEnabledLabelsCondition = new OrCondition(
            new PropertyCondition(AutomationElement.NameProperty, "mute: マイク"),
            new PropertyCondition(AutomationElement.NameProperty, "マイクのミュート")
        );
        private static readonly Condition CheckMicEnabledCondition = new AndCondition(
            ButtonCondition,
            MicEnabledLabelsCondition
        );

        private const string KEIKOCHAN_ADDRESS = "192.168.66.104";

        private static readonly HttpClient httpClient = new HttpClient();
        private static readonly Uri MicEnabledUri = new Uri($"http://{KEIKOCHAN_ADDRESS}/api/control?acop=x1xxxxxx");
        private static readonly Uri MicDisabledUri = new Uri($"http://{KEIKOCHAN_ADDRESS}/api/control?acop=x0xxxxxx");

        static async Task Main(string[] args)
        {
            Console.WriteLine("Teams マイク状態モニター");
            Console.WriteLine("=======================================");

            while (true)
            {
                await Task.Delay(250);
                await CheckMicState();
            }
        }

        private static Process? FindTeamsProcess()
        {
            var processes = Process.GetProcessesByName("ms-teams");
            if (processes.Length > 0) return processes[0];
            return null;
        }

        private static async Task<bool?> GetMicState()
        {
            try
            {
                var proc = FindTeamsProcess();

                if (proc == null)
                {
                    if (debugMode) Console.WriteLine("[DEBUG] Teams プロセスなし → マイク無効");
                    return false;
                }

                if (proc.HasExited || proc.MainWindowHandle == IntPtr.Zero)
                {
                    if (debugMode) Console.WriteLine("[DEBUG] Teams ウィンドウ取得不可 → マイク無効");
                    return false;
                }

                var root = AutomationElement.FromHandle(proc.MainWindowHandle);
                if (root == null)
                {
                    if (debugMode) Console.WriteLine("[DEBUG] AutomationElement 取得失敗 → マイク無効");
                    return false;
                }

                // ルールに基づく接頭辞
                var disabledPrefixes = new[] { "unmute: マイク", "マイクのミュートを解除" };
                var enabledPrefixes = new[] { "mute: マイク", "マイクのミュート" };

                // マイク有効チェック
                var enabledElem = FindElementByNamePrefix(root, CheckMicEnabledCondition);
                if (enabledElem != null)
                {
                    if (debugMode) Console.WriteLine($"[DEBUG] 有効判定: {enabledElem.Current.Name}");
                    return true;
                }
            }
            catch (Exception ex)
            {
                if (debugMode) Console.WriteLine($"[DEBUG] 取得エラー: {ex.Message} → マイク無効");
                return false;
            }

            // 判定材料がない場合はマイク無効
            if (debugMode) Console.WriteLine("[DEBUG] 判定材料なし → マイク無効");
            return false;
        }

        private static AutomationElement? FindElementByNamePrefix(AutomationElement root, Condition condition)
        {
            try
            {
                AutomationElement hit = root.FindFirst(TreeScope.Descendants, condition);
                if (hit != null) return hit;
            }
            catch
            {
                return null;
            }
            return null;
        }

        private static async Task NotifyKeikochanAsync(bool isEnabled)
        {
            var uri = isEnabled ? MicEnabledUri : MicDisabledUri;
            try
            {
                var resp = await httpClient.GetAsync(uri);
                if (debugMode)
                {
                    Console.WriteLine($"[DEBUG] 警子ちゃん通知 {uri} → {(int)resp.StatusCode} {resp.ReasonPhrase}");
                }
            }
            catch (Exception ex)
            {
                if (debugMode) Console.WriteLine($"[DEBUG] 警子ちゃん通知エラー: {ex.Message}");
            }
        }

        private static async Task CheckMicState()
        {
            var current = await GetMicState();

            if (debugMode && current.HasValue)
            {
                Console.WriteLine($"[DEBUG] 現在: {current}, 前回: {lastMicState}");
            }

            if (current.HasValue && lastMicState != current)
            {
                lastMicState = current.Value;
                var timestamp = DateTime.Now.ToString("HH:mm:ss");
                Console.WriteLine($"{timestamp} {GetMicStatusText(lastMicState.Value)}");
                await NotifyKeikochanAsync(lastMicState.Value);
            }
        }

        private static string GetMicStatusText(bool isEnabled)
        {
            return isEnabled ? "🟢 マイク有効" : "🔴 マイク無効";
        }
    }
}
