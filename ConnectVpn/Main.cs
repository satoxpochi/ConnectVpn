using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ConnectVpn
{
    class Main
    {
        // -------------------------
        // 定数
        // -------------------------
        const string COMMAND_FILE_NAME = "rasphone.exe";
        const string CONNECT_ARGUMENT = "-d ";
        const string DISCONNECT_ARGUMENT = "-h ";
        const string RASPHONE_PATH = @"\Microsoft\Network\Connections\Pbk\rasphone.pbk";
        const string PREVIEW_USESR_PW_0 = "PreviewUserPw=0";
        const string PREVIEW_USESR_PW_1 = "PreviewUserPw=1";

        // -------------------------
        // 変数
        // -------------------------
        string myRasphonePath;
        string connectName;

        // -------------------------
        // Constructor
        // -------------------------
        public Main(string[] args)
        {
            // VPN設定ファイルパス
            this.myRasphonePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + RASPHONE_PATH;

            if (File.Exists(this.myRasphonePath))
            {
                // 接続名を取得
                this.connectName = GetConnectName();
                SelectExecution();
            }
            else { NoVpnSetting(); }
        }

        // -------------------------
        // 接続先名取得
        // -------------------------
        private string GetConnectName()
        {
            string name = ""; // 接続先名用変数

            // 1行ずつ読み込み
            StreamReader sr = new StreamReader(this.myRasphonePath, Encoding.UTF8);
            while (sr.Peek() > -1)
            {
                string line = sr.ReadLine();
                // 接続名があればループを抜ける
                if (name == "" && line.Contains("[") && line.Contains("]"))
                {
                    name = line.Trim('[', ']');// [ と ] を削除
                    break;
                }
            }
            sr.Close();
            return name;
        }

        // -------------------------
        // Rasphoneファイルを書き換える
        // -------------------------
        private void RewriteRasphoneFile()
        {
            //ファイルを読み込みで開く
            StreamReader sr = new StreamReader(this.myRasphonePath);
            //一時ファイルを作成する
            string tmpPath = Path.GetTempFileName();
            //一時ファイルを書き込みで開く
            StreamWriter sw = new StreamWriter(tmpPath);

            // 1行ずつ読み込み
            while (sr.Peek() > -1)
            {
                string line = sr.ReadLine();
                // PreviewUserPw=1をPreviewUserPw=0に書き換える
                if (line == PREVIEW_USESR_PW_1)
                { line = line.Replace(PREVIEW_USESR_PW_1, PREVIEW_USESR_PW_0); }
                // 一時ファイルに書き込む
                sw.WriteLine(line);
            }
            // 閉じる
            sr.Close();
            sw.Close();

            // 一時ファイルと入れ替える
            File.Copy(tmpPath, this.myRasphonePath, true);
            File.Delete(tmpPath);
        }

        // -------------------------
        // VPN接続されているかどうか
        // -------------------------
        private bool IsVpnStatusActive()
        {
            // Processオブジェクトを作成
            Process cmd = new Process();

            cmd.StartInfo.FileName = "ipconfig";
            cmd.StartInfo.UseShellExecute = false;
            cmd.StartInfo.RedirectStandardOutput = true;
            cmd.StartInfo.RedirectStandardInput = false;
            cmd.StartInfo.CreateNoWindow = true;
            cmd.Start();

            // 出力を読み取る
            string result = cmd.StandardOutput.ReadToEnd();

            // プロセス終了まで待機する
            cmd.WaitForExit();
            cmd.Close();

            bool isConnected = result.Contains(this.connectName) ? true : false;
            return isConnected;
        }

        // -------------------------
        // メッセージボックスで選択
        // -------------------------
        private void SelectExecution()
        {
            DialogResult result;

            // すでに接続されているか
            if (IsVpnStatusActive())
            {
                // メッセージボックス表示
                result = MessageBox.Show(this.connectName + "から切断しますか？"
                    , "切断確認"
                    , MessageBoxButtons.YesNo
                    , MessageBoxIcon.Question
                    , MessageBoxDefaultButton.Button2); // 「いいえ」にフォーカス

                // 「はい」なら切断
                if (result == DialogResult.Yes) { ExecuteConnect(DISCONNECT_ARGUMENT); }
                else { Environment.Exit(0); }
            }
            else
            {
                // メッセージボックス表示
                result = MessageBox.Show(this.connectName + "に接続しますか？"
                    , "接続確認"
                    , MessageBoxButtons.YesNo
                    , MessageBoxIcon.Question
                    , MessageBoxDefaultButton.Button2); // 「いいえ」にフォーカス

                // 「はい」なら接続
                if (result == DialogResult.Yes)
                {
                    RewriteRasphoneFile();
                    ExecuteConnect(CONNECT_ARGUMENT);
                }
                else { Environment.Exit(0); }
            }
        }

        // -------------------------
        // 接続or切断実行
        // -------------------------
        private void ExecuteConnect(string argument)
        {
            //Processオブジェクトを作成
            Process cmd = new Process();

            cmd.StartInfo.FileName = COMMAND_FILE_NAME;
            cmd.StartInfo.UseShellExecute = false;
            cmd.StartInfo.RedirectStandardOutput = true;
            cmd.StartInfo.RedirectStandardInput = false;
            cmd.StartInfo.CreateNoWindow = true;
            cmd.StartInfo.Arguments = $"{argument} \"{this.connectName}\"";  //コマンドラインを指定
            cmd.Start();

            //プロセス終了まで待機する
            cmd.WaitForExit();
            cmd.Close();
        }

        // -------------------------
        // VPN設定がないとき
        // -------------------------
        private void NoVpnSetting()
        { MessageBox.Show("VPN設定がありません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }
}