using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using System.Management;
using System.Runtime.Serialization.Json;


namespace BLEMotionInstaller
{
    public partial class Form1 : Form
    {  
        /// <summary>
        /// PLEN2接続要求ポート名リスト
        /// PLEN2と接続したいBLEドングルは，このリストにポート名を追加すると，PLEN2接続用スレッドで接続処理を行ってくれる．
        /// </summary>
        public List<string> bleConnectingRequestPortList = new List<string>();

        /// <summary>
        /// 通信メソッドハッシュテーブル（キー：COMポート名)
        /// </summary>
        private Dictionary<string, Thread> threadDict = new Dictionary<string, Thread>();
        /// <summary>
        /// シリアルポートインスタンスハッシュテーブル（キー：COMポート名）
        /// </summary>
        private Dictionary<string, SerialCommProcess> portInstanceDict = new Dictionary<string, SerialCommProcess>();
        /// <summary>
        /// PC上に存在するすべてのCOMポートのハッシュテーブル
        /// （キー：「ポート名(COM4とか) - Caption(BlueGiga Bluetooth Low Enegeryとか)」，値：ポート名）
        /// </summary>
        private Dictionary<string, string> portDict = new Dictionary<string, string>();
        /// <summary>
        /// 送信するモーションファイル（コマンド済）のリスト
        /// </summary>
        private List<PLEN.BLECommand> sendCommandList = new List<PLEN.BLECommand>();
        /// <summary>
        /// 送信完了PLEN台数
        /// </summary>
        private int commandSendedPLENCnt;
        /// <summary>
        /// PLEN2接続スレッド
        /// </summary>
        private Thread bleConnectingThread;

        private const string BLE_COMPORT_NAME = "Bluegiga Bluetooth Low Energy";
        private const string USB_PLEN_COMPORT_NAME = "Arduino Micro";


        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// フォームロード完了メソッド（イベント呼び出し）
        /// </summary>
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                // PCに接続されているCOMポートの一覧を取得．portDictに格納

                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_SerialPort");
                foreach (ManagementObject queryObj in searcher.Get())
                {
                    string key = String.Format("{0} - {1}", queryObj["DeviceID"], queryObj["Caption"]);
                    portDict.Add(key, String.Format("{0}", queryObj["DeviceID"]));
                    listBox1.Items.Add(key);
                    // BlueGigaのドングルに関してはあらかじめlistBoxを選択状態にしておく
                    if (key.Contains(BLE_COMPORT_NAME))
                    {
                        listBox1.SetSelected(listBox1.Items.Count - 1, true);
                    }
                    cmbBoxMode.SelectedIndex = 0;
                }
            }
            catch (ManagementException ex)
            {
                portDict.Add("0", "Error " + ex.Message);
                
            }

            // プログラム起動時の引数がある場合，引数が指定するJSONファイルを読み込み
            // Note...args[0]：本プログラムのパス（関係なし），args[1]：JSONファイルのパス，args[2]：JSONファイル名
            string[] args = Environment.GetCommandLineArgs();
            if (args.Length > 2)
            {
                // JSONファイル読み込み．送信データに変換．（変換後送信ボタンをフォーカスさせる）
                string jsonPath = Uri.UnescapeDataString(args[1].Replace('+', ' '));
                string name = Uri.UnescapeDataString(args[2].Replace('+', ' '));
                using (System.IO.FileStream stream = new System.IO.FileStream(jsonPath, System.IO.FileMode.Open))
                {
                    readJsonFile(stream, name);
                    button1.Focus();
                }
         }
        }
        /// <summary>
        /// PLEN2接続スレッドメソッド（bleConnectingThread上で動作）
        /// </summary>
        private void bleConnectingThreadFunc()
        {
            try
            {
                while (true)
                {
                    // PLEN2に接続したいBLEドングルはbleConnectingRequestPortListにポート名を追加している
                    if (bleConnectingRequestPortList.Count > 0)
                    {
                        // ログ用テキストボックス更新
                        this.Invoke((MethodInvoker)delegate
                        {
                            textBox1.AppendText("Connection Thread     >> [" + bleConnectingRequestPortList[0] + "] Start PLEN2-Connection." + System.Environment.NewLine);
                        });
                        // 接続要求をしているCOMポートに対して，PLEN2接続処理を行う
                        ((SerialCommProcessBLE)portInstanceDict[bleConnectingRequestPortList[0]]).bleConnect();
                        // 接続
                        bleConnectingRequestPortList.RemoveAt(0);
                    }
                    else
                        Thread.Sleep(1);
                 }

            }
            catch (Exception) { }
        }

        /// <summary>
        /// モーションファイル読み込みボタン投下メソッド（イベント呼び出し）
        /// </summary>
        private void button3_Click(object sender, EventArgs e)
        {
            sendCommandList.Clear();
            labelSendCmdCnt.Text = "0";

            using (OpenFileDialog of = new OpenFileDialog())
            {
                of.Filter = "モーションファイル|*.mfx;*.json";
                of.Multiselect = true;  //複数のファイルを選択できるようにする
                of.ShowDialog();

                foreach (string fileName in of.FileNames)
                {
                    // 選択されたファイルが存在しない
                    if (!System.IO.File.Exists(fileName))
                    {
                        textBox1.AppendText("error : モーションファイルの検索に失敗しました。" + System.Environment.NewLine);
                        return;

                    }
                    // ファイルの拡張子に応じてモーションファイルを読み込む
                    using (System.IO.FileStream stream = new System.IO.FileStream(fileName, System.IO.FileMode.Open))
                    {
                        if (System.IO.Path.GetExtension(fileName) == ".mfx")
                            readMfxFile(stream, fileName);
                        else if (System.IO.Path.GetExtension(fileName) == ".json")
                            readJsonFile(stream, fileName);
                        else
                            textBox1.AppendText("error : モーションファイル(.mfx)を選択してください。" + System.Environment.NewLine);

                    }
                }
            }
        }
        /// <summary>
        /// モーションデータ（MFXファイル（XML記述））読み込みメソッド
        /// </summary>
        /// <param name="stream">ストリーム</param>
        /// <param name="fileName">ファイル名</param>
        private void readMfxFile(System.IO.Stream stream, string fileName)
        {
            // モーションファイル（XML形式)
            PLEN.MFX.XmlMfxModel tagMfx;
            /*----- XML読み出し -----*/
            try
            {
                System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(PLEN.MFX.XmlMfxModel));
                tagMfx = (PLEN.MFX.XmlMfxModel)serializer.Deserialize(stream);
            }
            catch (Exception ex)
            {
                textBox1.AppendText("error : XML解析に失敗しました。選択したモーションファイルが破損している恐れがあります。" + System.Environment.NewLine);
                textBox1.AppendText("（" + ex.Message + "）");

                return;
            }

            /*----- XML→送信データ(string)への変換 -----*/
            textBox1.AppendText("【" + fileName + "】モーションファイルを送信データとして変換します..." + System.Environment.NewLine);

            PLEN.MFX.BLEMfxCommand bleMfx = new PLEN.MFX.BLEMfxCommand(tagMfx);
            if (bleMfx.convertCommand() == false)
            {
                textBox1.AppendText("送信データの変換に失敗しました。" + System.Environment.NewLine);
                return;
            }
            // 送信データに変換できたモーションファイルをリストに送信リストに追加
            sendCommandList.Add(bleMfx);
            //textBox1.AppendText(bleMfx.strConvertedMfxForDisplay + System.Environment.NewLine);
            textBox1.AppendText(string.Format("***** モーションファイルを送信データに変換しました。（{0}バイト） *****", bleMfx.convertedStr.Length) + System.Environment.NewLine + System.Environment.NewLine);

            // 送信モーションファイル数の画面表示
            labelSendCmdCnt.Text = sendCommandList.Count.ToString();    
        }
        /// <summary>
        /// モーションデータ（JSON形式）読み込みメソッド
        /// </summary>
        /// <param name="stream">ストリーム</param>
        /// <param name="fileName">ファイル名</param>
        private void readJsonFile(System.IO.Stream stream, string fileName)
        {
            // モーションファイル（JSON形式)
            PLEN.JSON.Main jsonData;
            /*----- JSON読み出し -----*/
            try
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(PLEN.JSON.Main));
                jsonData = (PLEN.JSON.Main)serializer.ReadObject(stream);
            }
            catch (Exception ex)
            {
                textBox1.AppendText("error : JSONファイルの解析に失敗しました。選択したモーションファイルが破損している恐れがあります。" + System.Environment.NewLine);
                textBox1.AppendText("（" + ex.Message + "）");

                return;
            }

            /*----- JSON→送信データ(string)への変換 -----*/
            textBox1.AppendText("【" + fileName + "】モーションファイルを送信データとして変換します..." + System.Environment.NewLine);

            PLEN.JSON.BLEJsonCommand bleJson = new PLEN.JSON.BLEJsonCommand(jsonData);
            if (bleJson.convertCommand() == false)
            {
                textBox1.AppendText("送信データの変換に失敗しました。" + System.Environment.NewLine);
                return;
            }
            // 送信データに変換できたモーションファイルをリストに送信リストに追加
            sendCommandList.Add(bleJson);
//            textBox1.AppendText(bleJson.convertedStrForDisplay);
            textBox1.AppendText(string.Format("***** モーションファイルを送信データに変換しました。（{0}バイト） *****", bleJson.convertedStr.Length) + System.Environment.NewLine + System.Environment.NewLine);
            
            // 送信モーションファイル数の画面表示
            labelSendCmdCnt.Text = sendCommandList.Count.ToString();
        }


        /// <summary>
        /// 通信開始ボタン投下メソッド（イベント呼び出し）
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            // 一つ以上のアイテムがlistboxで選択されている時，処理を実行
            if (listBox1.SelectedItems.Count >= 1)
            {
                textBox1.Clear();
                textBox1.AppendText("***** 通信を開始します..... *****" + System.Environment.NewLine);

                // 送信すべきモーションファイルが一つも読み込まれていないとき
                if (sendCommandList.Count <= 0)
                {
                    textBox1.AppendText("error : 送信するモーションファイルが読み込まれていません。" + System.Environment.NewLine);
                    return;
                }

                try
                {
                    // 今現在存在するシリアルポートのインスタンスをすべてリセット（削除）
                    portInstanceDict.Clear();

                    // 今現在動作している通信用スレッドをすべて停止．スレッドテーブルをリセット
                    if (threadDict.Count > 0)
                    {
                        foreach (string key in threadDict.Keys)
                        {
                            threadDict[key].Abort();
                        }
                    }
                    // 各リスト，テーブルをリセット
                    threadDict.Clear();
                    SerialCommProcess.connectedDict.Clear();
                    bleConnectingRequestPortList.Clear();

                    // 選択されているポート名を用いてシリアル通信スレッドを作成．実行．
                    foreach (string portDictKey in listBox1.SelectedItems)
                    {
                        string portName = portDict[portDictKey];

                        // SelectedIndex：0..BLE接続，1..USB接続
                        if (cmbBoxMode.SelectedIndex == 0)
                            portInstanceDict.Add(portName, new SerialCommProcessBLE(this, portName, sendCommandList));
                        else
                            portInstanceDict.Add(portName, new SerialCommProcessUSB(this, portName, sendCommandList));

                    // イベント登録
                        portInstanceDict[portName].serialCommProcessMessage += new SerialCommProcessMessageHandler(serialCommEventProcessMessage);
                        portInstanceDict[portName].serialCommProcessFinished += new SerialCommProcessFinishedHandler(serialCommEventProcessFinished);
                        portInstanceDict[portName].serialCommProcessConnected += new SerialCommProcessBLEConnectedHander(serialCommProcessEventBLEConncted);
                        portInstanceDict[portName].serialCommProcessCommandSended += new SeiralCommProcessMfxCommandSendedHandler(serialCommProcessEventMfxCommandSended);

                        // スレッドテーブルに新規のシリアル通信スレッドを登録し，実行
                        threadDict.Add(portName, new Thread(portInstanceDict[portName].start));
                        threadDict[portName].Name = "SerialCommThread_" + portName;
                        threadDict[portName].Start();
                    }
                    // PLEN2接続スレッドを起動
                    bleConnectingThread = new Thread(bleConnectingThreadFunc);
                    bleConnectingThread.Name = "BLEConnectingThread";
                    bleConnectingThread.Start();

                    toolStripStatusLabel2.Text = "コマンド送信完了PLEN数：0";
                    commandSendedPLENCnt = 0;
                    button1.Enabled = false;
                    button2.Enabled = true;
                }
                // 例外が発生→すべてのスレッドを停止させる
                catch (Exception ex)
                {
                    textBox1.AppendText("エラーが発生しました． " + System.Environment.NewLine + "massage  [ " + ex.Message + " ]" + System.Environment.NewLine);
                    foreach (string key in threadDict.Keys)
                    {
                        threadDict[key].Abort();
                    }
                    if (bleConnectingThread != null)
                        bleConnectingThread.Abort();

                    button1.Enabled = true;
                    button2.Enabled = false;
                }
            }
        }
        /// <summary>
        /// モーションデータ送信完了メソッド（イベント呼び出し）
        /// </summary>
        void serialCommProcessEventMfxCommandSended(SerialCommProcess sender)
        {
            // 進捗度を示すラベルの更新
            ThreadSafeDelegate(() => toolStripStatusLabel1Update());
        }
        
        /// <summary>
        /// BLE接続完了メソッド（イベント呼び出し）
        /// </summary>
        void serialCommProcessEventBLEConncted(SerialCommProcess sender)
        {
            // 進捗度を示すラベルの更新
            ThreadSafeDelegate(() => toolStripStatusLabel1Update());
        }
        /// <summary>
        /// 進捗表示ラベルの更新メソッド
        /// </summary>
        void toolStripStatusLabel1Update()
        {
            // COMポート名を昇順にソートしてからラベル文字をいじる
            List<string> sortedPortInstanceDictKeys = portInstanceDict.Keys.ToList();
            sortedPortInstanceDictKeys.Sort();
            toolStripStatusLabel1.Text = "";
            foreach (string key in sortedPortInstanceDictKeys)
            {
                // PLENと接続が完了してるポートのみラベル
                if (portInstanceDict[key].ConnectState == SerialState.Connected || portInstanceDict[key].ConnectState == SerialState.SendCompleted)
                    toolStripStatusLabel1.Text += "[ " + portInstanceDict[key].PortName + "   " + portInstanceDict[key].sendedCommandCnt.ToString() + " / " + sendCommandList.Count + " ]    ";
            }
        }

        /// <summary>
        /// 全モーションコマンド送信完了メソッド（イベント呼び出し）
        /// </summary>
        void serialCommEventProcessFinished(object sender, SerialCommProcessFinishedEventArgs args)
        {
            // コマンド送信完了PLEN数更新．画面表示．
            commandSendedPLENCnt++;
            ThreadSafeDelegate(delegate
            {
                //textBox1.AppendText(" ***** [" + args.PortName + "] Finished *****" + System.Environment.NewLine);
                toolStripStatusLabel2.Text = "コマンド送信完了PLEN数：" + commandSendedPLENCnt.ToString();
            });

            // 通知元のスレッドを終了させ，テーブルから削除（ただし自動継続モードでないとき）
            ThreadSafeDelegate(delegate
            {
                textBox1.AppendText(" ***** [" + args.PortName + "] Finished *****" + System.Environment.NewLine);
                if (threadDict.Keys.Contains(args.PortName))
                {
                    threadDict[args.PortName].Abort();
                    threadDict.Remove(args.PortName);
                }
                // すべてのスレッドが終了した時
                if (threadDict.Values.Count <= 0)
                {
                    button1.Enabled = true;
                    button2.Enabled = false;
                    if (bleConnectingThread != null)                    
                    {
                        bleConnectingThread.Abort();
                        bleConnectingThread.Join();
                    }
                    textBox1.AppendText(System.Environment.NewLine + "***** すべてのモーションデータの送信が完了しました。*****" + System.Environment.NewLine + System.Environment.NewLine);
                }
            });
        }

        /// <summary>
        /// シリアル通信スレッドからのメッセージ受信メソッド（イベント呼び出し）
        /// </summary>
        private void serialCommEventProcessMessage(SerialCommProcess sender, string message)
        {
            if (textBox1.IsDisposed == false && textBox1.Disposing == false)
                // 受信したメッセージを画面表示
                ThreadSafeDelegate(delegate { textBox1.AppendText("Comm Thread[" + sender.PortName + "] >> " + message + System.Environment.NewLine); });
        }


        /// <summary>
        /// メソッドデリゲート呼び出しメソッド
        /// Note...ListBoxやTextboxなど，Formに関するオブジェクトはメインスレッド以外からの操作ができない．
        // 　      本メソッドはメインスレッド上で動かす必要があるメソッドをメインスレッドに委託する操作を行う．
        /// </summary>
        /// <param name="method"></param>
        public void ThreadSafeDelegate(MethodInvoker method)
        {
            if (InvokeRequired)
                BeginInvoke(method);
            else
                method.Invoke();
        }

        /*****  *****/
        /// <summary>
        /// 全通信停止ボタン投下メソッド（イベント飛び出し）
        /// </summary>
        private void button2_Click(object sender, EventArgs e)
        {
            foreach (string key in threadDict.Keys)
            {
                threadDict[key].Abort();
                threadDict[key].Join(1500);
            }

            bleConnectingThread.Abort();
            bleConnectingThread.Join();

            threadDict.Clear();
            MessageBox.Show("全通信を終了しました。", "通信を終了しました", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            textBox1.AppendText("***** All Communication Stopped *****" + System.Environment.NewLine);

            button1.Enabled = true;
            button2.Enabled = false;
        }

        /// <summary>
        /// フォームが閉じられるときに呼ばれるメソッド（イベント呼び出し）
        /// </summary>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // すべての通信スレッドに停止命令を送信
            foreach (string key in threadDict.Keys)
            {
                if (threadDict[key] != null)
                {
                    threadDict[key].Abort();
                    threadDict[key].Join(1500);
                }
            }
            if (bleConnectingThread != null)
            {
                bleConnectingThread.Abort();
                bleConnectingThread.Join(1500);
            }
        }

        private void cmbBoxMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox1.SelectedIndex = -1;
            if (cmbBoxMode.SelectedIndex == 0)
            {
                foreach (Object listItemObj in listBox1.Items)
                {
                    string listItemStr = (string)listItemObj;
                    if (listItemStr.Contains(BLE_COMPORT_NAME))
                    {
                        listBox1.SelectedItems.Add(listItemObj);
                        break;
                    }
                }
            }
            else
            {
                foreach (Object listItemObj in listBox1.Items)
                {
                    string listItemStr = (string)listItemObj;
                    if (listItemStr.Contains(USB_PLEN_COMPORT_NAME))
                    {
                        listBox1.SelectedIndex = listBox1.FindString(listItemStr);
                        break;
                    }
                }
            }
        }

    }
}
