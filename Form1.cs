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


namespace BLEMotionInstaller
{
    public partial class Form1 : Form
    {
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
        private List<PLEN.MFX.BLEMfxCommand> sendMfxCommandList = new List<PLEN.MFX.BLEMfxCommand>();
        /// <summary>
        /// 送信完了PLEN台数
        /// </summary>
        private int commandSendedPLENCnt;

        public Form1()
        {
            InitializeComponent();
        }

        /***** フォームロード完了メソッド（イベント呼び出し） *****/
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
                    if (key.Contains("Bluegiga Bluetooth Low Energy"))
                    {
                        listBox1.SetSelected(listBox1.Items.Count - 1, true);
                    }
                }
            }
            catch (ManagementException ex)
            {
                portDict.Add("0", "Error " + ex.Message);
            }
        }
        /***** モーションファイル読み込みボタン投下メソッド（イベント呼び出し） *****/
        private void button3_Click(object sender, EventArgs e)
        {
            sendMfxCommandList.Clear();
            labelSendMfxCnt.Text = "0";

            using (OpenFileDialog of = new OpenFileDialog())
            {
                of.Filter = "モーションファイル（.mfx）|*.mfx";
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
                    // 選択されたファイルがモーションファイル(*.mfx)でない
                    if (System.IO.Path.GetExtension(fileName) != ".mfx")
                    {
                        textBox1.AppendText("error : モーションファイル(.mfx)を選択してください。" + System.Environment.NewLine);
                        return;
                    }

                    try
                    {
                        using (System.IO.FileStream stream = new System.IO.FileStream(fileName, System.IO.FileMode.Open))
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
                            if (bleMfx.convertMfxCommand() == false)
                            {
                                textBox1.AppendText("送信データの変換に失敗しました。" + System.Environment.NewLine);
                                return;
                            }
                            // 送信データに変換できたモーションファイルをリストに送信リストに追加
                            sendMfxCommandList.Add(bleMfx);
                            //textBox1.AppendText(bleMfx.strConvertedMfxForDisplay + System.Environment.NewLine);
                            textBox1.AppendText(string.Format("***** モーションファイルを送信データに変換しました。（{0}バイト） *****", bleMfx.strConvertedMfx.Length) + System.Environment.NewLine + System.Environment.NewLine);

                        }
                    }
                    catch (Exception ex)
                    {
                        textBox1.AppendText("error : モーションファイルの読み込みに失敗しました。" + System.Environment.NewLine);
                        textBox1.AppendText("（" + ex.Message + "）");
                        return;
                    }
                    // 送信モーションファイル数の画面表示
                    labelSendMfxCnt.Text = sendMfxCommandList.Count.ToString();
                }
            }
        }

        /***** 通信開始ボタン投下メソッド（イベント呼び出し） *****/
        private void button1_Click(object sender, EventArgs e)
        {
            // 一つ以上のアイテムがlistboxで選択されている時，処理を実行
            if (listBox1.SelectedItems.Count >= 1)
            {
                textBox1.Clear();
                textBox1.AppendText("***** 通信を開始します..... *****" + System.Environment.NewLine);

                // 送信すべきモーションファイルが一つも読み込まれていないとき
                if (sendMfxCommandList.Count <= 0)
                {
                    textBox1.AppendText("error : 送信するモーションファイルが読み込まれていません。" + System.Environment.NewLine);
                    return;
                }
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
                threadDict.Clear();
                SerialCommProcess.connectedDict.Clear();
                // 選択されているポート名を用いてシリアル通信スレッドを作成．実行．
                foreach (string portDictKey in listBox1.SelectedItems)
                {
                    string portName = portDict[portDictKey];
                    // シリアル通信スレッド用インスタンス作成
                    // イベント登録
                    portInstanceDict.Add(portName, new SerialCommProcess(portName, sendMfxCommandList, checkBox1.Checked));
                    portInstanceDict[portName].serialCommProcessMessage += new SerialCommProcessMessageHandler(serialCommEventProcessMessage);
                    portInstanceDict[portName].serialCommProcessFinished += new SerialCommProcessFinishedHandler(serialCommEventProcessFinished);
                    portInstanceDict[portName].serialCommProcessBLEConncted += new SerialCommProcessBLEConnectedHander(serialCommProcessEventBLEConncted);
                    portInstanceDict[portName].serialCommProcessMfxCommandSended += new SeiralCommProcessMfxCommandSendedHandler(serialCommProcessEventMfxCommandSended);
                    // スレッドテーブルに新規のシリアル通信スレッドを登録し，実行
                    threadDict.Add(portName, new Thread(portInstanceDict[portName].start));
                    threadDict[portName].Name = "SerialCommThread_" + portName;
                    threadDict[portName].Start();
                }
                toolStripStatusLabel2.Text = "コマンド送信完了PLEN数：0";
                commandSendedPLENCnt = 0;
                button1.Enabled = false;
                button2.Enabled = true;
            }
        }
        /***** モーションデータ送信完了メソッド（イベント呼び出し） *****/
        void serialCommProcessEventMfxCommandSended(SerialCommProcess sender)
        {
            // 進捗度を示すラベルの更新
            ThreadSafeDelegate(() => toolStripStatusLabel1Update());
        }
        /***** BLE接続完了メソッド（イベント呼び出し） *****/
        void serialCommProcessEventBLEConncted(SerialCommProcess sender)
        {
            // 進捗度を示すラベルの更新
            ThreadSafeDelegate(() => toolStripStatusLabel1Update());
        }
        void toolStripStatusLabel1Update()
        {
            // COMポート名を昇順にソートしてからラベル文字をいじる
            List<string> sortedPortInstanceDictKeys = portInstanceDict.Keys.ToList();
            sortedPortInstanceDictKeys.Sort();
            toolStripStatusLabel1.Text = "";
            foreach (string key in sortedPortInstanceDictKeys)
            {
                // PLENと接続が完了してるポートのみラベル
                if (portInstanceDict[key].BLEConnectState == BLEState.Connected)

                    toolStripStatusLabel1.Text += "[ " + portInstanceDict[key].PortName + "   " + portInstanceDict[key].sendedMfxCommandCnt.ToString() + " / " + sendMfxCommandList.Count + " ]    ";
            }
        }

        /***** 全モーションコマンド送信完了メソッド（イベント呼び出し） *****/
        void serialCommEventProcessFinished(object sender, SerialCommProcessFinishedEventArgs args)
        {
            // コマンド送信完了PLEN数更新．画面表示．
            commandSendedPLENCnt++;
            ThreadSafeDelegate(delegate
            {
                textBox1.AppendText(" ***** [" + args.PortName + "] Finished *****" + System.Environment.NewLine);
                toolStripStatusLabel2.Text = "コマンド送信完了PLEN数：" + commandSendedPLENCnt.ToString();
            });

            // 通知元のスレッドを終了させ，テーブルから削除（ただし自動継続モードでないとき）
            if (checkBox1.Checked == false)
            {

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
                        textBox1.AppendText(System.Environment.NewLine + "***** すべてのモーションデータの送信が完了しました。*****" + System.Environment.NewLine + System.Environment.NewLine);
                    }
                });
            }
        }

        /***** シリアル通信スレッドからのメッセージ受信メソッド（イベント呼び出し） *****/
        private void serialCommEventProcessMessage(SerialCommProcess sender, string message)
        {
            if (textBox1.IsDisposed == false && textBox1.Disposing == false)
                // 受信したメッセージを画面表示
                ThreadSafeDelegate(delegate { textBox1.AppendText("Comm Thread[" + sender.PortName + "] >> " + message + System.Environment.NewLine); });
        }


        /***** メソッドデリゲート呼び出しメソッド *****/
        // ※ListBoxやTextboxなど，Formに関するオブジェクトはメインスレッド以外からの操作ができない．
        // 　本メソッドはメインスレッド上で動かす必要があるメソッドをメインスレッドに委託する操作を行う．
        public void ThreadSafeDelegate(MethodInvoker method)
        {
            if (InvokeRequired)
                BeginInvoke(method);
            else
                method.Invoke();
        }

        /***** 全通信停止ボタン投下メソッド（イベント飛び出し） *****/
        private void button2_Click(object sender, EventArgs e)
        {
            foreach (string key in threadDict.Keys)
            {
                threadDict[key].Abort();
            }
            threadDict.Clear();
            MessageBox.Show("全通信を終了しました。", "通信を終了しました", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            textBox1.AppendText("***** All Communication Stopped *****" + System.Environment.NewLine);

            button1.Enabled = true;
            button2.Enabled = false;
        }
        /****** フォームが閉じられるときに呼ばれるメソッド（イベント呼び出し） *****/
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // すべての通信スレッドに停止命令を送信
            foreach (string key in threadDict.Keys)
            {
                threadDict[key].Abort();
            }
        }


    }
}
