using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Threading;

namespace BLEMotionInstaller
{
    public delegate void SerialCommProcessMessageHandler(SerialCommProcess sender, string message);
    public delegate void SerialCommProcessFinishedHandler(SerialCommProcess sender, SerialCommProcessFinishedEventArgs args);
    public delegate void SerialCommProcessBLEConnectedHander(SerialCommProcess sender);
    public delegate void SeiralCommProcessMfxCommandSendedHandler(SerialCommProcess sender);

    /// <summary>
    /// シリアル通信メソッド終了イベント引数クラス
    /// </summary>
    public class SerialCommProcessFinishedEventArgs : EventArgs
    {
        public string PortName;
        List<PLEN.BLECommand> sendedCommandList;
        public byte[] ReceiveData;

        public SerialCommProcessFinishedEventArgs(string portname, List<PLEN.BLECommand> sendedList)
        {
            PortName = portname;
            sendedCommandList = sendedList;
        }
    }
    /// <summary>
    /// BLE接続状態列挙型
    /// </summary>
    public enum SerialState
    {
        /// <summary>
        /// 未接続
        /// </summary>
        NotConnected,
        /// <summary>
        /// 接続要求中
        /// </summary>
        Connecting,
        /// <summary>
        /// 接続完了
        /// </summary>
        Connected,
        /// <summary>
        /// 全データ送信完了（接続切断完了）
        /// </summary>
        SendCompleted,
        /// <summary>
        /// PLEN2でない（接続する必要なし）
        /// </summary>
        NotPLEN2
    }

    /// <summary>
    /// シリアル通信スレッド用クラス
    /// </summary>
    public abstract class SerialCommProcess
    {
        /// <summary>
        /// シリアル通信メソッドからのメッセージ受信イベント
        /// </summary>
        public abstract event SerialCommProcessMessageHandler serialCommProcessMessage;
        /// <summary>
        /// シリアル通信メソッド終了イベント
        /// </summary>
        public abstract event SerialCommProcessFinishedHandler serialCommProcessFinished;
        /// <summary>
        /// BLEドングル接続完了イベント
        /// </summary>
        public abstract event SerialCommProcessBLEConnectedHander serialCommProcessConnected;
        /// <summary>
        /// モーションデータ送信完了イベント（1モーション送信完了ごとに呼び出される）
        /// </summary>
        public abstract event SeiralCommProcessMfxCommandSendedHandler serialCommProcessCommandSended;
        /// <summary>
        /// COMポート名
        /// </summary>
        public string PortName
        {
            get { return serialPort.PortName; }
        }
        /// <summary>
        /// 本インスタンスのBLEコネクション状態
        /// </summary>
        public SerialState ConnectState
        {
            get { return connectState; }
        }
        /// <summary>
        /// 送信完了モーションファイル数
        /// </summary>
        public int sendedCommandCnt;
        /// <summary>
        /// 接続・通信完了通信相手リスト（PLEN問わず）
        /// </summary>
        static public Dictionary<Int64, SerialState> connectedDict = new Dictionary<Int64, SerialState>();
        /// <summary>
        /// モーションデータ送信間隔（単位：[ms]）
        /// </summary>
        protected readonly int DELAY_INTERVAL = 10;
        /// <summary>
        /// PLEN2キャラスティックUUID
        /// </summary>
        protected readonly byte[] PLEN2_TX_CHARACTERISTIC_UUID =
	        {
		        0xF9, 0x0E, 0x9C, 0xFE, 0x7E, 0x05, 0x44, 0xA5, 0x9D, 0x75, 0xF1, 0x36, 0x44, 0xD6, 0xF6, 0x45
	        };
        /// <summary>
        /// .NET シリアルポートコンポーネントインスタンス
        /// </summary>
        protected SerialPort serialPort;
        /// <summary>
        /// BLEドングルの通信接続状態
        /// </summary>
        protected SerialState connectState;
        /// <summary>
        /// 送信するモーションコマンドリスト
        /// </summary>
        protected List<PLEN.BLECommand> sendCommandList;
        /// <summary>
        /// 親インスタンス（自インスタンスの作成元）
        /// </summary>
        protected Form1 formSender;

        /// <summary>
        /// 通信開始（ スレッド用メソッド．直接呼びださない．）
        /// </summary>
        public abstract void start();
    }
}
