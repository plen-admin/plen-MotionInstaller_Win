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
        NotConnected,
        Connecting,
        ScanServices,
        ScanCharacteristics,
        Connected,
        SendCompleted,
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
