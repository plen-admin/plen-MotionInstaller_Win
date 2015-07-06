﻿using System;
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
    public enum BLEState
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
    public class SerialCommProcess
    {
        /// <summary>
        /// シリアル通信メソッドからのメッセージ受信イベント
        /// </summary>
        public event SerialCommProcessMessageHandler serialCommProcessMessage;
        /// <summary>
        /// シリアル通信メソッド終了イベント
        /// </summary>
        public event SerialCommProcessFinishedHandler serialCommProcessFinished;
        /// <summary>
        /// BLEドングル接続完了イベント
        /// </summary>
        public event SerialCommProcessBLEConnectedHander serialCommProcessBLEConncted;
        /// <summary>
        /// モーションデータ送信完了イベント（1モーション送信完了ごとに呼び出される）
        /// </summary>
        public event SeiralCommProcessMfxCommandSendedHandler serialCommProcessMfxCommandSended;
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
        public BLEState BLEConnectState
        {
            get { return bleConnectState; }
        }
        /// <summary>
        /// 接続・通信完了通信相手リスト（PLEN問わず）
        /// </summary>
        static public Dictionary<Int64, BLEState> connectedDict = new Dictionary<Int64, BLEState>();
        /// <summary>
        /// 送信完了モーションファイル数
        /// </summary>
        public int sendedMfxCommandCnt;
        /// <summary>
        /// モーションデータ送信間隔（単位：[ms]）
        /// </summary>
        private readonly int DELAY_INTERVAL = 10;
        /// <summary>
        /// PLEN2キャラスティックUUID
        /// </summary>
        private readonly byte[] PLEN2_TX_CHARACTERISTIC_UUID =
	        {
		        0xF9, 0x0E, 0x9C, 0xFE, 0x7E, 0x05, 0x44, 0xA5, 0x9D, 0x75, 0xF1, 0x36, 0x44, 0xD6, 0xF6, 0x45
	        };
        /// <summary>
        /// .NET シリアルポートコンポーネントインスタンス
        /// </summary>
        private SerialPort serialPort;
        /// <summary>
        /// BLEドングルの通信接続状態
        /// </summary>
        private BLEState bleConnectState;
        /// <summary>
        /// BLE接続先クライアントキー
        /// </summary>
        private Int64 connectedBLEKey = 0;
        /// <summary>
        /// 送信するモーションコマンドリスト
        /// </summary>
        private List<PLEN.BLECommand> sendCommandList;
        /// <summary>
        /// BgLibインスタンス
        /// </summary>
        private Bluegiga.BGLib bgLib;
        /// <summary>
        /// シリアルポート受信データ格納配列（生データ）
        /// </summary>
        private byte[] receiveRawData;
        /// <summary>
        /// 
        /// </summary>
        private bool isAttributeWrited;
        /// <summary>
        /// 親インスタンス（自インスタンスの作成元）
        /// </summary>
        private Form1 formSender;


        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="sender">インスタンス作成元Form</param>
        /// <param name="portName">シリアルポート名</param>
        /// <param name="mfxCommandList">送信モーションファイルリスト</param>
        /// <param name="isContinuationMode">自動継続モードにするかしないか</param>
        public SerialCommProcess(object sender, string portName, List<PLEN.BLECommand> commandList)
        {
            formSender = (Form1)sender;

            // シリアルポート初期化
            serialPort = new SerialPort();
            serialPort.PortName = portName;
            serialPort.Handshake = System.IO.Ports.Handshake.RequestToSend;
            serialPort.BaudRate = 115200;
            serialPort.DataBits = 8;
            serialPort.StopBits = System.IO.Ports.StopBits.One;
            serialPort.Parity = System.IO.Ports.Parity.None;
            serialPort.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(SerialDataReceived);
            // 送信モーションコマンドリストを登録
            sendCommandList = commandList;
            // bgLib初期化．イベント登録．
            bgLib = new Bluegiga.BGLib();
            bgLib.BLEEventGAPScanResponse += new Bluegiga.BLE.Events.GAP.ScanResponseEventHandler(bgLib_BLEEventGAPScanResponse);
            bgLib.BLEEventConnectionStatus += new Bluegiga.BLE.Events.Connection.StatusEventHandler(bgLib_BLEEventConnectionStatus);
            bgLib.BLEEventATTClientProcedureCompleted += new Bluegiga.BLE.Events.ATTClient.ProcedureCompletedEventHandler(bgLib_BLEEventATTClientProcedureCompleted);

            try
            {
                serialPort.Open();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// AtrributeWrite-データ送信完了メソッド（イベント呼び出し）
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">Args</param>
        void bgLib_BLEEventATTClientProcedureCompleted(object sender, Bluegiga.BLE.Events.ATTClient.ProcedureCompletedEventArgs e)
        {
            //serialCommProcessMessage(this, "ProcedureCompleted Result : " + e.result);

            // e.result == 0 ⇒データ送信正常完了
            if (e.result == 0)
                isAttributeWrited = true;
        }
        /// <summary>
        /// PLEN2接続メソッド（注：同時に複数のドングルがPLEN2に接続しないよう，本メソッドはシングルタスクで動作させることを推奨！）
        /// （現在Form1のbleConnectingThreadのみが接続処理を行うようにしている．）
        /// </summary>
        public void bleConnect()
        {
            /*----- 半二重通信ここから -----*/
            bgLib.SendCommand(serialPort, bgLib.BLECommandGAPEndProcedure());
            while (bgLib.IsBusy() == true)
                Thread.Sleep(1);

            Thread.Sleep(10);
            bgLib.SendCommand(serialPort, bgLib.BLECommandConnectionDisconnect(0));
            while (bgLib.IsBusy() == true)
                Thread.Sleep(1);

            // PLEN2との接続を試みる
            Thread.Sleep(10);
            serialCommProcessMessage(this, "PLEN2 searching...");
            bleConnectState = BLEState.NotConnected;
            bgLib.SendCommand(serialPort, bgLib.BLECommandGAPDiscover(1));
            while (bleConnectState != BLEState.Connected)
                Thread.Sleep(1);
        }

        /// <summary>
        /// 通信開始（ スレッド用メソッド．直接呼びださない．）
        /// </summary>
        public void start()
        {
            // ※親スレッドが本スレッドを破棄しようとするとThreadAbortExceptionが発生
            // 　自身が破棄される前に必ずシリアルポートを閉じる
            try
            {
             halfDuplexComm();
            }
            catch (Exception) {  }
            finally
            {
                // クライアントと接続を切断
                if (serialPort.IsOpen == true)
                {
                    int i = 0;
                    const int TIMEOUT = 50;

                    bgLib.SendCommand(serialPort, bgLib.BLECommandConnectionDisconnect(0));
                    for (i = 0; i < TIMEOUT && bgLib.IsBusy() == true; i++)
                    {
                        Thread.Sleep(1);
                    }
                    if (i >= TIMEOUT)
                    {
                        serialCommProcessMessage(this, "PLEN2との接続が解除できませんでした．BLEドングルを抜き差ししてください．");
                    }
                }
                // 今回接続したクライアントのキーをリストから削除
                if (connectedDict.Keys.Contains(connectedBLEKey))
                    connectedDict.Remove(connectedBLEKey);

                serialPort.Close();
            }
        }
        
        /***** GAPScanResponse検知メソッド（bgLibイベント呼び出し) *****/
        private void bgLib_BLEEventGAPScanResponse(object sender, Bluegiga.BLE.Events.GAP.ScanResponseEventArgs e)
        {
            // データパケットの長さが25以上であれば、UUIDが乗っていないかチェックする
            if (e.data.Length > 25)
            {
                byte[] data_buf = new byte[16];

                for (int i = 0; i < data_buf.Length; i++)
                    data_buf[i] = e.data[i + 9];

                // データパケットにUUIDが乗っていた場合
                if (Object.ReferenceEquals(data_buf, PLEN2_TX_CHARACTERISTIC_UUID) && bleConnectState == BLEState.NotConnected)
                {
                    bleConnectState = BLEState.Connecting;
                    // PLEN2からのアドバタイズなので、接続を試みる
                    bgLib.SendCommand(serialPort, bgLib.BLECommandGAPConnectDirect(e.sender, 0, 60, 76, 100, 0));
                }
            }
            else
            {
                // 本来の接続手順を実装する。(具体的には以下の通りです。)
                // 1. ble_cmd_gap_connect_direct()
                // 2. ble_cmd_attclient_find_information()
                // 3. ble_evt_attclient_find_information_found()を処理し、UUIDを比較
                //     a. UUIDが一致する場合は、そのキャラクタリスティックハンドルを取得 → 接続完了
                //     b. 全てのキャラクタリスティックについてUUIDが一致しない場合は、4.以降の処理へ
                // 4. MACアドレスを除外リストに追加した後、ble_cmd_connection_disconnect()
                // 5. 再度ble_cmd_gap_discove()
                // 6. 1.へ戻る。ただし、除外リストとMACアドレスを比較し、該当するものには接続をしない。

                // CAUTION!: 以下は横着した実装。本来は上記の手順を踏むべき。
                // PLEN2からのアドバタイズなので、接続を試みる
                Int64 key = 0;
                // キー作成
                for (int index = 0; index < 6; index++)
                    key += (Int64)e.sender[index] << (index * 8);

                // 取得したキーに対してまだ接続していない場合，接続を試みる
                // ※connectedDictは他スレッドからの参照もありうるので排他制御
                if (bleConnectState == BLEState.NotConnected)
                {
                    lock(connectedDict)
                    {
                        // 取得したキーに対して，すでに接続を試みたならばスルー
                        if (connectedDict.ContainsKey(key) == true && connectedDict[key] != BLEState.NotConnected)
                        {
                            return;
                        }
                        // 取得したキーをリストに追加，状態を接続要求中へ
                        if (connectedDict.ContainsKey(key) == true)
                            connectedDict.Add(key, BLEState.Connecting);
                        else
                            connectedDict[key] = BLEState.Connecting;
                        // 自分の接続状態を更新
                        bleConnectState = BLEState.Connecting;
                        // 取得したキーに対して接続を試みる(接続完了後イベント発生)
                        bgLib.SendCommand(serialPort, bgLib.BLECommandGAPConnectDirect(e.sender, 0, 60, 76, 100, 0));
                        
                    }
                }
                

            }
        }
        /// <summary>
        /// BLEクライアント接続完了メソッド（イベント呼び出し）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bgLib_BLEEventConnectionStatus(object sender, Bluegiga.BLE.Events.Connection.StatusEventArgs e)
        {
            // PLENと接続完了
            if ((e.flags & 0x01) != 0)
            {
                // アドレスリストを更新（接続中→接続完了へ）
                // ※connectedDictは他スレッドからも操作されるので排他制御
                lock (connectedDict)
                {
                    // アドレス作成
                    Int64 key = 0;
                    for (int index = 0; index < 6; index++)
                        key += (Int64)e.address[index] << (index * 8);
                    // 状態更新（接続完了）
                    connectedDict[key] = BLEState.Connected;
                    connectedBLEKey = key;
                }
                serialCommProcessMessage(this, "Connected [" + connectedBLEKey.ToString() + "]");
                bleConnectState = BLEState.Connected;
            }
            // 再度接続を試みる
            else
                bgLib.SendCommand(serialPort, bgLib.BLECommandGAPDiscover(1));
        }


        /// <summary>
        /// 半二重通信メソッド
        /// </summary>
        private void halfDuplexComm()
        {

            serialCommProcessMessage(this, "HalfDuplexCommunication Started");
            sendedMfxCommandCnt = 0;    // カウントリセット

            bleConnectState = BLEState.NotConnected;
            // PLEN2との接続処理をPLEN2接続用スレッド上（シングルタスク）上で行うため，接続要求リストにセット
            // ※PLEN2接続スレッドはbleConnectingRequestPortListにアイテムが追加されると自動的にテーブルの1番目から接続処理を行う
            if(!formSender.bleConnectingRequestPortList.Contains(PortName))
                formSender.bleConnectingRequestPortList.Add(PortName);
            serialCommProcessMessage(this, "BLE Connecting Thread Waiting...");
            while (bleConnectState != BLEState.Connected)
                Thread.Sleep(1);

            /*-- ここからPLEN2と接続中 --*/
            serialCommProcessMessage(this, "PLEN2 Connected");
            serialCommProcessBLEConncted(this);

            try
            {
                foreach (PLEN.BLECommand sendCommand in sendCommandList)
                {
                    // 送信データを文字列からbyte配列に変換
                    byte[] mfxCommandArray = System.Text.Encoding.ASCII.GetBytes(sendCommand.convertedStr);
                    serialCommProcessMessage(this, "【" + sendCommand.Name + "】 is sending...");

                    /*---- ここからモーションデータ送信 -----*/
                    /*-- header --*/
                    isAttributeWrited = false;
                    byte[] test = System.Text.Encoding.ASCII.GetBytes("#IN");
                    bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, 31, System.Text.Encoding.ASCII.GetBytes("#IN")));
                    while (isAttributeWrited == false)
                        Thread.Sleep(1);

                    Thread.Sleep(DELAY_INTERVAL);

                    isAttributeWrited = false;
                    bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, 31, mfxCommandArray.Take(20).ToArray()));
                    while (isAttributeWrited == false)
                        Thread.Sleep(1);

                    Thread.Sleep(DELAY_INTERVAL);

                    isAttributeWrited = false;
                    bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, 31, mfxCommandArray.Skip(20).Take(10).ToArray()));
                    while (isAttributeWrited == false)
                        Thread.Sleep(1);

                    Thread.Sleep(DELAY_INTERVAL);
                    serialCommProcessMessage(this, "header written.");

                    Thread.Sleep(50);
                    /*-- frame --*/
                    for (int index = 0; index < (mfxCommandArray.Length - 30) / 100; index++)
                    {
                        isAttributeWrited = false;
                        bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, 31, mfxCommandArray.Skip(30 + index * 100).Take(20).ToArray()));
                        while (isAttributeWrited == false)
                            Thread.Sleep(1);

                        Thread.Sleep(DELAY_INTERVAL);

                        isAttributeWrited = false;
                        bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, 31, mfxCommandArray.Skip(50 + index * 100).Take(20).ToArray()));
                        while (isAttributeWrited == false)
                            Thread.Sleep(1);

                        Thread.Sleep(DELAY_INTERVAL);

                        isAttributeWrited = false;
                        bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, 31, mfxCommandArray.Skip(70 + index * 100).Take(20).ToArray()));
                        while (isAttributeWrited == false)
                            Thread.Sleep(1);
                        Thread.Sleep(DELAY_INTERVAL);
                        isAttributeWrited = false;
                        bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, 31, mfxCommandArray.Skip(90 + index * 100).Take(20).ToArray()));
                        while (isAttributeWrited == false)
                            Thread.Sleep(1);

                        Thread.Sleep(DELAY_INTERVAL);

                        isAttributeWrited = false;
                        bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, 31, mfxCommandArray.Skip(110 + index * 100).Take(20).ToArray()));
                        while (isAttributeWrited == false)
                            Thread.Sleep(1);

                        Thread.Sleep(DELAY_INTERVAL);

                        serialCommProcessMessage(this, "frame written. [" + (index + 1).ToString() + "/" + ((mfxCommandArray.Length - 30) / 100).ToString() + "]");
                        Thread.Sleep(50);
                    }
                    serialCommProcessMessage(this, "【" + sendCommand.Name + "】send Complete. ...");
                    sendedMfxCommandCnt++;
                    serialCommProcessMfxCommandSended(this);
                    Thread.Sleep(500);
                }
                // 終了イベントを発生
                serialCommProcessMessage(this, "Communication Finished");
                serialCommProcessFinished(this, new SerialCommProcessFinishedEventArgs(serialPort.PortName, sendCommandList));
                connectedDict[connectedBLEKey] = BLEState.SendCompleted;
            }
            catch (Exception e)
            {
                serialCommProcessMessage(this, e.Message);
            }
            finally
            {
                /*----- 半二重通信ここまで -----*/
                bgLib.SendCommand(serialPort, bgLib.BLECommandConnectionDisconnect(0));
                while (bgLib.IsBusy() == true)
                    Thread.Sleep(1);
            }
        }

        /// <summary>
        /// シリアルポートデータ受信メソッド（イベント発生）
        /// </summary>
        private void SerialDataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            //string rcvDataStr = "";

            receiveRawData = new byte[serialPort.BytesToRead];
            serialPort.Read(receiveRawData, 0, serialPort.BytesToRead);

            //bgLibに処理を委譲（受信データに応じたイベントが発生）
            for (int i = 0; i < receiveRawData.Length; i++)
                bgLib.Parse(receiveRawData[i]);
        }

    }
}
