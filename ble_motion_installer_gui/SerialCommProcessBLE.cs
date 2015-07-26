using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Threading;

namespace BLEMotionInstaller
{
    class SerialCommProcessBLE : SerialCommProcess
    {
        /// <summary>
        /// シリアル通信メソッドからのメッセージ受信イベント
        /// </summary>
        public override event SerialCommProcessMessageHandler serialCommProcessMessage;
        /// <summary>
        /// シリアル通信メソッド終了イベント
        /// </summary>
        public override event SerialCommProcessFinishedHandler serialCommProcessFinished;
        /// <summary>
        /// BLEドングル接続完了イベント
        /// </summary>
        public override event SerialCommProcessBLEConnectedHander serialCommProcessConnected;
        /// <summary>
        /// モーションデータ送信完了イベント（1モーション送信完了ごとに呼び出される）
        /// </summary>
        public override event SeiralCommProcessMfxCommandSendedHandler serialCommProcessCommandSended;
        /// <summary>
        /// PLEN2キャラスティックUUID
        /// </summary>
        protected readonly byte[] PLEN2_TX_CHARACTERISTIC_UUID =
	        {
		        0xF9, 0x0E, 0x9C, 0xFE, 0x7E, 0x05, 0x44, 0xA5, 0x9D, 0x75, 0xF1, 0x36, 0x44, 0xD6, 0xF6, 0x45
	        };
        private readonly byte[] PLEN_CONTROL_SERVICE_UUID =
          {
                0xE1, 0xF4, 0x04, 0x69, 0xCF, 0xE1, 0x43, 0xC1, 0x83, 0x8D, 0xDD, 0xBC, 0x9D, 0xAF, 0xDD, 0xE6
          };

        private ushort plenTxAtthandle = 0;
        private ushort serviceStartHandle = 0;
        private ushort serviceEndHandle = 0;
        /// <summary>
        /// BLE接続先クライアントキー
        /// </summary>
        private Int64 connectedBLEKey = 0;
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
        /// コンストラクタ
        /// </summary>
        /// <param name="sender">インスタンス作成元Form</param>
        /// <param name="portName">シリアルポート名</param>
        /// <param name="mfxCommandList">送信モーションファイルリスト</param>
        public SerialCommProcessBLE(object sender, string portName, List<PLEN.BLECommand> commandList)
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
            bgLib.BLEEventATTClientGroupFound += new Bluegiga.BLE.Events.ATTClient.GroupFoundEventHandler(bgLib_BLEEventGroupFound);
            bgLib.BLEEventATTClientFindInformationFound += new Bluegiga.BLE.Events.ATTClient.FindInformationFoundEventHandler(bleLib_BLEEventFindInformationFound);
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

        Int64 getBleAddress(byte[] addressArray)
        {
            if (addressArray.Length < 6)
                return 0;

            Int64 key = 0;
            for (int index = 0; index < 6; index++)
                key += (Int64)addressArray[index] << (index * 8);

            return key;
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
            serialCommProcessMessage(this, "PLEN searching...");
            connectState = SerialState.NotConnected;
            bgLib.SendCommand(serialPort, bgLib.BLECommandGAPDiscover(1));
            while (connectState != SerialState.Connected)
                Thread.Sleep(1);
        }

        /// <summary>
        /// 通信開始（ スレッド用メソッド．直接呼びださない．）
        /// </summary>
        public override void start()
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
            Int64 key = getBleAddress(e.sender);

            // 取得したキーに対してまだ接続していない場合，接続を試みる
            // ※connectedDictは他スレッドからの参照もありうるので排他制御
            if (connectState == SerialState.NotConnected)
            {
                // PLEN BLEのアドレスはパブリックアドレスである
                if (e.address_type != 0)
                {
                    return;
                }
                lock(connectedDict)
                {
                    // 取得したキーに対して，すでに接続を試みたならばスルー
                    if (connectedDict.ContainsKey(key) == true && connectedDict[key] != SerialState.NotConnected)
                    {
                        return;
                    }
 
                    // 取得したキーをリストに追加，状態を接続要求中へ
                    if (connectedDict.ContainsKey(key) == true)
                        connectedDict.Add(key, SerialState.Connecting);
                    else
                        connectedDict[key] = SerialState.Connecting;

                }
                // 自分の接続状態を更新
                connectState = SerialState.Connecting;
                plenTxAtthandle = 0;
                serviceStartHandle = 0;
                serviceEndHandle = 0;
                // 取得したキーに対して接続を試みる(接続完了後イベント発生)
                //bgLib.BLECommandGAPEndProcedure();
                bgLib.SendCommand(serialPort, bgLib.BLECommandGAPConnectDirect(e.sender, e.address_type, 60, 76, 100, 0));
            }
        }
        /// <summary>
        /// BLEクライアント接続完了メソッド（bgLibイベント呼び出し）
        /// </summary>
        private void bgLib_BLEEventConnectionStatus(object sender, Bluegiga.BLE.Events.Connection.StatusEventArgs e)
        {
            // アドレス作成
            Int64 key = getBleAddress(e.address);

            // ペリフェラルと接続完了
            if ((e.flags & 0x01) != 0)
            {
                // アドレスリストを更新（接続要求中→サービス検知へ）
                // ※connectedDictは他スレッドからも操作されるので排他制御
                lock (connectedDict)
                {
                    serialCommProcessMessage(this, "[" + key.ToString() + "] connected. Scan Services...");

                    lock(connectedDict)
                        connectedDict[key] = SerialState.ScanServices;
                    connectState = SerialState.ScanServices;
                    connectedBLEKey = key;
                    bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientReadByGroupType(e.connection, 0x0001, 0xFFFF, new byte[] { 0x00, 0x28 }));
                }

            }
            // 再度接続を試みる
            else
            {
                lock (connectedDict)
                    connectedDict[key] = SerialState.NotConnected;
                connectState = SerialState.NotConnected;
                bgLib.SendCommand(serialPort, bgLib.BLECommandGAPDiscover(1));
            }
        }
        /// <summary>
        /// サービス検知メソッド（bgLibイベント呼び出し）
        /// </summary>
        void bgLib_BLEEventGroupFound(object sender, Bluegiga.BLE.Events.ATTClient.GroupFoundEventArgs e)
        {
            // e.uuidはリトルエンディアン
            if (e.uuid.SequenceEqual(PLEN_CONTROL_SERVICE_UUID.Reverse()))
            {
                serviceStartHandle = e.start;
                serviceEndHandle = e.end;
            }
        }
        /// <summary>
        /// キャラクタリスティック検知メソッド（bgLibイベント呼び出し）
        /// </summary>
        void bleLib_BLEEventFindInformationFound(object sender, Bluegiga.BLE.Events.ATTClient.FindInformationFoundEventArgs e)
        {
            // e.uuidはリトルエンディアン
            if(e.uuid.SequenceEqual(PLEN2_TX_CHARACTERISTIC_UUID.Reverse()))
            {
                plenTxAtthandle = e.chrhandle;
            }
        }

        /// <summary>
        /// ATTClientProcedureCompletedメソッド（bgLibイベント呼び出し）
        /// </summary>
        void bgLib_BLEEventATTClientProcedureCompleted(object sender, Bluegiga.BLE.Events.ATTClient.ProcedureCompletedEventArgs e)
        { 
            switch (connectState)
            {
                // Attribute Write完了
                case SerialState.Connected:
                     if (e.result == 0)
                        isAttributeWrited = true;
                    break;
                // Serviceスキャン完了
                case SerialState.ScanServices:
                    // PLEN Control Serviceあり
                    if (serviceEndHandle > 0)
                    {
                        lock(connectedDict)
                            connectedDict[connectedBLEKey] = SerialState.ScanCharacteristics;

                        serialCommProcessMessage(this, "PLEN Control Service Detected. Scan Characteristics...");
                        connectState = SerialState.ScanCharacteristics;
                        bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientFindInformation(e.connection, serviceStartHandle, serviceEndHandle));
                    }
                    // PLEN Control Serivceなし→PLENでない（接続リストから除外し再探索）
                    else
                    {
                        lock(connectedDict)
                            connectedDict[connectedBLEKey] = SerialState.NotPLEN2;

                        serialCommProcessMessage(this, "[" + connectedBLEKey.ToString() + "] is not PLEN.");
                        bgLib.SendCommand(serialPort, bgLib.BLECommandConnectionDisconnect(0));
                        Thread.Sleep(250);

                        connectState = SerialState.NotConnected;
                        bgLib.SendCommand(serialPort, bgLib.BLECommandGAPDiscover(1));
                        serialCommProcessMessage(this, "PLEN re-searching...");
                    }

                    break;
                // Characteristicスキャン完了
                case SerialState.ScanCharacteristics:
                    // PLEN TX Characteristicあり→接続しているペリフェラルがPLENであることが確定
                    if (plenTxAtthandle > 0)
                    {
                        lock (connectedDict)
                            connectedDict[connectedBLEKey] = SerialState.Connected;

                        connectState = SerialState.Connected;
                        serialCommProcessMessage(this, "[" + connectedBLEKey.ToString() + "] is PLEN !!");
                    }
                    // PLEN TX Characteristicなし→PLENでない（接続リストから除外し再探索）
                    else
                    {
                        lock (connectedDict)
                            connectedDict[connectedBLEKey] = SerialState.NotPLEN2;

                        serialCommProcessMessage(this, "[" + connectedBLEKey.ToString() + "] is not PLEN.");
                        bgLib.SendCommand(serialPort, bgLib.BLECommandConnectionDisconnect(0));
                        Thread.Sleep(250);

                        connectState = SerialState.NotConnected;
                        bgLib.SendCommand(serialPort, bgLib.BLECommandGAPDiscover(1));
                        serialCommProcessMessage(this, "PLEN re-searching...");
                    }
                    break;
            }
        }

        /// <summary>
        /// 半二重通信メソッド
        /// </summary>
        private void halfDuplexComm()
        {

            serialCommProcessMessage(this, "HalfDuplexCommunication Started");
            sendedCommandCnt = 0;    // カウントリセット

            connectState = SerialState.NotConnected;
            // PLEN2との接続処理をPLEN2接続用スレッド上（シングルタスク）上で行うため，接続要求リストにセット
            // ※PLEN2接続スレッドはbleConnectingRequestPortListにアイテムが追加されると自動的にテーブルの1番目から接続処理を行う
            if(!formSender.bleConnectingRequestPortList.Contains(PortName))
                formSender.bleConnectingRequestPortList.Add(PortName);
            serialCommProcessMessage(this, "BLE Connecting Thread Waiting...");
            while (connectState != SerialState.Connected)
                Thread.Sleep(1);

            /*-- ここからPLEN2と接続中 --*/
            serialCommProcessMessage(this, "PLEN Connected");
            serialCommProcessConnected(this);

            Thread.Sleep(100);

            try
            {
                foreach (PLEN.BLECommand sendCommand in sendCommandList)
                {
                    // 送信データを文字列からbyte配列に変換
                    byte[] commandArray = System.Text.Encoding.ASCII.GetBytes(sendCommand.convertedStr);
                    serialCommProcessMessage(this, "【" + sendCommand.Name + "】 is sending...");

                    /*---- ここからモーションデータ送信 -----*/
                    /*-- header --*/
                    isAttributeWrited = false;
                    bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, plenTxAtthandle, System.Text.Encoding.ASCII.GetBytes("#IN")));
                    while (isAttributeWrited == false)
                        Thread.Sleep(1);

                    Thread.Sleep(DELAY_INTERVAL);

                    isAttributeWrited = false;
                    bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, plenTxAtthandle, commandArray.Take(20).ToArray()));
                    while (isAttributeWrited == false)
                        Thread.Sleep(1);

                    Thread.Sleep(DELAY_INTERVAL);

                    isAttributeWrited = false;
                    bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, plenTxAtthandle, commandArray.Skip(20).Take(10).ToArray()));
                    while (isAttributeWrited == false)
                        Thread.Sleep(1);

                    Thread.Sleep(DELAY_INTERVAL);
                    serialCommProcessMessage(this, "header written.");

                    Thread.Sleep(50);
                    /*-- frame --*/
                    for (int index = 0; index < (commandArray.Length - 30) / 100; index++)
                    {
                        isAttributeWrited = false;
                        bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, plenTxAtthandle, commandArray.Skip(30 + index * 100).Take(20).ToArray()));
                        while (isAttributeWrited == false)
                            Thread.Sleep(1);

                        Thread.Sleep(DELAY_INTERVAL);

                        isAttributeWrited = false;
                        bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, plenTxAtthandle, commandArray.Skip(50 + index * 100).Take(20).ToArray()));
                        while (isAttributeWrited == false)
                            Thread.Sleep(1);

                        Thread.Sleep(DELAY_INTERVAL);

                        isAttributeWrited = false;
                        bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, plenTxAtthandle, commandArray.Skip(70 + index * 100).Take(20).ToArray()));
                        while (isAttributeWrited == false)
                            Thread.Sleep(1);
                        Thread.Sleep(DELAY_INTERVAL);
                        isAttributeWrited = false;
                        bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, plenTxAtthandle, commandArray.Skip(90 + index * 100).Take(20).ToArray()));
                        while (isAttributeWrited == false)
                            Thread.Sleep(1);

                        Thread.Sleep(DELAY_INTERVAL);

                        isAttributeWrited = false;
                        bgLib.SendCommand(serialPort, bgLib.BLECommandATTClientAttributeWrite(0, plenTxAtthandle, commandArray.Skip(110 + index * 100).Take(20).ToArray()));
                        while (isAttributeWrited == false)
                            Thread.Sleep(1);

                        Thread.Sleep(DELAY_INTERVAL);

                        serialCommProcessMessage(this, "frame written. [" + (index + 1).ToString() + "/" + ((commandArray.Length - 30) / 100).ToString() + "]");
                        Thread.Sleep(50);
                    }
                    serialCommProcessMessage(this, "【" + sendCommand.Name + "】send Complete. ...");
                    sendedCommandCnt++;
                    serialCommProcessCommandSended(this);
                    Thread.Sleep(500);
                }
                // 終了イベントを発生
                serialCommProcessMessage(this, "Communication Finished");
                serialCommProcessFinished(this, new SerialCommProcessFinishedEventArgs(serialPort.PortName, sendCommandList));
                connectedDict[connectedBLEKey] = SerialState.SendCompleted;
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
            int size = serialPort.BytesToRead;

            receiveRawData = new byte[size];
            serialPort.Read(receiveRawData, 0, size);

            //bgLibに処理を委譲（受信データに応じたイベントが発生）
            for (int i = 0; i < receiveRawData.Length; i++)
                bgLib.Parse(receiveRawData[i]);
        }
    }
}
