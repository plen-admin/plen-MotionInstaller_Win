using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Threading;

namespace BLEMotionInstaller
{
    class SerialCommProcessUSB : SerialCommProcess
    {
        /// <summary>
        /// シリアル通信メソッドからの受信イベント
        /// </summary>
        public override event SerialCommProcessMessageHandler serialCommProcessMessage;
        /// <summary>
        /// シリアル通信メソッド終了イベント
        /// </summary>
        public override event SerialCommProcessFinishedHandler serialCommProcessFinished;
        /// <summary>
        /// BLEドングル接続完了イベント（本クラスでは使用しない）
        /// </summary>
        public override event SerialCommProcessBLEConnectedHander serialCommProcessConnected;
        /// <summary>
        /// モーションデータ送信完了イベント（1モーション送信完了ごとに呼び出される）
        /// </summary>
        public override event SeiralCommProcessMfxCommandSendedHandler serialCommProcessCommandSended;

        private string readMessage;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="sender">インスタンス作成元Form</param>
        /// <param name="portName">シリアルポート名</param>
        /// <param name="mfxCommandList">送信モーションファイルリスト</param>
        public SerialCommProcessUSB(object sender, string portName, List<PLEN.BLECommand> commandList)
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
        /// 通信開始（ スレッド用メソッド．直接呼びださない．）
        /// </summary>
        public override void start()
        {
            // ※親スレッドが本スレッドを破棄しようとするとThreadAbortExceptionが発生
            // 　自身が破棄される前に必ずシリアルポートを閉じる
            try
            {
                readMessage = "";
                halfDuplexComm();
            }
            catch (Exception) { }
            finally
            {
                serialPort.Close();
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

            /*-- ここからPLEN2と接続中 --*/
            serialCommProcessMessage(this, "PLEN2 Connected");
            connectState = SerialState.Connected;
            serialCommProcessConnected(this);
            try
            {
                foreach (PLEN.BLECommand sendCommand in sendCommandList)
                {
                    // 送信データを文字列からbyte配列に変換
                    byte[] commandArray = System.Text.Encoding.ASCII.GetBytes(sendCommand.convertedStr);
                    serialCommProcessMessage(this, "【" + sendCommand.Name + "】 is sending...");

                    /*---- ここからモーションデータ送信 -----*/
                    /*-- header --*/
                    serialPort.Write("#IN");
                    serialPort.Write(commandArray, 0, 30);

                    Thread.Sleep(DELAY_INTERVAL);
                    serialCommProcessMessage(this, "header written.");

                    Thread.Sleep(50);
                    /*-- frame --*/
                    for (int index = 0; index < (commandArray.Length - 30) / 100; index++)
                    {
                        serialPort.Write(commandArray, 30 + index * 100, 100);

                        Thread.Sleep(DELAY_INTERVAL);

                        serialCommProcessMessage(this, "frame written. [" + (index + 1).ToString() + "/" + ((commandArray.Length - 30) / 100).ToString() + "]");
                        Thread.Sleep(20);
                    }
                    serialCommProcessMessage(this, "【" + sendCommand.Name + "】send Complete. ...");
                    sendedCommandCnt++;
                    serialCommProcessCommandSended(this);
                    Thread.Sleep(200);
                    // Debug
                    /*serialCommProcessMessage(this, "[ReadMessage] : " + Environment.NewLine + readMessage);
                    readMessage = "";
                    serialPort.Write("#DM" + sendCommand.Slot.ToString("x2"));
                    Thread.Sleep(100);
                    serialCommProcessMessage(this, "[Dump] : " + Environment.NewLine + readMessage); */
                }
                // 終了イベントを発生
                serialCommProcessMessage(this, "Communication Finished");
                Thread.Sleep(100);
                // Debug
                //serialCommProcessMessage(this, "[ReadMessage] : " + Environment.NewLine + readMessage);
                serialCommProcessFinished(this, new SerialCommProcessFinishedEventArgs(serialPort.PortName, sendCommandList));
                
                connectState = SerialState.SendCompleted;
            }
            catch (Exception e)
            {
                serialCommProcessMessage(this, e.Message);
            }
            finally
            {
            }
        }


        private void SerialDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            byte[] readBuff = new byte[serialPort.BytesToRead];
            serialPort.Read(readBuff, 0, readBuff.Length);
            readMessage += Encoding.ASCII.GetString(readBuff);
        }
    }
}
