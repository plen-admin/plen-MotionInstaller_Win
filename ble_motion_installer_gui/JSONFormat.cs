using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PLEN.JSON;

namespace PLEN.JSON
{
    class BLEJsonCommand : PLEN.BLECommand
    {
	       /// <summary>
        /// コマンド変換元Mfxデータ（XML形式）
        /// </summary>
        public Main jsonData;
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="mfxModel">コマンド変換元Mfxデータ（XML形式）</param>
        public BLEJsonCommand(Main readJsonData)
        {
            jsonData = readJsonData;
        }
        /// <summary>
        /// モーションデータ名
        /// </summary>
        public override string Name
        {
            get
            {
                return jsonData.name;
            }
        }
        /// <summary>
        /// スロット番号
        /// </summary>
        public override short Slot
        {
            get
            {
                return jsonData.slot;
            }
        }
        private readonly int[] JOINT_MAP = {0, 1, 2, 3, 4, 5, 6, 7, 8, -1, -1, -1, 9, 10, 11, 12, 13, 14, 15, 16, 17, -1, -1, -1};

        /// <summary>
        /// モーションデータ変換メソッド
        /// </summary>
        /// <returns>false：失敗，true：成功</returns>
        public override bool convertCommand()
        {
            convertedStr = "";
            convertedStrForDisplay = "";
            isConverted = false;
            try
            {
                /*----- システムコマンド「slotNum」 -----*/
                convertedStr += jsonData.slot.ToString("x2");
                convertedStrForDisplay += "[slotNum : " + convertedStr + "] ";

                /*----- システムコマンド「name」 -----*/
                convertedStr += string.Format("{0,-20}", jsonData.name);
                convertedStrForDisplay += "[name : " + string.Format("{0,-20}", jsonData.name) + "] ";  
     
                /*----- システムコマンド「config」（フォーマットが決まり次第実装） -----*/
                // Paramはid:0，id:1の2つしかない
                convertedStr += "000000";
                convertedStrForDisplay += "[config : " + "000000" + "] ";

                /*----- システムコマンド「frameNum」 -----*/
                convertedStr += jsonData.frames.Count.ToString("x2");
                convertedStrForDisplay += "[frameNum : " + jsonData.frames.Count.ToString("x2") + "] ";

                /*----- システムコマンド「frame」 -----*/
   
                foreach (Frame frame in jsonData.frames)
                {
                    convertedStrForDisplay += "[frame : ";
                    string[] jointStrs = Enum.GetNames(typeof(JointName));

                    convertedStr += frame.transition_time_ms.ToString("x4");
                    convertedStrForDisplay += " " + frame.transition_time_ms.ToString("x4") + " ";
                    // JSONファイル中の関節名（output.device）を送信用のインデックスに変換
                    // Note...インデックスは一部連続でないためJOINT_MAPによってマッピングを行う.
                    //        あまりのインデックスは0埋め
                    for (int i = 0; i < JOINT_MAP.Length; i++)
                    {
                        // モーションに関係ないインデックス（0埋め）
                        if (JOINT_MAP[i] < 0)
                        {
                            convertedStr += "0000";
                            convertedStrForDisplay += "0000 ";
                        }
                        else
                        {
                            foreach (Output output in frame.outputs)
                            {
                                // インデックスを0から順にデータ格納する必要がある．
                                // したがって送信インデックスに応じたデータをJSONファイルから探索し，データ格納を行う．
                                if (output.device == jointStrs[JOINT_MAP[i]])
                                {   
                                    convertedStr += output.value.ToString("x4");
                                    convertedStrForDisplay += output.value.ToString("x4") + " ";
                                    break;
                                }


                            }
                        }
                    }

                    convertedStrForDisplay += "]";
                }
            }
            catch (Exception)
            {
                return false;
            }

            isConverted = true;
            return true;
        }

    }
}
	