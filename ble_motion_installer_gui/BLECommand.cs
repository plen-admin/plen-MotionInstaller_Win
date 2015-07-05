using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PLEN
{
    public abstract class BLECommand
    {
        /// <summary>
        /// 変換完了モーションデータ（文字列形式）
        /// </summary>
        public string convertedStr;
        /// <summary>
        /// 変換完了モーションデータ（文字列形式．画面表示用．）
        /// </summary>
        public string convertedStrForDisplay;
        /// <summary>
        /// モーションデータ変換完了フラグ
        /// </summary>
        public bool isConverted = false;

        /// <summary>
        /// モーションデータ名
        /// </summary>
        public abstract string Name
        {
            get;
        }

        /// <summary>
        /// モーションデータ変換メソッド
        /// </summary>
        /// <returns>false：失敗，true：成功</returns>
        public abstract bool convertCommand();
    }
}
