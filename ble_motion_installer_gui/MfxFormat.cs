using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PLEN.MFX
{
    /// <summary>
    /// モーションファイル送信データ化メソッド
    /// </summary>
    public class BLEMfxCommand : PLEN.BLECommand
    {
        /// <summary>
        /// コマンド変換元Mfxデータ（XML形式）
        /// </summary>
        public XmlMfxModel xmlMfx;
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="mfxModel">コマンド変換元Mfxデータ（XML形式）</param>
        public BLEMfxCommand(XmlMfxModel mfxModel)
        {
            xmlMfx = mfxModel;
        }
        /// <summary>
        /// モーションデータ名
        /// </summary>
        public override string Name
        {
            get
            {
                // 先頭のモーションのNameをこのクラスのNameとして返す（特に深い意味はない）
                if (xmlMfx.Motion.Count > 0)
                    return xmlMfx.Motion[0].Name;
                else
                    return "";
            }
        }
        public override short Slot
        {
            get
            {
                return slot;
            }
        }

        private short slot;
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
                foreach (TagMotionModel tagMotion in xmlMfx.Motion)
                {
                    /*----- システムコマンド「slotNum」 -----*/
                    slot = short.Parse(tagMotion.ID);
                    convertedStr += slot.ToString("x2");
                    convertedStrForDisplay += "[slotNum : " + slot.ToString("x2") + "] ";

                    /*----- システムコマンド「name」 -----*/
                    convertedStr += string.Format("{0,-20}", tagMotion.Name);
                    convertedStrForDisplay += "[name : " + string.Format("{0,-20}", tagMotion.Name) + "] ";  
     
                    /*----- システムコマンド「config」 -----*/
                    // Paramはid:0，id:1の2つしかない
                    if (tagMotion.Extra.Param.Count != 2)
                        return false;
                    // ParamリストをIDの昇順に並び替え
                    tagMotion.Extra.Param.Sort((o1, o2) => (Int32.Parse(o1.ID)).CompareTo(Int32.Parse(o2.ID)));
                    // コマンドを追加
                    convertedStr += byte.Parse(tagMotion.Extra.Function).ToString("x2");
                    convertedStr += byte.Parse(tagMotion.Extra.Param[0].Param).ToString("x2");
                    convertedStr += byte.Parse(tagMotion.Extra.Param[1].Param).ToString("x2");
                    convertedStrForDisplay += "[config : " + int.Parse(tagMotion.Extra.Function).ToString("x2");
                    convertedStrForDisplay += byte.Parse(tagMotion.Extra.Param[0].Param).ToString("x2");
                    convertedStrForDisplay += byte.Parse(tagMotion.Extra.Param[1].Param).ToString("x2") + "] ";

                    /*----- システムコマンド「frameNum」 -----*/
                    convertedStr += int.Parse(tagMotion.FrameNum).ToString("x2");
                    convertedStrForDisplay += "[frameNum : " + byte.Parse(tagMotion.FrameNum).ToString("x2") + "] ";

                    /*----- システムコマンド「frame」 -----*/
                    // FrameリストをIDの昇順に並び替え
                    tagMotion.Frame.Sort((o1, o2) => (Int32.Parse(o1.ID)).CompareTo(Int32.Parse(o2.ID)));
                    convertedStrForDisplay += "[frame : ";
                    foreach (TagFrameModel tagFrame in tagMotion.Frame)
                    {
                        convertedStr += short.Parse(tagFrame.Time).ToString("x4");
                        convertedStrForDisplay += " " + short.Parse(tagFrame.Time).ToString("x4");
                        // JointリストをIDの昇順に並べ替え
                        tagFrame.Joint.Sort((o1, o2) => (Int32.Parse(o1.ID)).CompareTo(Int32.Parse(o2.ID)));
                        foreach (TagJointModel tagJoint in tagFrame.Joint)
                        {
                            convertedStr += short.Parse(tagJoint.Joint).ToString("x4");
                            convertedStrForDisplay += short.Parse(tagJoint.Joint).ToString("x4");
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
    /// <summary>
    /// モーションファイル（XML形式）
    /// </summary>
    [System.Xml.Serialization.XmlRoot("mfx")]
    public class XmlMfxModel
    {
        [System.Xml.Serialization.XmlElement("motion")]
        public List<PLEN.MFX.TagMotionModel> Motion { get; set; }
    }
    /// <summary>
    /// モーションファイル：motionタグ
    /// </summary>
    public class TagMotionModel
    {
        [System.Xml.Serialization.XmlAttribute("id")]
        public String ID { get; set; }

        [System.Xml.Serialization.XmlElement("name")]
        public string Name { get; set; }

        [System.Xml.Serialization.XmlElement("extra")]
        public TagExtraModel Extra { get; set; }

        [System.Xml.Serialization.XmlElement("frameNum")]
        public string FrameNum { get; set; }

        [System.Xml.Serialization.XmlElement("frame")]
        public List<TagFrameModel> Frame { get; set; }
    }
    /// <summary>
    /// モーションファイル：extraタグ
    /// </summary>
    public class TagExtraModel
    {

        [System.Xml.Serialization.XmlElement("function")]
        public string Function { get; set; }

        [System.Xml.Serialization.XmlElement("param")]
        public List<TagParamModel> Param { get; set; }
    }
    /// <summary>
    /// モーションファイル：paramタグ
    /// </summary>
    public class TagParamModel
    {
        [System.Xml.Serialization.XmlAttribute("id")]
        public String ID { get; set; }

        [System.Xml.Serialization.XmlText()]
        public string Param { get; set; }
    }
    /// <summary>
    /// モーションファイル：frameタグ
    /// </summary>
    public class TagFrameModel
    {
        [System.Xml.Serialization.XmlAttribute("id")]
        public String ID { get; set; }

        [System.Xml.Serialization.XmlElement("time")]
        public string Time { get; set; }

        [System.Xml.Serialization.XmlElement("joint")]
        public List<TagJointModel> Joint { get; set; }
    }
    /// <summary>
    /// モーションファイル：jointタグ
    /// </summary>
    public class TagJointModel
    {
        [System.Xml.Serialization.XmlAttribute("id")]
        public String ID { get; set; }

        [System.Xml.Serialization.XmlText()]
        public string Joint { get; set; }
    }
}
